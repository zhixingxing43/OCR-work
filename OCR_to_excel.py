"""
============================================================
发票 & 合同 OCR 识别工具
环境：PaddleOCR 2.9.1 + PyMuPDF + OpenCV + pandas + openpyxl
功能：
    - 批量处理指定文件夹中的所有图片/PDF
    - 提取发票关键字段（号码、代码、日期、金额等）
    - 合同等非发票文件保存全文识别文本
    - 导出 Excel 表格
使用方法：
    1. 安装依赖：pip install -r requirements.txt
    2. 修改 INPUT_DIR 为你的图片目录
    3. 运行：python invoice_ocr.py
============================================================
"""
import os
os.environ['FLAGS_use_mkldnn'] = '0'   # 禁用 oneDNN，强制使用原生 CPU 推理
os.environ['KMP_DUPLICATE_LIB_OK'] = 'TRUE'  # 避免 OpenMP 冲突（Windows 常见）

import os
import re
import cv2
import numpy as np
import pandas as pd
from paddleocr import PaddleOCR
from PIL import Image
from collections import defaultdict

# ==================== 配置 ====================
INPUT_DIR = r"D:\A\Agent-renting\demo"          # 发票/合同图片所在目录
OUTPUT_EXCEL = "output/invoices.xlsx"           # 输出 Excel 路径
SUPPORTED_IMG_EXTS = (".jpg", ".jpeg", ".png", ".bmp", ".tiff")
SUPPORTED_PDF_EXTS = (".pdf",)

# 发票关键字段关键词映射（可自定义扩展）
INVOICE_KEYWORDS = {
    "发票号码": ["发票号码", "No.", "号码"],
    "发票代码": ["发票代码", "Code"],
    "开票日期": ["开票日期", "日期"],
    "校验码": ["校验码", "验证码"],
    "购买方名称": ["购买方名称", "购货方", "名称（购买方）"],
    "购买方识别号": ["购买方纳税人识别号", "纳税人识别号（购）"],
    "销售方名称": ["销售方名称", "销货方", "名称（销售方）"],
    "销售方识别号": ["销售方纳税人识别号", "纳税人识别号（销）"],
    "合计金额": ["合计金额", "金额（小写）", "小写金额"],
    "合计税额": ["合计税额", "税额（小写）"],
    "价税合计": ["价税合计", "价税合计（大写）", "合计（大写）"],
    "备注": ["备注"],
}

# ==================== 工具函数 ====================
def enhance_image(img: np.ndarray) -> np.ndarray:
    """
    CLAHE 对比度增强（保留颜色）
    """
    lab = cv2.cvtColor(img, cv2.COLOR_BGR2LAB)
    l, a, b = cv2.split(lab)
    clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8))
    l = clahe.apply(l)
    lab = cv2.merge((l, a, b))
    return cv2.cvtColor(lab, cv2.COLOR_LAB2BGR)




def get_bbox_xyxy(box: list) -> tuple:
    """
    将四点坐标转换为 (xmin, ymin, xmax, ymax)
    box: [[x1,y1], [x2,y2], [x3,y3], [x4,y4]]
    """
    xs = [p[0] for p in box]
    ys = [p[1] for p in box]
    return int(min(xs)), int(min(ys)), int(max(xs)), int(max(ys))


def compute_center(box: tuple) -> tuple:
    xmin, ymin, xmax, ymax = box
    return (xmin + xmax) / 2, (ymin + ymax) / 2


def group_lines(ocr_results: list, y_tol: int = 10) -> list:
    """
    将 OCR 结果按行分组（y 坐标接近的视为同一行）
    返回二维列表，每个子列表为同一行的 OCR 条目，按 x 坐标排序
    """
    if not ocr_results:
        return []
    # 按 y 中心排序
    sorted_results = sorted(ocr_results, key=lambda x: (compute_center(get_bbox_xyxy(x[0]))[1], compute_center(get_bbox_xyxy(x[0]))[0]))
    lines = []
    current_line = [sorted_results[0]]
    _, y_center_prev = compute_center(get_bbox_xyxy(sorted_results[0][0]))
    for item in sorted_results[1:]:
        _, y_center = compute_center(get_bbox_xyxy(item[0]))
        if abs(y_center - y_center_prev) <= y_tol:
            current_line.append(item)
        else:
            lines.append(sorted(current_line, key=lambda x: compute_center(get_bbox_xyxy(x[0]))[0]))
            current_line = [item]
        y_center_prev = y_center
    lines.append(sorted(current_line, key=lambda x: compute_center(get_bbox_xyxy(x[0]))[0]))
    return lines


def find_value_by_keyword(keyword: str, lines: list) -> str:
    """
    在按行分组的 OCR 结果中查找 keyword，然后获取其右侧或下一行的值
    """
    # 扁平化所有文本项，以便按坐标查找
    all_items = [(get_bbox_xyxy(item[0]), item[1][0]) for line in lines for item in line]
    for idx, (bbox, text) in enumerate(all_items):
        # 匹配关键词
        if any(kw in text for kw in [keyword]):  # 只匹配传入的具体关键词
            # 1. 尝试同行右侧文本
            for bbox2, text2 in all_items:
                if bbox2[0] > bbox[2] and abs(bbox[1] - bbox2[1]) <= 10:  # 右侧相邻
                    # 取最近的一个
                    dist = bbox2[0] - bbox[2]
                    if dist < 200:  # 避免跨太远
                        return re.sub(r'[：: ]', '', text2).replace(keyword, '').strip()
            # 2. 尝试下一行
            for line in lines:
                line_center_y = np.mean([compute_center(get_bbox_xyxy(item[0]))[1] for item in line])
                if abs(line_center_y - (bbox[1]+bbox[3])/2) <= 15:  # 下一行
                    for item in line:
                        bbox2, text2 = get_bbox_xyxy(item[0]), item[1][0]
                        if bbox2[0] >= bbox[0] - 50:  # 大致同列
                            return re.sub(r'[：: ]', '', text2).strip()
    return ""


def extract_invoice_fields(ocr_results: list) -> dict:
    """
    从 OCR 结果中提取发票关键字段
    """
    lines = group_lines(ocr_results, y_tol=12)
    fields = {}
    # 先扁平化所有文本，供某些需要区分的字段使用（如购买方 vs 销售方）
    all_texts = [item[1][0] for line in lines for item in line]

    for field_en, keywords in INVOICE_KEYWORDS.items():
        value = ""
        for kw in keywords:
            value = find_value_by_keyword(kw, lines)
            if value:
                break
        # 特殊处理：购买方/销售方名称可能以“名称：”单独出现，需要结合上下文
        if field_en == "购买方名称" and not value:
            # 搜索“名称”但排除“销售方”区域
            for line in lines:
                for item in line:
                    bbox, text = get_bbox_xyxy(item[0]), item[1][0]
                    if "名称" in text and "销售" not in text:
                        # 尝试取其右侧或下一行
                        value = find_value_by_keyword("名称", lines)
                        if value:
                            break
        fields[field_en] = value

    # 如果未提取到核心字段，可能不是发票，返回 None
    if not fields["发票号码"] and not fields["发票代码"]:
        return None  # 大概率是合同
    return fields


def recognize_image(image) -> list:
    """
    对 PIL Image 或 numpy array 进行 OCR，返回标准的 ocr 结果列表
    """
    if isinstance(image, Image.Image):
        image_np = cv2.cvtColor(np.array(image), cv2.COLOR_RGB2BGR)
    else:
        image_np = image
    # 图像增强
    enhanced = enhance_image(image_np)
    # 执行 OCR
    result = ocr.ocr(enhanced, cls=True)
    if not result or not result[0]:
        return []
    return result[0]  # 第一页


def process_file(file_path: str) -> list:
    """
    处理单个文件（图片或 PDF），返回每页的 OCR 结果列表
    """
    ext = os.path.splitext(file_path)[1].lower()
    ocr_results_per_page = []

    if ext in SUPPORTED_IMG_EXTS:
        img = Image.open(file_path).convert("RGB")
        ocr_res = recognize_image(img)
        if ocr_res:
            ocr_results_per_page.append(ocr_res)

    else:
        print(f"⚠️ 不支持的文件类型: {file_path}")
    return ocr_results_per_page


# ==================== 主流程 ====================
if __name__ == "__main__":
    # 初始化 PaddleOCR（使用 PP-OCRv4 中文模型，自动下载）
    print("正在加载 OCR 模型...")
    ocr = PaddleOCR(
        use_textline_orientation=True,       # 文字方向分类（新版参数）
        lang="ch",
        text_det_thresh=0.3,        # 检测阈值，调高可减少误检（新版参数）
        text_det_box_thresh=0.5,    # 检测框阈值（新版参数）
        text_recognition_batch_size=6, # 识别批量大小（新版参数）
    )
    print("模型加载完成。")

    # 遍历输入目录
    all_records = []
    if not os.path.exists(INPUT_DIR):
        print(f"❌ 目录不存在: {INPUT_DIR}")
        exit(1)

    files = [f for f in os.listdir(INPUT_DIR) if os.path.isfile(os.path.join(INPUT_DIR, f))]
    if not files:
        print("⚠️ 目录中没有可识别文件。")
        exit(0)

    for filename in files:
        file_path = os.path.join(INPUT_DIR, filename)
        print(f"正在处理: {filename}")
        try:
            pages_ocr = process_file(file_path)
        except Exception as e:
            print(f"❌ 处理失败: {filename}, 错误: {e}")
            continue

        for i, ocr_result in enumerate(pages_ocr):
            # 尝试发票抽取
            fields = extract_invoice_fields(ocr_result)
            record = {
                "文件名": filename,
                "页码": i+1,
                "全文本": " ".join([line[1][0] for line in ocr_result])  # 保存全文
            }
            if fields:  # 是发票
                record.update(fields)
            else:
                # 合同等其他文件只保留全文
                record["备注"] = "非发票文件，请查看全文本"
            all_records.append(record)

    if not all_records:
        print("未提取到任何信息。")
        exit(0)

    # 转换为 DataFrame 并导出 Excel
    df = pd.DataFrame(all_records)
    # 重新排列列顺序：发票字段在前，全文在后
    invoice_columns = [
        "文件名", "页码",
        "发票号码", "发票代码", "开票日期", "校验码",
        "购买方名称", "购买方识别号",
        "销售方名称", "销售方识别号",
        "合计金额", "合计税额", "价税合计",
        "备注", "全文本"
    ]
    # 仅保留实际存在的列
    existing_columns = [col for col in invoice_columns if col in df.columns]
    df = df[existing_columns]

    # 保存 Excel
    os.makedirs("output", exist_ok=True)
    df.to_excel(OUTPUT_EXCEL, index=False, engine="openpyxl")
    print(f"✅ 识别完成！结果已保存至 {OUTPUT_EXCEL}")
