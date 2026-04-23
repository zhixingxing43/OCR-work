"""
基于 PaddleOCR + UIE-X 的发票关键信息抽取工具
功能：批量识别发票图片/PDF，提取发票头关键字段，输出 Excel
作者：财务智能化工具
版本：2.0 (UIE-X 高精度版)
"""

import os
import re
import json
import pandas as pd
from PIL import Image
from paddleocr import PaddleOCR
from paddlenlp import Taskflow

# ================== 配置区域（请根据实际情况修改） ==================
IMAGE_FOLDER = "./invoices"          # 存放发票图片/PDF的文件夹路径
OUTPUT_EXCEL = "./发票识别结果_UIEX.xlsx"  # 输出 Excel 文件路径
GPU_ENABLED = False                  # 是否使用 GPU 加速（需安装 paddlepaddle-gpu）
OCR_USE_ANGLE_CLS = True             # OCR 方向校正
OCR_LANG = 'ch'                      # 识别语言
# ===================================================================

# 支持的图片格式
SUPPORTED_EXT = ('.jpg', '.jpeg', '.png', '.bmp', '.tiff', '.pdf')

# UIE-X 信息抽取的 Schema 定义（发票关键字段）
# 使用 UIE-X 的跨模态文档抽取能力，通过自然语言描述字段即可
SCHEMA = [
    '发票代码',
    '发票号码',
    '开票日期',
    '校验码',
    '价税合计',
    '购买方名称',
    '购买方纳税人识别号',
    '销售方名称',
    '销售方纳税人识别号'
]

# 字段名称映射（用于统一输出列名）
FIELD_MAP = {
    '发票代码': '发票代码',
    '发票号码': '发票号码',
    '开票日期': '开票日期',
    '校验码': '校验码',
    '价税合计': '价税合计(小写)',
    '购买方名称': '购买方名称',
    '购买方纳税人识别号': '购买方税号',
    '销售方名称': '销售方名称',
    '销售方纳税人识别号': '销售方税号'
}

# 可选：补充大写金额字段（UIE-X 不一定直接输出，通过 OCR 文本正则补充）
EXTRA_FIELDS = ['价税合计(大写)']


def init_ocr():
    """
    初始化 PaddleOCR 引擎（用于获取纯文本行，辅助 UIE-X 抽取不足的字段）
    """
    return PaddleOCR(
        use_angle_cls=OCR_USE_ANGLE_CLS,
        lang=OCR_LANG,
        show_log=False,
        use_gpu=GPU_ENABLED
    )


def init_uie_x():
    """
    初始化 UIE-X 文档信息抽取模型（基于多模态跨模态架构）
    首次运行会自动下载约 2GB 模型文件
    """
    print("正在加载 UIE-X 模型（首次运行需下载约 2GB 模型，请耐心等待）...")
    ie = Taskflow(
        'information_extraction',
        schema=SCHEMA,
        model='uie-x-base',
        task_path=None,               # 使用官方预训练模型
        device_id=0 if GPU_ENABLED else -1  # -1 表示 CPU
    )
    print("UIE-X 模型加载完成！")
    return ie


def extract_with_uiex(ie, image_path):
    """
    使用 UIE-X 直接对文档图像进行关键信息抽取
    返回字典，键为字段名，值为识别结果（列表形式）
    """
    try:
        # UIE-X 支持直接传入图片路径
        result = ie({'doc': image_path})
        return result
    except Exception as e:
        print(f"    UIE-X 抽取出错: {e}")
        return {}


def extract_full_text(ocr_engine, image_path):
    """
    使用 PaddleOCR 获取全部文本行（用于正则提取大写金额等补充字段）
    返回拼接后的全文和行列表
    """
    try:
        result = ocr_engine.ocr(image_path, cls=OCR_USE_ANGLE_CLS)
        if not result or not result[0]:
            return '', []
        lines = [line[1][0] for line in result[0]]
        full_text = ''.join(lines)
        return full_text, lines
    except Exception as e:
        print(f"    OCR 文本提取出错: {e}")
        return '', []


def parse_uiex_result(uiex_result, full_text=''):
    """
    将 UIE-X 返回的嵌套结构转换为扁平的字段字典
    UIE-X 返回格式示例：
    {
      '发票代码': [{'text': '123456789012', 'probability': 0.98, ...}],
      ...
    }
    """
    fields = {}
    for key in SCHEMA:
        value_list = uiex_result.get(key, [])
        if value_list:
            # 取置信度最高的一项
            best_item = max(value_list, key=lambda x: x.get('probability', 0))
            fields[FIELD_MAP.get(key, key)] = best_item.get('text', '').strip()
        else:
            fields[FIELD_MAP.get(key, key)] = ''

    # 补充大写金额（UIE-X 可能未定义，从 OCR 全文正则提取）
    capital_amount = extract_capital_amount(full_text)
    fields['价税合计(大写)'] = capital_amount

    # 补充校验码（部分发票版式 UIE-X 可能漏掉，用正则兜底）
    if not fields.get('校验码'):
        fields['校验码'] = extract_check_code(full_text)

    return fields


def extract_capital_amount(full_text):
    """
    从文本中提取中文大写金额
    """
    patterns = [
        r'[（(]大写[）)][：:\s]*([壹贰叁肆伍陆柒捌玖拾佰仟万亿元角分整零]+)',
        r'价税合计.*?大写.*?[：:\s]*([壹贰叁肆伍陆柒捌玖拾佰仟万亿元角分整零]+)',
        r'人民币.*?大写.*?[：:\s]*([壹贰叁肆伍陆柒捌玖拾佰仟万亿元角分整零]+)',
        r'金额大写[：:\s]*([壹贰叁肆伍陆柒捌玖拾佰仟万亿元角分整零]+)'
    ]
    for pat in patterns:
        match = re.search(pat, full_text)
        if match:
            return match.group(1).strip()
    return ''


def extract_check_code(full_text):
    """
    提取校验码（通常 6 位数字或字母数字组合）
    """
    match = re.search(r'(?:校验码|效验码)[：:\s]*([A-Za-z0-9]{6,8})', full_text)
    if match:
        return match.group(1)
    # 尝试单独匹配 6 位数字（位于发票右上角附近，但 OCR 文本中可能独立出现）
    # 为避免误匹配，只在前述正则失败时采用较宽松策略
    match = re.search(r'(\d{6})\s*$', full_text.strip())
    if match:
        return match.group(1)
    return ''


def process_images():
    """
    主处理函数：遍历文件夹 -> UIE-X 抽取 -> 补充正则 -> 保存 Excel
    """
    if not os.path.exists(IMAGE_FOLDER):
        print(f"错误：图片文件夹不存在 -> {IMAGE_FOLDER}")
        return

    # 获取所有支持的文件
    all_files = os.listdir(IMAGE_FOLDER)
    image_files = [f for f in all_files if f.lower().endswith(SUPPORTED_EXT)]
    if not image_files:
        print(f"警告：在 {IMAGE_FOLDER} 中没有找到支持的图片/PDF文件")
        return

    print(f"共发现 {len(image_files)} 个文件，开始处理...")
    print("=" * 60)

    # 初始化引擎（只需一次）
    ie = init_uie_x()
    ocr = init_ocr()  # 仅用于补充字段，也可不加，若不需要补充可注释

    records = []
    total = len(image_files)

    for idx, filename in enumerate(image_files, 1):
        filepath = os.path.join(IMAGE_FOLDER, filename)
        print(f"[{idx}/{total}] 正在处理: {filename}")

        try:
            # 1. UIE-X 抽取主要字段
            uiex_result = extract_with_uiex(ie, filepath)

            # 2. 获取 OCR 全文（用于补充正则）
            full_text, _ = extract_full_text(ocr, filepath)

            # 3. 解析并合并字段
            fields = parse_uiex_result(uiex_result, full_text)
            fields['文件名'] = filename

            records.append(fields)

            # 打印摘要
            print(f"  -> 发票号码: {fields.get('发票号码', '')}, "
                  f"价税合计: {fields.get('价税合计(小写)', '')}, "
                  f"销售方: {fields.get('销售方名称', '')[:20]}")

        except Exception as e:
            print(f"  -> 处理失败: {e}")
            # 添加一条空记录，便于定位问题文件
            records.append({
                '文件名': filename,
                '发票代码': '', '发票号码': '', '开票日期': '', '校验码': '',
                '价税合计(小写)': '', '价税合计(大写)': '',
                '购买方名称': '', '购买方税号': '', '销售方名称': '', '销售方税号': '',
                '备注': f'处理异常: {str(e)}'
            })

    # 保存 Excel
    if records:
        df = pd.DataFrame(records)
        # 定义列顺序
        columns_order = [
            '文件名', '发票代码', '发票号码', '开票日期', '校验码',
            '价税合计(小写)', '价税合计(大写)',
            '购买方名称', '购买方税号',
            '销售方名称', '销售方税号'
        ]
        # 可能存在额外列（如备注），一并保留
        existing_cols = [c for c in columns_order if c in df.columns]
        other_cols = [c for c in df.columns if c not in columns_order]
        df = df[existing_cols + other_cols]

        df.to_excel(OUTPUT_EXCEL, index=False, engine='openpyxl')
        print("=" * 60)
        print(f"✅ 处理完成！共处理 {len(records)} 个文件，结果保存至: {OUTPUT_EXCEL}")
    else:
        print("未生成任何有效记录。")


if __name__ == "__main__":
    process_images()
