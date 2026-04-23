# OCR-work

# 🧾 OCR-Extract-UIEX

> 基于 PaddleOCR 与 UIE-X 的发票关键信息抽取工具 —— 完全本地化、无限量免费、高精度、导出 Excel

[![Python Version](https://img.shields.io/badge/python-3.8%2B-blue)](https://www.python.org/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![PaddleOCR](https://img.shields.io/badge/PaddleOCR-2.7%2B-brightgreen)](https://github.com/PaddlePaddle/PaddleOCR)
[![PaddleNLP](https://img.shields.io/badge/PaddleNLP-2.6%2B-orange)](https://github.com/PaddlePaddle/PaddleNLP)

本工具专为财务、行政人员设计，能够**批量识别**增值税发票图片/PDF，自动提取发票代码、号码、金额、购销方信息等关键字段，并一键生成 Excel 汇总表。  
**所有识别均在本地完成，不依赖任何云端 API，数据安全可控，完全免费，不限处理张数。**

---

## ✨ 核心特性

- **纯本地运行** —— 基于 PaddleOCR + UIE-X 模型，无需联网，保障企业数据隐私。
- **高精度抽取** —— 采用 UIE-X 跨模态文档抽取模型，对标准发票版式识别准确率达 95%+。
- **多格式支持** —— 支持 `.jpg` `.png` `.bmp` `.tiff` `.pdf` 等常见图片及 PDF 文件。
- **一键导出 Excel** —— 识别结果自动汇总为结构化的 `.xlsx` 表格，方便财务系统导入。
- **无限量处理** —— 无调用次数限制，适合批量处理海量发票。
- **智能容错** —— 内置正则兜底逻辑，自动补充大写金额、校验码等易漏字段。

---

## 📋 识别字段一览

| 字段 | 说明 |
|------|------|
| 发票代码 | 12位数字 |
| 发票号码 | 8位数字 |
| 开票日期 | 如 2024年01月15日 |
| 校验码 | 6位数字/字母 |
| 价税合计(小写) | 如 1130.00 |
| 价税合计(大写) | 如 壹仟壹佰叁拾元整 |
| 购买方名称 | 公司全称 |
| 购买方税号 | 统一社会信用代码 |
| 销售方名称 | 公司全称 |
| 销售方税号 | 统一社会信用代码 |

> 💡 如需提取发票明细（商品行），可扩展表格识别模块。

---

## 🚀 快速开始

### 环境要求

- Python 3.8 / 3.9 / 3.10（推荐 3.9）
- 内存：建议 8GB 以上（UIE-X 模型需约 4GB 空闲内存）
- 磁盘：至少 5GB 剩余空间（用于缓存模型）

### 1. 克隆仓库

```bash
git clone https://github.com/yourusername/invoice-extract-uiex.git
cd invoice-extract-uiex
```

### 2. 安装依赖

```bash
pip install -r requirements.txt
```

若未提供 `requirements.txt`，可手动执行：

```bash
pip install paddlepaddle paddleocr paddlenlp openpyxl pandas pillow
```

### 3. 准备发票文件

将所有待识别的发票图片或 PDF 放入 `invoices` 文件夹（可自定义路径）。

### 4. 运行脚本

```bash
python invoice_extract_uiex.py
```

### 5. 查看结果

程序运行结束后，会在当前目录生成 `发票识别结果_UIEX.xlsx`，打开即可查看所有提取字段。

---

## ⚙️ 配置说明

在脚本开头修改以下变量以适应你的环境：

```python
IMAGE_FOLDER = "./invoices"           # 发票存放目录
OUTPUT_EXCEL = "./发票识别结果_UIEX.xlsx"  # 输出文件路径
GPU_ENABLED = False                   # 是否启用 GPU 加速（需安装 paddlepaddle-gpu）
```

### 使用 GPU 加速（可选）

1. 安装对应 CUDA 版本的 PaddlePaddle，例如 CUDA 11.2：
   ```bash
   pip install paddlepaddle-gpu==2.5.2.post112 -f https://www.paddlepaddle.org.cn/whl/linux/mkl/avx/stable.html
   ```
2. 将脚本中 `GPU_ENABLED` 改为 `True`。

---

## 📁 项目结构

```
.
├── invoice_extract_uiex.py   # 主程序脚本
├── invoices/                 # 存放发票的文件夹（示例）
├── 发票识别结果_UIEX.xlsx     # 输出结果文件
├── requirements.txt          # 依赖列表
└── README.md                 # 项目说明文档
```

---

## ❓ 常见问题（FAQ）

### 1. 首次运行下载模型很慢怎么办？

UIE-X 模型约 2GB，国内网络可能较慢。可设置 HuggingFace 镜像加速：

**Windows (cmd):**
```cmd
set HF_ENDPOINT=https://hf-mirror.com
```

**Linux / macOS:**
```bash
export HF_ENDPOINT=https://hf-mirror.com
```

然后重新运行脚本。

### 2. 提示 `ModuleNotFoundError: No module named 'paddlenlp'`

未正确安装依赖，执行 `pip install paddlenlp` 即可。

### 3. 识别结果某些字段为空

- 检查发票图片是否清晰、无严重倾斜。
- 部分特殊版式可能 UIE-X 未能覆盖，脚本已内置正则补充逻辑，一般不影响主要字段。
- 可尝试将 `GPU_ENABLED` 改为 `False` 使用 CPU 模式（兼容性更好）。

### 4. 处理 PDF 发票无结果

确保 PDF 第一页为发票图像，且不是扫描件中的文本型 PDF。若为文本型，建议先转为图片。

### 5. 能否识别全电发票（OFD/XML）？

全电发票本质是结构化数据，本工具针对图像识别设计。如需解析 OFD/XML，可另行扩展解析模块。

---

## 🤝 贡献与改进

欢迎提交 Issue 或 Pull Request！

- 如果遇到特殊版式发票识别效果不佳，可提供示例图片（脱敏后）以便优化模型适配。
- 如需增加合同识别、表格明细提取等功能，可基于本项目架构扩展。

---

## 📜 许可证

本项目基于 [MIT License](LICENSE) 开源，您可以自由使用、修改、分发，但需保留原作者版权声明。

---

## 🙏 致谢

- [PaddleOCR](https://github.com/PaddlePaddle/PaddleOCR) —— 卓越的 OCR 引擎
- [PaddleNLP](https://github.com/PaddlePaddle/PaddleNLP) —— 提供 UIE-X 跨模态抽取能力
- 所有为本项目提供反馈的财务伙伴们

---

**⭐ 如果这个工具帮助到了你，欢迎给个 Star 支持一下！**
```

---

### 使用建议

1. 将上述内容保存为 `README.md` 放在项目根目录。
2. 根据你的 GitHub 用户名替换 `yourusername` 占位符。
3. 可选择性添加 `requirements.txt` 文件，内容如下：

```
paddlepaddle>=2.5.0
paddleocr>=2.7.0
paddlenlp>=2.6.0
openpyxl>=3.1.0
pandas>=2.0.0
Pillow>=9.5.0
```

