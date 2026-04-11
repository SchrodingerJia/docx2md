# PDF/DOCX 转 Markdown 转换器

这是一个将 Microsoft Word (.docx) 文件或 PDF (.pdf) 文件转换为 Markdown (.md) 格式的 Python 工具。它能够解析文档中的文本、样式、表格、图片、超链接和数学公式，并生成结构化的 Markdown 文档。

## 项目结构

```
docx2md/
├── core/                   # 核心功能模块
│   ├── __init__.py
│   ├── docx_handler.py     # DOCX 文件解析与处理
│   ├── pdf_handler.py      # PDF 文件解析与处理
│   ├── translator.py       # 结构化数据转 Markdown
│   └── main.py             # 主程序入口
├── assets/                 # 资源文件
│   └── images/             # 从源文件提取的图片
├── requirements.txt        # 项目依赖
└── README.md               # 项目说明
```

## 功能特性

- **完整内容解析**：提取文档中的段落、标题、列表、表格、图片、超链接和数学公式。
- **样式保留**：支持加粗、斜体、下划线、字体颜色等文本样式的 Markdown 转换。
- **表格转换**：将 Word 表格转换为 Markdown 表格格式。
- **图片提取**：自动提取文档中的图片并保存到本地目录，在 Markdown 中引用相对路径。
- **数学公式支持**：初步支持行内和块级数学公式的提取与转换。
- **结构化输出**：生成易于阅读和编辑的 Markdown 文档。

## 使用方法

### 1. 安装依赖

确保已安装 Python 3.7+，然后安装所需依赖：

```bash
pip install -r requirements.txt
```

### 2. 运行转换

修改 `core/main.py` 中的文件路径，然后运行：

```bash
python core/main.py
```

### 3. 输出示例

转换后的 Markdown 文件将包含：
- 标题（# Heading 1, ## Heading 2 等）
- 列表（- 项目）
- 表格（| 列1 | 列2 |）
- 图片（![图片](assets/images/image1.png)）
- 超链接（[链接文本](url)）
- 数学公式（$行内公式$ 或 $$块级公式$$）

## 注意事项

- 数学公式转换目前基于简单的文本提取，复杂公式可能需要进一步处理。
- 某些高级 Word 样式（如艺术字、复杂边框）可能无法完全转换。
- 图片将保存到 `assets/images/` 目录，请确保该目录可写。

## 依赖库

主要依赖：
- `python-docx`：用于读取和操作 DOCX 文件。
- `marker-pdf`：用于读取和操作 PDF 文件。

## 许可证

本项目仅供学习参考，可根据需要修改和使用。