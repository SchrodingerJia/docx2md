import os
import tempfile
from typing import List, Dict, Any
from marker.converters.pdf import PdfConverter
from marker.models import create_model_dict

class PdfHandler:
    """
    基于 marker-pdf 库的 PDF 处理器。
    使用 Python API 解析 PDF 并生成与 DocxHandler 格式相同的结构化 JSON。
    """
    def __init__(self, file_path: str|None = None, image_dir: str = "images"):
        """
        初始化 PdfHandler。
        :param file_path: PDF 文件路径
        :param image_dir: 图片保存目录
        """
        self.file_path = file_path
        self.image_dir = image_dir
        self.raw_json_data = None  # 存储 marker 转换后的内部对象

        if not os.path.exists(image_dir):
            os.makedirs(image_dir)

        if file_path and os.path.exists(file_path):
            self._parse_pdf(file_path)

    def _parse_pdf(self, pdf_path: str):
        """
        使用 marker 的 Python API 解析 PDF 并提取图片。
        """
        # 创建临时目录用于 marker 输出图片
        with tempfile.TemporaryDirectory() as tmpdir:
            # 配置 converter
            config = {
                "output_format": "json",   # 获取结构化数据
                "extract_images": True,    # 提取图片
                "image_dir": tmpdir,       # marker 会将图片输出到此目录
            }
            converter = PdfConverter(
                artifact_dict=create_model_dict(),
                config=config
            )

            # 执行转换，返回一个 RenderedDocument 对象
            rendered = converter(pdf_path)
            self.raw_json_data = rendered

            # 将 marker 提取的图片移动到我们的 image_dir
            marker_images_dir = os.path.join(tmpdir, "images")
            if os.path.exists(marker_images_dir):
                for img_file in os.listdir(marker_images_dir):
                    src = os.path.join(marker_images_dir, img_file)
                    dst = os.path.join(self.image_dir, img_file)
                    if not os.path.exists(dst):
                        os.rename(src, dst)

    def get_full_details(self) -> List[Dict[str, Any]]:
        """
        返回与 DocxHandler.get_full_details() 格式完全一致的结构化数据。
        """
        if not self.raw_json_data:
            return []

        content_details = []
        # rendered.children 是页面列表
        for page in self.raw_json_data.children:
            for block in page.children:
                block_type = block.block_type

                if block_type in ("Text", "TextInlineMath"):
                    content_details.append(self._parse_text_block(block))
                elif block_type == "Formula":
                    content_details.append(self._parse_formula_block(block))
                elif block_type == "Table":
                    content_details.append(self._parse_table_block(block))
                elif block_type == "Figure":
                    content_details.append(self._parse_figure_block(block))
                # 可根据需要扩展其他类型，如 ListItem

        return content_details

    def _parse_text_block(self, block) -> Dict[str, Any]:
        """解析文本块，处理格式（加粗、斜体、公式等）"""
        runs = []
        for child in getattr(block, 'children', []):
            child_type = child.block_type

            if child_type == "Span":
                # 普通文本片段
                runs.append({
                    "type": "text",
                    "text": getattr(child, 'text', ""),
                    "bold": getattr(child, 'bold', False),
                    "italic": getattr(child, 'italic', False),
                    "underline": False,
                    "color": "Default",
                    "font_size": getattr(child, 'font_size', None)
                })
            elif child_type == "Formula":
                # 行内公式
                runs.append({
                    "type": "formula",
                    "text": getattr(child, 'text', ""),
                    "is_block": False
                })
            elif child_type == "Picture":
                # 行内图片
                img_filename = getattr(child, 'image_filename', None)
                if img_filename:
                    runs.append({
                        "type": "image",
                        "src": os.path.join(self.image_dir, img_filename)
                    })

        # 样式和列表层级
        style = "Default"
        if "Heading" in block.block_type:
            style = block.block_type  # e.g., "Heading1"
        list_level = getattr(block, 'list_level', None)

        return {
            "type": "paragraph",
            "style": style,
            "list_level": list_level,
            "runs": runs
        }

    def _parse_formula_block(self, block) -> Dict[str, Any]:
        """解析块级公式"""
        return {
            "type": "paragraph",
            "style": "Formula",
            "list_level": None,
            "runs": [{
                "type": "formula",
                "text": getattr(block, 'text', ""),
                "is_block": True
            }]
        }

    def _parse_table_block(self, block) -> Dict[str, Any]:
        """解析表格"""
        rows_data = []
        for row in getattr(block, 'rows', []):
            row_cells = []
            for cell in row.cells:
                cell_paras = []
                for cell_block in cell.children:
                    if cell_block.block_type in ("Text", "TextInlineMath"):
                        cell_paras.append(self._parse_text_block(cell_block))
                row_cells.append(cell_paras)
            rows_data.append(row_cells)
        return {"type": "table", "rows": rows_data}

    def _parse_figure_block(self, block) -> Dict[str, Any]:
        """解析图片块"""
        img_filename = getattr(block, 'image_filename', None)
        if img_filename:
            return {
                "type": "paragraph",
                "style": "Figure",
                "list_level": None,
                "runs": [{
                    "type": "image",
                    "src": os.path.join(self.image_dir, img_filename)
                }]
            }
        return {"type": "paragraph", "style": "Figure", "list_level": None, "runs": []}

    def find_paragraphs_with_keyword(self, keyword: str) -> List[str]:
        """查找包含关键字的段落文本"""
        details = self.get_full_details()
        found = []
        for item in details:
            if item["type"] == "paragraph":
                for run in item["runs"]:
                    if run["type"] == "text" and keyword in run.get("text", ""):
                        found.append(run["text"])
        return found

    def get_tables_data(self) -> List[List[List[str]]]:
        """获取所有表格的纯文本数据"""
        details = self.get_full_details()
        all_tables = []
        for item in details:
            if item["type"] == "table":
                table_data = []
                for row in item["rows"]:
                    row_data = []
                    for cell in row:
                        cell_text = ""
                        for para in cell:
                            for run in para["runs"]:
                                if run["type"] == "text":
                                    cell_text += run["text"]
                            cell_text += "\n"
                        row_data.append(cell_text.strip())
                    table_data.append(row_data)
                all_tables.append(table_data)
        return all_tables