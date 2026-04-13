import os
import pymupdf
from typing import List, Dict, Any

class PdfHandler:
    """
    基于 pymupdf 库的 PDF 处理器。
    解析 PDF 并生成与 DocxHandler 格式相同的结构化 JSON。
    """
    def __init__(self, file_path: str|None = None, image_dir: str = "images"):
        """
        初始化 PdfHandler。
        :param file_path: PDF 文件路径
        :param image_dir: 图片保存目录
        """
        self.file_path = file_path
        self.image_dir = image_dir
        self.raw_json_data = None  # 存储 pymupdf 转换后的内部对象

        if not os.path.exists(image_dir):
            os.makedirs(image_dir)

        if file_path and os.path.exists(file_path):
            self._parse_pdf(file_path)

    def _parse_pdf(self, pdf_path: str):
        """
        使用 pymupdf 解析 PDF 并提取图片。
        pymupdf 的 Document 对象本身就包含了所有页面信息，
        因此直接存储 doc 对象以便后续逐步解析。
        """
        self.doc = pymupdf.open(pdf_path)
        self._extract_all_images()

    def _extract_all_images(self):
        """
        提取 PDF 中所有图片并保存到 self.image_dir。
        """
        for page_num, page in enumerate(self.doc):
            image_list = page.get_images()
            for img_index, img in enumerate(image_list):
                xref = img[0]
                try:
                    pix = pymupdf.Pixmap(self.doc, xref)
                    if pix.n - pix.alpha > 3:
                        pix = pymupdf.Pixmap(pymupdf.csRGB, pix)
                    img_filename = f"page_{page_num+1}_img_{img_index+1}.png"
                    pix.save(os.path.join(self.image_dir, img_filename))
                    pix = None
                except Exception as e:
                    print(f"Error extracting image on page {page_num+1}: {e}")

    def get_full_details(self) -> List[Dict[str, Any]]:
        """
        返回与 DocxHandler.get_full_details() 格式完全一致的结构化数据。
        """
        if not self.doc:
            return []

        content_details = []
        for page_num, page in enumerate(self.doc):
            # 获取页面的结构化字典
            page_dict = page.get_text("dict")
            # 1. 先处理图片块（Figure）
            content_details.extend(self._extract_figure_blocks_from_page(page_dict, page_num))
            # 2. 处理文本块（包括表格识别）
            content_details.extend(self._extract_text_blocks_from_page(page, page_dict, page_num))
        return content_details

    def _extract_figure_blocks_from_page(self, page_dict: dict, page_num: int) -> List[Dict[str, Any]]:
        """
        从页面字典中提取独立的图片块（Figure）。
        """
        figures = []
        for block in page_dict.get("blocks", []):
            if block.get("type") == 1:  # 图片块
                img_filename = f"page_{page_num+1}_block_{block['number']}.png"
                figures.append({
                    "type": "paragraph",
                    "style": "Figure",
                    "list_level": None,
                    "runs": [{
                        "type": "image",
                        "src": os.path.join(self.image_dir, img_filename)
                    }]
                })
        return figures

    def _extract_text_blocks_from_page(self, page: pymupdf.Page, page_dict: dict, page_num: int) -> List[Dict[str, Any]]:
        """
        处理文本块和表格。
        首先使用 find_tables() 识别表格区域，然后解析普通文本块时跳过这些区域。
        """
        content = []
        
        # 查找表格
        tables = page.find_tables()
        table_rects = []
        if tables and tables.tables:
            for table in tables:
                table_rects.append(table.bbox)
                content.append(self._parse_table_block(table))
        
        # 解析文本块，跳过表格区域
        for block in page_dict.get("blocks", []):
            if block.get("type") != 0:
                continue
            
            # 检查该文本块是否在表格区域内
            bbox = block.get("bbox")
            if bbox and self._is_inside_any_table(bbox, table_rects):
                continue
            
            # 处理普通文本块
            block_content = self._parse_text_block(block)
            if block_content:
                content.append(block_content)
        
        return content

    def _is_inside_any_table(self, block_bbox: List[float], table_rects: List[List[float]]) -> bool:
        """
        检查文本块是否位于任何表格区域内。
        """
        if not table_rects:
            return False
        x0, y0, x1, y1 = block_bbox
        for t_bbox in table_rects:
            tx0, ty0, tx1, ty1 = t_bbox
            if (x0 >= tx0 and x1 <= tx1 and y0 >= ty0 and y1 <= ty1):
                return True
        return False

    def _parse_text_block(self, block: dict) -> Dict[str, Any]:
        """
        解析文本块，处理格式（加粗、斜体、公式等）。
        pymupdf 的 dict 结构包含 lines -> spans 层次，spans 提供了字体、字号等信息。
        """
        runs = []
        style = "Default"
        
        # 检测标题（基于字号或字体）
        for line in block.get("lines", []):
            for span in line.get("spans", []):
                font_size = span.get("size", 0)
                font_name = span.get("font", "").lower()
                flags = span.get("flags", 0)
                
                # 构建 run 对象
                run = {
                    "type": "text",
                    "text": span.get("text", ""),
                    "bold": bool(flags & 2**4),  # 粗体标志
                    "italic": bool(flags & 2**1), # 斜体标志
                    "underline": bool(flags & 2**2), # 下划线标志
                    "color": self._get_span_color(span),
                    "font_size": font_size
                }
                runs.append(run)
                
                # 根据字号或字体判断标题样式
                if font_size > 14:
                    style = "Heading1"
                elif font_size > 12:
                    style = "Heading2"
                elif "bold" in font_name:
                    style = "Heading3"
        
        # 检测列表层级（根据缩进）
        list_level = None
        if block.get("bbox"):
            left_margin = block["bbox"][0]
            if left_margin > 50:
                list_level = int(left_margin / 30)
        
        return {
            "type": "paragraph",
            "style": style,
            "list_level": list_level,
            "runs": runs
        }

    def _get_span_color(self, span: dict) -> str:
        """
        从 span 中提取颜色信息（如果存在）。
        """
        color = span.get("color", 0)
        if color:
            # 简单处理：将整数颜色转换为十六进制
            return f"#{color:06x}"
        return "Default"

    def _parse_formula_block(self, block: dict) -> Dict[str, Any]:
        """
        解析块级公式（如果需要）。
        """
        # pymupdf 原生不直接识别公式，可以结合特殊字体或位置判断
        # 这里留作扩展点
        return {
            "type": "paragraph",
            "style": "Formula",
            "list_level": None,
            "runs": [{
                "type": "formula",
                "text": "",
                "is_block": True
            }]
        }

    def _parse_table_block(self, table) -> Dict[str, Any]:
        """
        解析表格。
        pymupdf 的 table 对象提供了 .extract() 方法直接获取纯文本内容。
        """
        rows_data = []
        cells = table.extract()  # 返回二维列表，每个元素为单元格文本
        if cells:
            for row_cells in cells:
                row_data = []
                for cell_text in row_cells:
                    # 每个单元格可以进一步解析为段落（这里简化处理）
                    cell_para = [{
                        "type": "paragraph",
                        "style": "Default",
                        "list_level": None,
                        "runs": [{"type": "text", "text": cell_text or ""}]
                    }]
                    row_data.append(cell_para)
                rows_data.append(row_data)
        
        return {"type": "table", "rows": rows_data}

    def _parse_figure_block(self, block: dict) -> Dict[str, Any]:
        """
        解析图片块（备用方法）。
        """
        return {
            "type": "paragraph",
            "style": "Figure",
            "list_level": None,
            "runs": []
        }

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

    def close(self):
        """释放资源"""
        if hasattr(self, 'doc') and self.doc:
            self.doc.close()