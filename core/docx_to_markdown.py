import re

class DocxToMarkdown:
    def __init__(self, structured_data):
        """
        :param structured_data: 由 EnhancedDocxReader 生成的结构化数据列表
        """
        self.data = structured_data
        # 映射 Word 样式名到 Markdown 标题
        self.heading_map = {
            'Heading 1': '# ',
            'Heading 2': '## ',
            'Heading 3': '### ',
            'Heading 4': '#### ',
            'Heading 5': '##### ',
            'Heading 6': '###### ',
            'Title': '# ',
            'Subtitle': '## '
        }
    def convert(self):
        """核心转换函数"""
        md_segments = []
        for item in self.data:
            if item['type'] == 'paragraph':
                md_segments.append(self._parse_paragraph(item))
            elif item['type'] == 'table':
                md_segments.append(self._parse_table(item))
        # 使用双换行符连接各块，保持段落间距
        return "\n\n".join(md_segments)
    def _parse_paragraph(self, para_info):
        content = ""
        prefix = self.heading_map.get(para_info['style'], "")
        if para_info.get('list_level') is not None:
            indent = "  " * para_info['list_level']
            prefix = f"{indent}- "
        for run in para_info['runs']:
            if run['type'] == 'text':
                run_text = self._format_text(run)
                if run.get('link_url'):
                    run_text = f"[{run_text}]({run['link_url']})"
                if content.endswith("**") and run_text.startswith("**"):
                    content = content[:-2] + run_text[2:]
                else:
                    content += run_text
            elif run['type'] == 'image':
                content += f"![图片]({run.get('src', '')})"
            # --- 新增：处理公式类型 ---
            elif run['type'] == 'formula':
                formula_text = run['text']
                if run.get('is_block'):
                    # 块级公式另起一行
                    content += f"\n$$\n{formula_text}\n$$\n"
                else:
                    # 行内公式
                    content += f"${formula_text}$"
        if not prefix and not content.strip():
            return ""
        return f"{prefix}{content}".strip()
    def _format_text(self, run):
        text = run['text']
        if not text or text.isspace(): return text
        leading_space = re.match(r"^\s*", text).group()
        trailing_space = re.search(r"\s*$", text).group()
        core_text = text.strip()
        if run.get('bold') and run.get('italic'):
            core_text = f"***{core_text}***"
        elif run.get('bold'):
            core_text = f"**{core_text}**"
        elif run.get('italic'):
            core_text = f"*{core_text}*"
        if run.get('underline'):
            core_text = f"<u>{core_text}</u>"
        color = run.get('color')
        if color and color != "Default" and color != "000000":
            hex_color = f"#{color}" if not color.startswith('#') else color
            core_text = f'<span style="color:{hex_color}">{core_text}</span>'
        return f"{leading_space}{core_text}{trailing_space}"
    def _parse_table(self, table_info):
        rows = table_info.get('rows', [])
        if not rows: return ""
        md_table = []
        column_count = 0
        for r_idx, row in enumerate(rows):
            formatted_cells = []
            for cell_paras in row:
                # cell_paras 现在直接复用了 _parse_paragraph 的返回结果
                cell_content_parts = [self._parse_paragraph(p) for p in cell_paras]
                cell_final_text = "<br>".join([c for c in cell_content_parts if c])
                cell_final_text = cell_final_text.replace("|", "\\|")
                formatted_cells.append(cell_final_text)
            if r_idx == 0:
                column_count = len(formatted_cells)
            md_table.append(f"| {' | '.join(formatted_cells)} |")
            if r_idx == 0:
                md_table.append(f"| {' | '.join(['---'] * column_count)} |")
        return "\n".join(md_table)
    def save_to_file(self, filename="output.md"):
        with open(filename, 'w', encoding='utf-8') as f:
            f.write(self.convert())