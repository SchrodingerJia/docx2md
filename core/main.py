from .docx_handler import DocxHandler
from .docx_to_markdown import DocxToMarkdown

if __name__ == "__main__":
    DocxFile = r'YOUR_FILE_PATH_HERE'
    MDFile = r'YOUR_OUTPUT_PATH_HERE'
    reader = DocxHandler(DocxFile)
    details = reader.get_full_details()
    converter = DocxToMarkdown(details)
    markdown_text = converter.convert()
    converter.save_to_file(MDFile)