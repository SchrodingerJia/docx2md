from .docx_handler import DocxHandler
from .pdf_handler import PdfHandler
from .translator import MarkdownGenerater

if __name__ == "__main__":
    File = r'YOUR_FILE_PATH_HERE'
    MDFile = r'YOUR_OUTPUT_PATH_HERE'
    if File.endswith('.docx'):
        reader = DocxHandler(File)
    elif File.endswith('.pdf'):
        reader = PdfHandler(File)
    details = reader.get_full_details()
    converter = MarkdownGenerater(details)
    markdown_text = converter.convert()
    converter.save_to_file(MDFile)