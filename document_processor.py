import os
from pdf2docx import Converter
from docx import Document
from utils.table_detector import TableDetector
from utils.style_manager import StyleManager
from config import TEMPLATE_DOCX

class DocumentProcessor:
    """Main document processor for converting PDFs and TXTs to Word documents."""
    
    def __init__(self, template_path: str = TEMPLATE_DOCX):
        self.template_path = template_path
        self.style_manager = StyleManager(template_path)
    
    def universal_extract(self, input_path: str, output_path: str) -> dict:
        """Universal Extraction Engine - Converts any supported file type."""
        result = {'success': False, 'output_file': output_path, 'tables_found': 0, 'input_type': None, 'error': None}
        
        try:
            file_ext = os.path.splitext(input_path)[1].lower()
            result['input_type'] = file_ext
            
            print(f"[Universal Extraction Engine] Processing: {input_path} ({file_ext})")
            
            if file_ext == '.pdf':
                result = self._extract_from_pdf(input_path, output_path)
            elif file_ext == '.txt':
                result = self._extract_from_text(input_path, output_path)
            elif file_ext == '.docx':
                result = self._extract_from_docx(input_path, output_path)
            else:
                result['error'] = f"Unsupported file type: {file_ext}"
                return result
            
            print(f"[Universal Extraction Engine] ✓ Success: {output_path} | Tables: {result.get('tables_found', 0)}")
            
        except Exception as e:
            result['error'] = str(e)
            print(f"[Universal Extraction Engine] Error: {str(e)}")
            
        return result
    
    def _extract_from_pdf(self, pdf_path: str, output_path: str) -> dict:
        result = {'success': False, 'output_file': output_path, 'tables_found': 0, 'input_type': '.pdf', 'error': None}
        try:
            table_detector = TableDetector(pdf_path)
            result['tables_found'] = table_detector.get_table_count()
            
            cv = Converter(pdf_path)
            cv.convert(output_path)
            cv.close()
            
            result['success'] = os.path.exists(output_path)
        except Exception as e:
            result['error'] = str(e)
        return result
    
    def _extract_from_text(self, text_path: str, output_path: str) -> dict:
        result = {'success': False, 'output_file': output_path, 'tables_found': 0, 'input_type': '.txt', 'error': None}
        try:
            with open(text_path, 'r', encoding='utf-8') as f:
                content = f.read()
            
            doc = Document()
            for line in content.split('\n'):
                doc.add_paragraph(line)
            
            doc.save(output_path)
            result['success'] = True
        except Exception as e:
            result['error'] = str(e)
        return result
    
    def _extract_from_docx(self, docx_path: str, output_path: str) -> dict:
        result = {'success': False, 'output_file': output_path, 'tables_found': 0, 'input_type': '.docx', 'error': None}
        try:
            doc = Document(docx_path)
            result['tables_found'] = len(doc.tables)
            doc.save(output_path)
            result['success'] = True
        except Exception as e:
            result['error'] = str(e)
        return result