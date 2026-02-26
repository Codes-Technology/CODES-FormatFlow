import pdfplumber
from typing import List, Dict, Optional


class TableDetector:
    """Detect and extract tables from PDF files using pdfplumber."""
    
    def __init__(self, pdf_path: str):
        """
        Initialize the table detector.
        
        Args:
            pdf_path: Path to the PDF file
        """
        self.pdf_path = pdf_path
        
    def detect_tables(self) -> bool:
        """
        Check if the PDF contains tables.
        
        Returns:
            bool: True if tables are detected, False otherwise
        """
        try:
            with pdfplumber.open(self.pdf_path) as pdf:
                for page in pdf.pages:
                    tables = page.extract_tables()
                    if tables:
                        return True
            return False
        except Exception as e:
            print(f"Error detecting tables: {str(e)}")
            return False
    
    def extract_tables(self) -> List[Dict]:
        """
        Extract all tables from the PDF.
        
        Returns:
            List of dictionaries containing table data and page info
        """
        tables_data = []
        
        try:
            with pdfplumber.open(self.pdf_path) as pdf:
                for page_num, page in enumerate(pdf.pages, start=1):
                    tables = page.extract_tables()
                    
                    for table_idx, table in enumerate(tables):
                        if table:
                            tables_data.append({
                                'page': page_num,
                                'table_index': table_idx,
                                'data': table,
                                'num_rows': len(table),
                                'num_cols': len(table[0]) if table else 0
                            })
        except Exception as e:
            print(f"Error extracting tables: {str(e)}")
        
        return tables_data
    
    def get_table_count(self) -> int:
        """
        Get the total number of tables in the PDF.
        
        Returns:
            int: Number of tables found
        """
        return len(self.extract_tables())
