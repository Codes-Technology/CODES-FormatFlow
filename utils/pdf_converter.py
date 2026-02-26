import os
import sys
import pythoncom
import win32com.client

def convert_to_pdf(docx_path, pdf_path):
    """
    Convert DOCX to PDF using direct win32com dispatch.
    Handles CoInitialize and cleanup explicitly.
    """
    docx_path = os.path.abspath(docx_path)
    pdf_path = os.path.abspath(pdf_path)
    
    word = None
    doc = None
    
    try:
        # Initialize COM for this thread
        pythoncom.CoInitialize()
        
        # Dispatch Word application
        try:
            word = win32com.client.Dispatch("Word.Application")
        except AttributeError:
             # Fallback for some caching issues
             from win32com.client import dynamic
             word = dynamic.Dispatch("Word.Application")
             
        word.Visible = False
        
        # Open document
        print(f"Opening: {docx_path}")
        doc = word.Documents.Open(docx_path, ReadOnly=True)
        
        # Save as PDF (Format 17)
        print(f"Saving as PDF: {pdf_path}")
        wdFormatPDF = 17
        doc.SaveAs(pdf_path, FileFormat=wdFormatPDF)
        
        print("Conversion successful.")
        
    except Exception as e:
        print(f"PDF Conversion Failed: {e}")
        raise e
        
    finally:
        # Cleanup
        if doc:
            try:
                doc.Close(SaveChanges=0) # wdDoNotSaveChanges
            except: pass
        if word:
            try:
                word.Quit()
            except: pass
        
        try:
            pythoncom.CoUninitialize()
        except: pass

if __name__ == "__main__":
    if len(sys.argv) > 2:
        convert_to_pdf(sys.argv[1], sys.argv[2])
    else:
        print("Usage: python pdf_converter.py input.docx output.pdf")
