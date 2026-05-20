"""
OCR Engine wrapper for EasyOCR
"""

import tempfile
import os
import fitz  # PyMuPDF
import easyocr

class OCREngine:
    """Handles OCR operations"""
    
    _instance = None
    _reader = None
    
    def __new__(cls):
        if cls._instance is None:
            cls._instance = super(OCREngine, cls).__new__(cls)
        return cls._instance
    
    def __init__(self):
        if self._reader is None:
            print("Chargement d'EasyOCR...")
            self._reader = easyocr.Reader(['fr', 'en'], gpu=False, verbose=False)
            print("EasyOCR prêt!")
    
    def ocr_page(self, page, dpi=150):
        """Perform OCR on a single PDF page"""
        zoom = dpi / 72
        mat = fitz.Matrix(zoom, zoom)
        pix = page.get_pixmap(matrix=mat)
        
        with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp_img:
            tmp_img.write(pix.tobytes("png"))
            img_path = tmp_img.name
        
        try:
            result = self._reader.readtext(img_path, paragraph=False)
            text = " ".join([item[1] for item in result])
            return text if text.strip() else ""
        except Exception as e:
            print(f"OCR error: {e}")
            return ""
        finally:
            if os.path.exists(img_path):
                os.unlink(img_path)
    
    def is_pdf_scanned(self, doc, pages_to_check=3):
        """Check if PDF is scanned (lacks extractable text)"""
        total_pages = len(doc)
        for page_num in range(min(pages_to_check, total_pages)):
            test_text = doc[page_num].get_text()
            if test_text and len(test_text.strip()) > 100:
                return False
        return True