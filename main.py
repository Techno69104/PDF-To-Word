import PyPDF2
import fitz  # PyMuPDF
from docx import Document
from docx.shared import Inches
from PIL import Image
import io
import os
import pytesseract
from pdf2image import convert_from_path

def extract_text_with_ocr(pdf_path):
    """Extract text from scanned PDFs using OCR"""
    try:
        # Convert PDF pages to images
        images = convert_from_path(pdf_path, dpi=300)
        all_text = []
        
        for page_num, image in enumerate(images):
            # Run OCR on each page
            text = pytesseract.image_to_string(image)
            all_text.append(text if text.strip() else "No text found on this page.")
        
        return all_text
    except Exception as e:
        print(f"OCR extraction error: {e}")
        return None

def extract_text_with_pypdf2(pdf_path):
    """Extract text from digital PDFs using PyPDF2"""
    text_by_page = []
    try:
        with open(pdf_path, "rb") as pdf_file:
            pdf_reader = PyPDF2.PdfReader(pdf_file)
            for page_num in range(len(pdf_reader.pages)):
                page = pdf_reader.pages[page_num]
                text = page.extract_text()
                text_by_page.append(text if text else "")
        return text_by_page
    except Exception as e:
        print(f"PyPDF2 error: {e}")
        return []

def pdf_to_word(pdf_path, word_path):
    """Convert PDF to Word - automatically chooses best method"""
    
    # First try regular text extraction
    text_by_page = extract_text_with_pypdf2(pdf_path)
    
    # Check if text was found (digital PDF)
    has_text = any(len(text.strip()) > 50 for text in text_by_page)
    
    if not has_text:
        # Fall back to OCR for scanned PDFs
        print("No selectable text found, using OCR...")
        text_by_page = extract_text_with_ocr(pdf_path)
        if text_by_page is None:
            text_by_page = ["Could not extract text. PDF may be corrupted."] * 10
    
    # Extract images (optional, can be disabled for cleaner text)
    images_by_page = extract_images_with_pymupdf(pdf_path)
    
    # Create Word document
    doc = Document()
    
    for page_num, text in enumerate(text_by_page, start=1):
        doc.add_heading(f'Page {page_num}', level=1)
        
        # Add cleaned text (remove extra spaces, fix line breaks)
        cleaned_text = ' '.join(text.split())
        doc.add_paragraph(cleaned_text)
        
        if page_num < len(text_by_page):
            doc.add_page_break()
    
    doc.save(word_path)
    return True
