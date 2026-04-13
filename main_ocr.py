"""
PDF to Word Converter with OCR Support
Handles both digital PDFs and scanned PDFs
"""

import PyPDF2
import fitz  # PyMuPDF
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from PIL import Image
import io
import os
import re

# Try to import OCR libraries (will be installed via requirements.txt)
try:
    import pytesseract
    from pdf2image import convert_from_path
    OCR_AVAILABLE = True
except ImportError:
    OCR_AVAILABLE = False
    print("Warning: OCR libraries not installed. Scanned PDFs will not work properly.")

def extract_text_with_pypdf2(pdf_path):
    """Extract text from digital PDFs using PyPDF2"""
    text_by_page = []
    try:
        with open(pdf_path, "rb") as pdf_file:
            pdf_reader = PyPDF2.PdfReader(pdf_file)
            for page_num in range(len(pdf_reader.pages)):
                page = pdf_reader.pages[page_num]
                text = page.extract_text()
                if text:
                    # Clean up the text
                    text = re.sub(r'\s+', ' ', text)
                    text_by_page.append(text.strip())
                else:
                    text_by_page.append("")
        return text_by_page
    except Exception as e:
        print(f"PyPDF2 error: {e}")
        return []

def extract_text_with_ocr(pdf_path):
    """Extract text from scanned PDFs using OCR"""
    if not OCR_AVAILABLE:
        return None
    
    text_by_page = []
    try:
        # Convert PDF pages to images at 300 DPI
        images = convert_from_path(pdf_path, dpi=300)
        
        for page_num, image in enumerate(images):
            # Preprocess image for better OCR
            # Convert to grayscale
            if image.mode != 'L':
                image = image.convert('L')
            
            # Run OCR on the image
            custom_config = r'--oem 3 --psm 6'
            text = pytesseract.image_to_string(image, config=custom_config)
            
            if text.strip():
                # Clean up OCR text
                text = re.sub(r'\s+', ' ', text)
                text_by_page.append(text.strip())
            else:
                text_by_page.append("No text found on this page.")
        
        return text_by_page
    except Exception as e:
        print(f"OCR error: {e}")
        return None

def extract_images_with_pymupdf(pdf_path):
    """Extract images from PDF using PyMuPDF"""
    images_by_page = {}
    try:
        pdf_document = fitz.open(pdf_path)
        for page_num in range(len(pdf_document)):
            page = pdf_document[page_num]
            image_list = page.get_images()
            images = []
            for img_index, img in enumerate(image_list):
                try:
                    xref = img[0]
                    base_image = pdf_document.extract_image(xref)
                    image_bytes = base_image["image"]
                    image = Image.open(io.BytesIO(image_bytes))
                    images.append(image)
                except Exception as img_error:
                    print(f"Error extracting image {img_index}: {img_error}")
            images_by_page[page_num + 1] = images
        pdf_document.close()
        return images_by_page
    except Exception as e:
        print(f"Image extraction error: {e}")
        return {}

def clean_ocr_text(text):
    """Clean and improve OCR extracted text"""
    if not text:
        return ""
    
    # Remove excessive spaces
    text = re.sub(r'\s+', ' ', text)
    
    # Fix common OCR issues
    replacements = {
        r'\|': 'I',  # Pipe to I
        r'0(?=\d)': 'O',  # Zero before number to O (contextual)
        r'1(?=\s)': 'I',  # One before space to I
    }
    
    for pattern, replacement in replacements.items():
        text = re.sub(pattern, replacement, text)
    
    return text.strip()

def pdf_to_word(pdf_path, word_path):
    """
    Convert PDF to Word document with text and images
    Automatically detects and handles scanned PDFs
    """
    try:
        print(f"Processing PDF: {pdf_path}")
        
        # First try regular text extraction (for digital PDFs)
        text_by_page = extract_text_with_pypdf2(pdf_path)
        
        # Check if we got meaningful text (at least 100 chars on first page)
        has_text = False
        if text_by_page and len(text_by_page) > 0:
            first_page_text = text_by_page[0]
            has_text = len(first_page_text) > 100
        
        # If no text found, use OCR (for scanned PDFs)
        if not has_text and OCR_AVAILABLE:
            print("No selectable text found, using OCR...")
            text_by_page = extract_text_with_ocr(pdf_path)
            if text_by_page:
                # Clean OCR text
                text_by_page = [clean_ocr_text(text) for text in text_by_page]
        elif not has_text and not OCR_AVAILABLE:
            print("Warning: No text found and OCR not available.")
            text_by_page = ["This appears to be a scanned PDF. OCR support is not installed on the server."]
        
        # Extract images (optional - can be disabled for cleaner output)
        images_by_page = extract_images_with_pymupdf(pdf_path)
        
        # Create Word document
        print("Creating Word document...")
        doc = Document()
        
        # Set default font
        style = doc.styles['Normal']
        style.font.name = 'Arial'
        style.font.size = Pt(11)
        
        # Add content page by page
        for page_num in range(len(text_by_page)):
            # Add page header
            heading = doc.add_heading(f'Page {page_num + 1}', level=1)
            heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            # Add text content
            if page_num < len(text_by_page) and text_by_page[page_num]:
                text = text_by_page[page_num]
                if len(text) > 10:  # Only add if there's meaningful text
                    paragraph = doc.add_paragraph()
                    run = paragraph.add_run(text)
                    run.font.size = Pt(11)
                else:
                    doc.add_paragraph("[No text extracted from this page]")
            else:
                doc.add_paragraph("[No text extracted from this page]")
            
            # Add images from this page (limit to first 3 images per page to keep file size reasonable)
            if page_num + 1 in images_by_page and images_by_page[page_num + 1]:
                images = images_by_page[page_num + 1][:3]  # Max 3 images per page
                if images:
                    doc.add_paragraph()  # Add spacing
                    for img_index, img in enumerate(images):
                        try:
                            # Resize large images
                            if img.width > 500:
                                ratio = 500 / img.width
                                new_width = 500
                                new_height = int(img.height * ratio)
                                img = img.resize((new_width, new_height), Image.Resampling.LANCZOS)
                            
                            # Save image temporarily
                            temp_img_path = f"temp_img_{page_num}_{img_index}.png"
                            img.save(temp_img_path, 'PNG')
                            doc.add_picture(temp_img_path, width=Inches(5))
                            # Clean up
                            if os.path.exists(temp_img_path):
                                os.remove(temp_img_path)
                        except Exception as img_error:
                            print(f"Error adding image: {img_error}")
            
            # Add page break except after last page
            if page_num < len(text_by_page) - 1:
                doc.add_page_break()
        
        # Save the document
        doc.save(word_path)
        print(f"Word document saved to: {word_path}")
        
        # Verify file was created
        if os.path.exists(word_path) and os.path.getsize(word_path) > 0:
            return True
        else:
            raise Exception("Output file is empty or was not created")
        
    except Exception as e:
        print(f"Error in pdf_to_word: {e}")
        raise e

# For command line testing
if __name__ == "__main__":
    import sys
    if len(sys.argv) > 1:
        input_pdf = sys.argv[1]
        output_word = sys.argv[2] if len(sys.argv) > 2 else "output.docx"
        pdf_to_word(input_pdf, output_word)
        print(f"Conversion complete! Output saved to {output_word}")
