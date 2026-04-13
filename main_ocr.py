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
import sys

# Try to import OCR libraries (will be installed via requirements.txt)
try:
    import pytesseract
    from pdf2image import convert_from_path
    OCR_AVAILABLE = True
    print("OCR libraries loaded successfully", file=sys.stderr)
except ImportError as e:
    OCR_AVAILABLE = False
    print(f"Warning: OCR libraries not installed. Scanned PDFs will not work properly. Error: {e}", file=sys.stderr)

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
        print(f"PyPDF2 error: {e}", file=sys.stderr)
        return []

def extract_text_with_ocr(pdf_path):
    """Extract text from scanned PDFs using OCR"""
    if not OCR_AVAILABLE:
        print("OCR not available - libraries missing", file=sys.stderr)
        return None
    
    text_by_page = []
    try:
        print(f"Starting OCR on: {pdf_path}", file=sys.stderr)
        
        # Convert PDF pages to images at 300 DPI
        images = convert_from_path(pdf_path, dpi=300)
        print(f"Converted {len(images)} pages to images", file=sys.stderr)
        
        for page_num, image in enumerate(images):
            print(f"Processing page {page_num + 1} with OCR...", file=sys.stderr)
            
            # Preprocess image for better OCR
            # Convert to grayscale
            if image.mode != 'L':
                image = image.convert('L')
            
            # Increase contrast for better recognition
            from PIL import ImageEnhance
            enhancer = ImageEnhance.Contrast(image)
            image = enhancer.enhance(1.5)
            
            # Run OCR on the image
            custom_config = r'--oem 3 --psm 6 -c tessedit_char_whitelist=ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789.,!?;:()[]{}<>/\\|@#$%^&*+=_- '
            text = pytesseract.image_to_string(image, config=custom_config)
            
            if text and text.strip():
                # Clean up OCR text
                text = re.sub(r'\s+', ' ', text)
                text_by_page.append(text.strip())
                print(f"Page {page_num + 1}: Extracted {len(text)} characters", file=sys.stderr)
            else:
                print(f"Page {page_num + 1}: No text found", file=sys.stderr)
                text_by_page.append("[No readable text found on this page]")
        
        return text_by_page if text_by_page else None
        
    except Exception as e:
        print(f"OCR error: {e}", file=sys.stderr)
        import traceback
        traceback.print_exc(file=sys.stderr)
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
                    print(f"Error extracting image {img_index}: {img_error}", file=sys.stderr)
            images_by_page[page_num + 1] = images
        pdf_document.close()
        return images_by_page
    except Exception as e:
        print(f"Image extraction error: {e}", file=sys.stderr)
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
        print(f"Processing PDF: {pdf_path}", file=sys.stderr)
        
        # Check if file exists
        if not os.path.exists(pdf_path):
            raise Exception(f"PDF file not found: {pdf_path}")
        
        # First try regular text extraction (for digital PDFs)
        text_by_page = extract_text_with_pypdf2(pdf_path)
        
        # Check if we got meaningful text (at least 100 chars on first page)
        has_text = False
        if text_by_page and len(text_by_page) > 0:
            first_page_text = text_by_page[0] if text_by_page[0] else ""
            has_text = len(first_page_text) > 100
            print(f"Digital text extraction: found {len(first_page_text)} chars on page 1", file=sys.stderr)
        
        # If no text found, use OCR (for scanned PDFs)
        if not has_text:
            print("No selectable text found, attempting OCR...", file=sys.stderr)
            if OCR_AVAILABLE:
                ocr_text = extract_text_with_ocr(pdf_path)
                if ocr_text and len(ocr_text) > 0:
                    text_by_page = ocr_text
                    # Clean OCR text
                    text_by_page = [clean_ocr_text(text) for text in text_by_page]
                    print(f"OCR successful! Extracted {len(text_by_page)} pages", file=sys.stderr)
                else:
                    print("OCR returned no text", file=sys.stderr)
                    text_by_page = ["[This appears to be a scanned PDF. OCR processing could not extract readable text. The document may be handwritten or have poor image quality.]"]
            else:
                print("OCR not available - install pytesseract and pdf2image", file=sys.stderr)
                text_by_page = ["[This appears to be a scanned PDF. OCR support is not installed on the server. Please use a digital PDF or contact support.]"]
        
        # Ensure text_by_page is a list with at least one element
        if not text_by_page or not isinstance(text_by_page, list):
            text_by_page = ["[No text could be extracted from this PDF]"]
        
        # Extract images (optional - can be disabled for cleaner output)
        images_by_page = extract_images_with_pymupdf(pdf_path)
        
        # Create Word document
        print("Creating Word document...", file=sys.stderr)
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
                    # Split long text into paragraphs
                    paragraphs = text.split('. ')
                    for para in paragraphs:
                        if para.strip():
                            paragraph = doc.add_paragraph()
                            run = paragraph.add_run(para.strip() + ('.' if not para.endswith('.') else ''))
                            run.font.size = Pt(11)
                else:
                    doc.add_paragraph(text)
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
                            temp_img_path = f"/tmp/temp_img_{page_num}_{img_index}.png"
                            img.save(temp_img_path, 'PNG')
                            doc.add_picture(temp_img_path, width=Inches(5))
                            # Clean up
                            if os.path.exists(temp_img_path):
                                os.remove(temp_img_path)
                        except Exception as img_error:
                            print(f"Error adding image: {img_error}", file=sys.stderr)
            
            # Add page break except after last page
            if page_num < len(text_by_page) - 1:
                doc.add_page_break()
        
        # Save the document
        doc.save(word_path)
        print(f"Word document saved to: {word_path} (Size: {os.path.getsize(word_path)} bytes)", file=sys.stderr)
        
        # Verify file was created
        if os.path.exists(word_path) and os.path.getsize(word_path) > 0:
            return True
        else:
            raise Exception("Output file is empty or was not created")
        
    except Exception as e:
        print(f"Error in pdf_to_word: {e}", file=sys.stderr)
        import traceback
        traceback.print_exc(file=sys.stderr)
        raise e

# For command line testing
if __name__ == "__main__":
    import sys
    if len(sys.argv) > 1:
        input_pdf = sys.argv[1]
        output_word = sys.argv[2] if len(sys.argv) > 2 else "output.docx"
        pdf_to_word(input_pdf, output_word)
        print(f"Conversion complete! Output saved to {output_word}")
