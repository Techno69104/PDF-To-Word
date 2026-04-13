from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import os
import tempfile
import uuid
import sys
import traceback

# Add current directory to path
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

# Import the enhanced OCR version
from main_ocr import pdf_to_word

app = Flask(__name__)
CORS(app)

# Configure upload limits (10MB)
app.config['MAX_CONTENT_LENGTH'] = 10 * 1024 * 1024

@app.route('/convert', methods=['POST'])
def convert_pdf_to_word():
    """Convert PDF to Word and return the DOCX file"""
    try:
        # Check if file was uploaded
        if 'file' not in request.files:
            return jsonify({'error': 'No file provided'}), 400
        
        file = request.files['file']
        
        if file.filename == '':
            return jsonify({'error': 'No file selected'}), 400
        
        if not file.filename.lower().endswith('.pdf'):
            return jsonify({'error': 'File must be PDF'}), 400
        
        print(f"Processing file: {file.filename}", file=sys.stderr)
        
        # Save uploaded PDF to temporary file
        with tempfile.NamedTemporaryFile(suffix='.pdf', delete=False) as temp_pdf:
            file.save(temp_pdf.name)
            pdf_path = temp_pdf.name
            print(f"Saved PDF to: {pdf_path}", file=sys.stderr)
        
        # Create output Word path
        output_dir = tempfile.mkdtemp()
        word_filename = f"{uuid.uuid4().hex}.docx"
        word_path = os.path.join(output_dir, word_filename)
        
        # Convert PDF to Word using OCR-enhanced function
        print(f"Starting conversion to: {word_path}", file=sys.stderr)
        result = pdf_to_word(pdf_path, word_path)
        
        # Check if conversion succeeded
        if not os.path.exists(word_path):
            return jsonify({'error': 'Conversion failed - output file not created'}), 500
        
        file_size = os.path.getsize(word_path)
        if file_size < 100:
            return jsonify({'error': 'Conversion produced an empty file'}), 500
        
        print(f"Conversion successful! File size: {file_size} bytes", file=sys.stderr)
        
        # Send the Word file back
        return send_file(
            word_path,
            as_attachment=True,
            download_name=f"{os.path.splitext(file.filename)[0]}.docx",
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )
    
    except Exception as e:
        error_msg = str(e)
        print(f"Error: {error_msg}", file=sys.stderr)
        traceback.print_exc(file=sys.stderr)
        return jsonify({'error': error_msg}), 500
    
    finally:
        # Clean up temporary files
        try:
            if 'pdf_path' in locals() and os.path.exists(pdf_path):
                os.unlink(pdf_path)
                print(f"Cleaned up: {pdf_path}", file=sys.stderr)
            if 'output_dir' in locals() and os.path.exists(output_dir):
                import shutil
                shutil.rmtree(output_dir, ignore_errors=True)
                print(f"Cleaned up: {output_dir}", file=sys.stderr)
        except Exception as cleanup_error:
            print(f"Cleanup error: {cleanup_error}", file=sys.stderr)

@app.route('/health', methods=['GET'])
def health_check():
    # Check if OCR is available
    ocr_status = "available" if __import__('main_ocr').OCR_AVAILABLE else "not_available"
    return jsonify({
        'status': 'healthy',
        'message': 'PDF to Word API with OCR is running',
        'ocr': ocr_status
    }), 200

@app.route('/', methods=['GET'])
def home():
    return jsonify({
        'service': 'PDF to Word Converter API',
        'version': '2.1',
        'features': ['OCR Support', 'Scanned PDF Support', 'Text Extraction'],
        'endpoints': {
            'convert': 'POST /convert (multipart/form-data with "file" field)',
            'health': 'GET /health'
        }
    }), 200

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
