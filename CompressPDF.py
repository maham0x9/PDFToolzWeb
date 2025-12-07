import os
import io
import fitz  # PyMuPDF
import pdfplumber
import pandas as pd
import zipfile
from PIL import Image 
from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from pdf2docx import Converter
from pptx import Presentation
from pptx.util import Inches
from werkzeug.utils import secure_filename

app = Flask(__name__)
CORS(app)

# Configure folders
UPLOAD_FOLDER = 'uploads'
DOWNLOAD_FOLDER = 'downloads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(DOWNLOAD_FOLDER, exist_ok=True)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['DOWNLOAD_FOLDER'] = DOWNLOAD_FOLDER

# --- HELPER: Format File Size ---
def get_size_format(b, factor=1024, suffix="B"):
    """Converts bytes to KB, MB, etc."""
    for unit in ["", "K", "M", "G", "T"]:
        if b < factor:
            return f"{b:.2f}{unit}{suffix}"
        b /= factor
    return f"{b:.2f}Y{suffix}"

# --- HELPER: Intelligent Image Compressor ---
def compress_images_in_pdf(doc, quality=50, max_width=1024):
    """
    Iterates through the PDF, finds images, downscales/compresses them,
    and replaces the original streams.
    """
    # Track images we've already compressed to avoid duplicates
    img_xrefs = set()
    
    for page_num in range(len(doc)):
        page = doc[page_num]
        images = page.get_images()
        
        for img in images:
            xref = img[0]
            if xref in img_xrefs:
                continue
            img_xrefs.add(xref)

            # Extract image
            try:
                pix = fitz.Pixmap(doc, xref)
                
                # Skip small images (likely icons/logos)
                if pix.width < 100 or pix.height < 100:
                    continue

                # Handle CMYK/Alpha conversion
                if pix.n - pix.alpha > 3:
                    pix = fitz.Pixmap(fitz.csRGB, pix, 0)

                # Convert to PIL Image
                img_data = pix.tobytes()
                pil_img = Image.open(io.BytesIO(img_data))

                # Resize if too large (Downsampling)
                if pil_img.width > max_width:
                    ratio = max_width / float(pil_img.width)
                    new_height = int(float(pil_img.height) * ratio)
                    pil_img = pil_img.resize((max_width, new_height), Image.Resampling.LANCZOS)

                # Compress to JPEG
                buffer = io.BytesIO()
                pil_img.save(buffer, format="JPEG", quality=quality, optimize=True)
                new_img_data = buffer.getvalue()

                # Update the PDF stream
                doc.update_stream(xref, new_img_data)
                
            except Exception as e:
                print(f"Skipping image {xref}: {e}")
                continue

# --- ROUTES (Previous routes kept same, showing compress_pdf update) ---

@app.route('/convert-to-word', methods=['POST'])
def convert_to_word():
    if 'file' not in request.files: return jsonify({'error': 'No file'}), 400
    file = request.files['file']
    if file.filename == '': return jsonify({'error': 'No file'}), 400
    filename = secure_filename(file.filename)
    pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    file.save(pdf_path)
    word_filename = filename.rsplit('.', 1)[0] + '.docx'
    word_path = os.path.join(app.config['DOWNLOAD_FOLDER'], word_filename)
    try:
        cv = Converter(pdf_path)
        cv.convert(word_path, start=0, end=None)
        cv.close()
        return jsonify({'message': 'Success', 'download_url': f'/download/{word_filename}'})
    except Exception as e: return jsonify({'error': str(e)}), 500

@app.route('/convert-to-ppt', methods=['POST'])
def convert_to_ppt():
    if 'file' not in request.files: return jsonify({'error': 'No file'}), 400
    file = request.files['file']
    if file.filename == '': return jsonify({'error': 'No file'}), 400
    filename = secure_filename(file.filename)
    pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    file.save(pdf_path)
    ppt_filename = filename.rsplit('.', 1)[0] + '.pptx'
    ppt_path = os.path.join(app.config['DOWNLOAD_FOLDER'], ppt_filename)
    try:
        convert_pdf_to_pptx_logic(pdf_path, ppt_path)
        return jsonify({'message': 'Success', 'download_url': f'/download/{ppt_filename}'})
    except Exception as e: return jsonify({'error': str(e)}), 500

@app.route('/convert-to-excel', methods=['POST'])
def convert_to_excel():
    if 'file' not in request.files: return jsonify({'error': 'No file'}), 400
    file = request.files['file']
    if file.filename == '': return jsonify({'error': 'No file'}), 400
    filename = secure_filename(file.filename)
    pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    file.save(pdf_path)
    excel_filename = filename.rsplit('.', 1)[0] + '.xlsx'
    excel_path = os.path.join(app.config['DOWNLOAD_FOLDER'], excel_filename)
    try:
        convert_pdf_to_excel_logic(pdf_path, excel_path)
        return jsonify({'message': 'Success', 'download_url': f'/download/{excel_filename}'})
    except Exception as e: return jsonify({'error': str(e)}), 500

@app.route('/merge-pdfs', methods=['POST'])
def merge_pdfs():
    uploaded_files = request.files.getlist('files')
    if not uploaded_files: return jsonify({'error': 'No files'}), 400
    result_doc = fitz.open()
    for file in uploaded_files:
        file_stream = file.read()
        src_doc = fitz.open("pdf", file_stream)
        result_doc.insert_pdf(src_doc)
        src_doc.close()
    output_filename = "merged.pdf"
    output_path = os.path.join(app.config['DOWNLOAD_FOLDER'], output_filename)
    result_doc.save(output_path)
    result_doc.close()
    return jsonify({'message': 'Success', 'download_url': f'/download/{output_filename}'})

@app.route('/split-pdf', methods=['POST'])
def split_pdf():
    if 'file' not in request.files: return jsonify({'error': 'No file'}), 400
    file = request.files['file']
    filename = secure_filename(file.filename)
    pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    file.save(pdf_path)
    start_page = request.form.get('start_page', type=int)
    end_page = request.form.get('end_page', type=int)
    base_name = filename.rsplit('.', 1)[0]
    zip_filename = f"{base_name}_split.zip"
    zip_path = os.path.join(app.config['DOWNLOAD_FOLDER'], zip_filename)
    try:
        doc = fitz.open(pdf_path)
        total = len(doc)
        s = (start_page - 1) if start_page else 0
        e = end_page if end_page else total
        if s < 0: s = 0
        if e > total: e = total
        with zipfile.ZipFile(zip_path, 'w') as zipf:
            for i in range(s, e):
                new_doc = fitz.open()
                new_doc.insert_pdf(doc, from_page=i, to_page=i)
                zipf.writestr(f"page_{i+1}.pdf", new_doc.tobytes())
                new_doc.close()
        doc.close()
        return jsonify({'message': 'Success', 'download_url': f'/download/{zip_filename}'})
    except Exception as e: return jsonify({'error': str(e)}), 500

@app.route('/organize-pdf', methods=['POST'])
def organize_pdf():
    if 'file' not in request.files: return jsonify({'error': 'No file'}), 400
    file = request.files['file']
    filename = secure_filename(file.filename)
    pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    file.save(pdf_path)
    page_order = request.form.get('page_order', '')
    output_filename = f"organized_{filename}"
    output_path = os.path.join(app.config['DOWNLOAD_FOLDER'], output_filename)
    try:
        doc = fitz.open(pdf_path)
        indices = parse_page_string(page_order, len(doc))
        doc.select(indices)
        doc.save(output_path)
        doc.close()
        return jsonify({'message': 'Success', 'download_url': f'/download/{output_filename}'})
    except Exception as e: return jsonify({'error': str(e)}), 500

# --- IMPROVED COMPRESS PDF ROUTE ---
@app.route('/compress-pdf', methods=['POST'])
def compress_pdf():
    if 'file' not in request.files: return jsonify({'error': 'No file part'}), 400
    file = request.files['file']
    if file.filename == '': return jsonify({'error': 'No selected file'}), 400

    if file:
        filename = secure_filename(file.filename)
        pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(pdf_path)
        
        # 1. Get Original Size
        original_size = os.path.getsize(pdf_path)

        level = request.form.get('level', 'recommended') 
        base_name = filename.rsplit('.', 1)[0]
        output_filename = f"{base_name}_compressed.pdf"
        output_path = os.path.join(app.config['DOWNLOAD_FOLDER'], output_filename)

        try:
            doc = fitz.open(pdf_path)
            
            # LOGIC: 
            # If 'extreme': Downsample images aggressively + Max Garbage Collection
            # If 'recommended': Clean structure + Light image optimization
            # If 'less': Just clean structure
            
            if level == 'extreme':
                # Aggressively shrink images inside the PDF
                compress_images_in_pdf(doc, quality=30, max_width=800)
                doc.save(output_path, garbage=4, deflate=True, clean=True)
                
            elif level == 'recommended':
                # Moderate image shrink + cleanup
                compress_images_in_pdf(doc, quality=60, max_width=1600)
                doc.save(output_path, garbage=4, deflate=True)
                
            else: # 'less'
                # Just structural cleanup (lossless)
                doc.save(output_path, garbage=3, deflate=True)

            doc.close()

            # 2. Get New Size
            new_size = os.path.getsize(output_path)
            
            # 3. Safety check: If compressed is BIGGER, return original
            if new_size >= original_size:
                # Just copy original to output if compression failed to reduce
                doc = fitz.open(pdf_path)
                doc.save(output_path)
                doc.close()
                new_size = original_size # Reflect that size didn't change

            size_comparison = f"{get_size_format(original_size)} ➔ {get_size_format(new_size)}"

            return jsonify({
                'message': 'Compression successful',
                'download_url': f'/download/{output_filename}',
                'size_comparison': size_comparison
            })

        except Exception as e:
            print(e)
            return jsonify({'error': str(e)}), 500

# ... (Helpers like parse_page_string, convert_pdf_to_pptx_logic remain same) ...
# (Ensure you kept the imports at the top: fitz, PIL, io, os, etc.)

@app.route('/download/<filename>', methods=['GET'])
def download_file(filename):
    return send_file(os.path.join(app.config['DOWNLOAD_FOLDER'], filename), as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True, port=5000)