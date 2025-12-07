import os
from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from pdf2docx import Converter
from werkzeug.utils import secure_filename

app = Flask(__name__)
CORS(app)  # Enable Cross-Origin requests so your HTML can talk to Python

# Configure folders
UPLOAD_FOLDER = 'uploads'
DOWNLOAD_FOLDER = 'downloads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(DOWNLOAD_FOLDER, exist_ok=True)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['DOWNLOAD_FOLDER'] = DOWNLOAD_FOLDER

@app.route('/convert-to-word', methods=['POST'])
def convert_to_word():
    if 'file' not in request.files:
        return jsonify({'error': 'No file part'}), 400
    
    file = request.files['file']
    
    if file.filename == '':
        return jsonify({'error': 'No selected file'}), 400

    if file:
        # 1. Save the uploaded PDF
        filename = secure_filename(file.filename)
        pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(pdf_path)

        # 2. Define the Output Word filename
        word_filename = filename.rsplit('.', 1)[0] + '.docx'
        word_path = os.path.join(app.config['DOWNLOAD_FOLDER'], word_filename)

        try:
            # 3. Perform the Conversion
            cv = Converter(pdf_path)
            cv.convert(word_path, start=0, end=None)
            cv.close()

            # 4. Return the download URL
            return jsonify({
                'message': 'Conversion successful',
                'download_url': f'/download/{word_filename}'
            })
        except Exception as e:
            return jsonify({'error': str(e)}), 500

@app.route('/download/<filename>', methods=['GET'])
def download_file(filename):
    # Serve the file to the user
    return send_file(os.path.join(app.config['DOWNLOAD_FOLDER'], filename), as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True, port=5000)