import os
import logging
from flask import Flask, render_template, request, send_file, flash
from werkzeug.utils import secure_filename
from converter import convert_word_to_excel, create_text_files
import tempfile

# Configure logging
logging.basicConfig(level=logging.DEBUG)

app = Flask(__name__)
app.secret_key = os.environ.get("SESSION_SECRET", "default-secret-key")

# Configure upload settings
ALLOWED_EXTENSIONS = {'docx'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        flash('No file part', 'error')
        return render_template('index.html')

    file = request.files['file']
    mode = request.form.get('mode', 'reader')

    if file.filename == '':
        flash('No file selected', 'error')
        return render_template('index.html')

    if not allowed_file(file.filename):
        flash('Invalid file type. Please upload a .docx file', 'error')
        return render_template('index.html')

    try:
        # Create temporary files for processing
        with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as temp_input:
            file.save(temp_input.name)
            temp_output = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx')

            # Convert the file using the selected mode
            convert_word_to_excel(temp_input.name, temp_output.name, mode)

            # Send the converted file
            return send_file(
                temp_output.name,
                as_attachment=True,
                download_name=f'converted_{mode}.xlsx',
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
    except Exception as e:
        logging.error(f"Conversion error: {str(e)}")
        flash('Error converting file. Please try again.', 'error')
        return render_template('index.html')
    finally:
        # Cleanup temporary files
        try:
            os.unlink(temp_input.name)
            os.unlink(temp_output.name)
        except:
            pass

@app.route('/generate-text-files', methods=['POST'])
def generate_text_files():
    if 'file' not in request.files:
        flash('No file part', 'error')
        return render_template('index.html')

    file = request.files['file']

    if file.filename == '':
        flash('No file selected', 'error')
        return render_template('index.html')

    if not allowed_file(file.filename):
        flash('Invalid file type. Please upload a .docx file', 'error')
        return render_template('index.html')

    try:
        # Create temporary file for processing
        with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as temp_input:
            file.save(temp_input.name)

            # Generate text files and get the zip file path
            zip_path = create_text_files(temp_input.name)

            # Send the zip file
            return send_file(
                zip_path,
                as_attachment=True,
                download_name='code_files.zip',
                mimetype='application/zip'
            )
    except Exception as e:
        logging.error(f"Text file generation error: {str(e)}")
        flash('Error generating text files. Please try again.', 'error')
        return render_template('index.html')
    finally:
        # Cleanup temporary files
        try:
            os.unlink(temp_input.name)
            os.unlink(zip_path)
        except:
            pass
