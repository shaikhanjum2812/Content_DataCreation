import os
import logging
from flask import Flask, render_template, request, send_file, flash
from werkzeug.utils import secure_filename
from converter import convert_word_to_excel
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
    # Initialize variables for cleanup
    temp_input = None
    temp_output = None

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
        # Get original filename without extension
        original_name = os.path.splitext(secure_filename(file.filename))[0]

        # Create temporary files for processing
        temp_input = tempfile.NamedTemporaryFile(delete=False, suffix='.docx')
        file.save(temp_input.name)
        temp_output = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx')

        # Convert the file
        convert_word_to_excel(temp_input.name, temp_output.name)

        # Send the converted file
        return send_file(
            temp_output.name,
            as_attachment=True,
            download_name=f'{original_name}.xlsx',
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    except Exception as e:
        logging.error(f"Conversion error: {str(e)}")
        flash('Error converting file. Please try again.', 'error')
        return render_template('index.html')
    finally:
        # Safe cleanup of temporary files
        try:
            if temp_input and os.path.exists(temp_input.name):
                os.unlink(temp_input.name)
            if temp_output and os.path.exists(temp_output.name):
                os.unlink(temp_output.name)
        except Exception as e:
            logging.error(f"Error cleaning up temporary files: {str(e)}")