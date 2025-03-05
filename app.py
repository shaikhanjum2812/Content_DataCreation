import os
import logging
from flask import Flask, render_template, request, send_file, flash, redirect, url_for
from werkzeug.utils import secure_filename
from converter import convert_word_to_excel, create_text_files
import tempfile

# Configure logging
logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)

# Initialize Flask app
app = Flask(__name__, template_folder='templates', static_folder='static')
app.secret_key = os.environ.get("SESSION_SECRET", "default-secret-key")

# Configure upload settings
ALLOWED_EXTENSIONS = {'docx'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def index():
    logger.debug("Accessing index route")
    try:
        return render_template('index.html')
    except Exception as e:
        logger.error(f"Error rendering template: {str(e)}", exc_info=True)
        return "Error loading page", 500

@app.route('/upload', methods=['POST'])
def upload_file():
    temp_input = None
    temp_output = None

    try:
        if 'file' not in request.files:
            flash('No file part', 'error')
            return redirect(url_for('index'))

        file = request.files['file']
        if file.filename == '':
            flash('No selected file', 'error')
            return redirect(url_for('index'))

        if not allowed_file(file.filename):
            flash('Invalid file type. Please upload a .docx file', 'error')
            return redirect(url_for('index'))

        # Create temporary files for processing
        temp_input = tempfile.NamedTemporaryFile(delete=False, suffix='.docx')
        file.save(temp_input.name)
        temp_output = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx')

        # Convert the file
        logger.info(f"Converting file: {file.filename}")
        success = convert_word_to_excel(temp_input.name, temp_output.name)

        if not success:
            flash('Error converting file. Please check the document format.', 'error')
            return redirect(url_for('index'))

        # Send the converted file
        return send_file(
            temp_output.name,
            as_attachment=True,
            download_name=f"{os.path.splitext(secure_filename(file.filename))[0]}.xlsx",
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

    except Exception as e:
        logger.error(f"Conversion error: {str(e)}", exc_info=True)
        flash('Error converting file. Please try again.', 'error')
        return redirect(url_for('index'))

    finally:
        # Cleanup temporary files
        try:
            if temp_input and os.path.exists(temp_input.name):
                os.unlink(temp_input.name)
            if temp_output and os.path.exists(temp_output.name):
                os.unlink(temp_output.name)
        except Exception as e:
            logger.error(f"Error cleaning up temporary files: {str(e)}")

@app.route('/extract-code', methods=['POST'])
def extract_code_files():
    temp_input = None
    zip_path = None

    try:
        if 'file' not in request.files:
            flash('No file part', 'error')
            return redirect(url_for('index'))

        file = request.files['file']
        if file.filename == '':
            flash('No selected file', 'error')
            return redirect(url_for('index'))

        if not allowed_file(file.filename):
            flash('Invalid file type. Please upload a .docx file', 'error')
            return redirect(url_for('index'))

        # Create temporary file for processing
        temp_input = tempfile.NamedTemporaryFile(delete=False, suffix='.docx')
        file.save(temp_input.name)

        # Extract code files
        zip_path = create_text_files(temp_input.name)

        # Send the zip file
        return send_file(
            zip_path,
            as_attachment=True,
            download_name=f"{os.path.splitext(secure_filename(file.filename))[0]}_code_files.zip",
            mimetype='application/zip'
        )

    except Exception as e:
        logger.error(f"Code extraction error: {str(e)}", exc_info=True)
        flash('Error extracting code files. Please try again.', 'error')
        return redirect(url_for('index'))

    finally:
        # Cleanup temporary files
        try:
            if temp_input and os.path.exists(temp_input.name):
                os.unlink(temp_input.name)
            if zip_path and os.path.exists(zip_path):
                os.unlink(zip_path)
        except Exception as e:
            logger.error(f"Error cleaning up temporary files: {str(e)}")

if __name__ == "__main__":
    app.run(host='0.0.0.0', port=5000, debug=True)