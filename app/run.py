from flask import Flask, render_template, request, send_file, jsonify, after_this_request
import os
import docx
import json
from docx import Document
import uuid
import cloudinary
import cloudinary.uploader
import cloudinary.api
import requests
from io import BytesIO
from docx2pdf import convert
from PyPDF2 import PdfReader, PdfWriter
import pythoncom
from dotenv import load_dotenv
import tempfile

load_dotenv()

app = Flask(__name__)

# Configure Cloudinary with your credentials
cloudinary.config(
    cloud_name=os.getenv("CLOUD_NAME"),     
    api_key=os.getenv("API_KEY"),          
    api_secret=os.getenv("API_SECRET"),     
)


@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({'error': 'No file part in the request'}), 400

    file = request.files['file']

    if file.filename == '':
        return jsonify({'error': 'No file selected'}), 400

    # Upload the file to Cloudinary
    try:
        upload_result = cloudinary.uploader.upload(
            file,
            folder="uploads",  
            resource_type="auto"  
        )
        return jsonify({
            'message': 'File uploaded successfully',
            'file_url': upload_result['secure_url']  # Cloudinary secure URL
        }), 200
    except Exception as e:
        return jsonify({'error': f"An error occurred: {str(e)}"}), 500


def generate_unique_filename(extension):
    return f"{uuid.uuid4()}.{extension}"


def safe_convert(docx_path, pdf_path):
    """
    Safely convert a .docx file to PDF by ensuring COM is properly initialized.
    """
    pythoncom.CoInitialize()  # Initialize COM for the current thread
    try:
        convert(docx_path, pdf_path)  # Perform the conversion
    finally:
        pythoncom.CoUninitialize()  # Ensure COM is uninitialized


@app.route('/convert', methods=['POST'])
def convert_to_pdf():
    """
    Converts a .docx file to PDF and returns the file as a response.
    """
    data = request.json
    if not data or 'file_url' not in data:
        return jsonify({'error': 'No file_url provided'}), 400

    file_url = data['file_url']

    try:
        # Download the .docx file
        response = requests.get(file_url)
        if response.status_code != 200:
            return jsonify({'error': 'Failed to download the file'}), 400

        # Create a temporary file for the .docx file
        with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as temp_docx:
            temp_docx.write(response.content)
            docx_filename = temp_docx.name

        # Create a temporary file for the PDF file
        with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as temp_pdf:
            pdf_filename = temp_pdf.name

        # Convert the .docx file to PDF
        safe_convert(docx_filename, pdf_filename)

        # Verify if the PDF was created successfully
        if not os.path.exists(pdf_filename):
            return jsonify({'error': 'PDF conversion failed'}), 500

        # Schedule cleanup of temporary files after the response is sent
        @after_this_request
        def cleanup(response):
            try:
                os.remove(docx_filename)  # Remove the temporary .docx file
                os.remove(pdf_filename)  # Remove the temporary PDF file
            except Exception as e:
                print(f"Error cleaning up files: {str(e)}")
            return response

        # Send the PDF file to the client
        return send_file(pdf_filename, as_attachment=True)

    except Exception as e:
        print(e)
        return jsonify({'error': f"An error occurred during conversion: {str(e)}"}), 500


def add_password_to_pdf(input_pdf_path, output_pdf_path, password):
    """
    Adds a password to a PDF file using PyPDF2.
    """
    try:
        reader = PdfReader(input_pdf_path)
        writer = PdfWriter()

        # Copy pages from the original PDF
        for page in reader.pages:
            writer.add_page(page)

        # Encrypt the PDF with the given password
        writer.encrypt(password)

        # Save the protected PDF
        with open(output_pdf_path, 'wb') as output_file:
            writer.write(output_file)
    except Exception as e:
        raise Exception(f"Error adding password to PDF: {e}")

@app.route('/password-protection', methods=['POST'])
def password_protected():
    data = request.json
    if not data or 'file_url' not in data or 'password' not in data:
        return jsonify({'error': 'Missing file_url or password'}), 400

    file_url = data['file_url']
    password = data['password']

    try:
        # Download the file from the given URL
        response = requests.get(file_url)
        if response.status_code != 200:
            return jsonify({'error': 'Failed to download the file'}), 400

        # Create temporary files for the docx and PDF
        with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as temp_docx_file:
            temp_docx_file.write(response.content)
            docx_filename = temp_docx_file.name

        with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as temp_pdf_file:
            pdf_filename = temp_pdf_file.name

        # Convert the .docx to PDF
        safe_convert(docx_filename, pdf_filename)

        if not os.path.exists(pdf_filename):
            return jsonify({'error': 'PDF conversion failed'}), 500

        # Create a temporary file for the password-protected PDF
        with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as protected_temp_file:
            protected_pdf_filename = protected_temp_file.name

        # Add password protection to the PDF
        add_password_to_pdf(pdf_filename, protected_pdf_filename, password)

        # Read the password-protected PDF into memory
        with open(protected_pdf_filename, 'rb') as pdf_file:
            pdf_data = pdf_file.read()

        # Clean up temporary files
        os.remove(docx_filename)
        os.remove(pdf_filename)
        os.remove(protected_pdf_filename)

        # Send the password-protected PDF as a response
        return send_file(
            BytesIO(pdf_data),  # PDF content in memory
            mimetype='application/pdf',
            as_attachment=True,
            download_name="protected_converted.pdf"
        )

    except Exception as e:
        return jsonify({'error': f"An error occurred: {str(e)}"}), 500


def extract_metadata(docx_file):
    """
    Extract metadata from a .docx file and ensure all values are JSON-serializable.
    """
    from datetime import datetime

    doc = docx.Document(docx_file)  # Create a Document object from the Word document file.
    core_properties = doc.core_properties  # Get the core properties of the document.
    metadata = {}

    # Extract core properties
    for prop in dir(core_properties):
        if prop.startswith('__'):
            continue  # Skip built-in properties
        value = getattr(core_properties, prop)
        if callable(value):
            continue  # Skip methods

        # Convert datetime properties to strings
        if isinstance(value, datetime):
            value = value.strftime('%Y-%m-%d %H:%M:%S') if value else None

        # Convert non-serializable types to strings
        try:
            import json
            json.dumps(value)  # Test if the value is JSON-serializable
            metadata[prop] = value
        except (TypeError, ValueError):
            metadata[prop] = str(value)  # Fallback to string conversion

    return metadata

@app.route('/metadata', methods=['POST'])
def get_metadata():
    """
    Extract metadata from a .docx file provided as a file_url.
    """
    data = request.json

    # Validate the input
    if not data or 'file_url' not in data:
        return jsonify({'error': 'No file_url provided'}), 400

    file_url = data['file_url']

    try:
        # Download the .docx file from the file_url
        response = requests.get(file_url)
        if response.status_code != 200:
            return jsonify({'error': 'Failed to download the file from the provided URL'}), 400

        # Load the .docx file from the downloaded content
        docx_file = BytesIO(response.content)

        # Extract metadata
        metadata = extract_metadata(docx_file)

        # Return metadata as JSON response
        return jsonify({'metadata': metadata}), 200

    except Exception as e:
        return jsonify({'error': f'An error occurred while processing the file: {str(e)}'}), 500

if __name__ == "__main__":
    app.run(debug=True)