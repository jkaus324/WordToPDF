
# Word to PDF Converter

A web-based application that allows users to upload Word files (.docx), convert them to PDF, add password protection, and extract metadata from PDFs. This project is built using Flask (backend) and a simple HTML/JavaScript frontend with Bootstrap.



## Features

- Upload .docx files to the server.
- Convert Word files to PDF.
- Add password protection to the generated PDF.
- Extract metadata from the uploaded or converted PDF.
- Simple and intuitive UI built with Bootstrap.
- Axios for handling API requests to the backend.



## Tech Stack

### *Frontend*
- HTML5, CSS3
- Bootstrap 5
- JavaScript
- Axios for API requests

### *Backend*
- Flask (Python)
- Flask-CORS for Cross-Origin Resource Sharing
- Cloudinary (for file management, if used)
- Docker for containerization

## Run Locally

### *1. Clone the Repository*
bash
git clone https://github.com/yourusername/WordToPDF.git
cd WordToPDF


### *2. Create a Virtual Environment*
Set up a Python virtual environment and activate it:
bash
python -m venv venv
# On Windows
venv\Scripts\activate
# On macOS/Linux
source venv/bin/activate


### *3. Install Dependencies*
Install the required Python libraries:
bash
pip install -r requirements.txt


### *4. Set Up Environment Variables*
Create a .env file in the project root and configure your environment variables:
env
CLOUD_NAME=your-cloud-name
API_KEY=your-api-key
API_SECRET=your-api-secret


### *5. Run the Application*
Start the Flask development server:
bash
python ./app/run.py


The application will be accessible at http://localhost:5000.
### *1. Clone the Repository*
bash
git clone https://github.com/yourusername/WordToPDF.git
cd WordToPDF


### *2. Create a Virtual Environment*
Set up a Python virtual environment and activate it:
bash
python -m venv venv
# On Windows
venv\Scripts\activate
# On macOS/Linux
source venv/bin/activate


### *3. Install Dependencies*
Install the required Python libraries:
bash
pip install -r requirements.txt


### *4. Set Up Environment Variables*
Create a .env file in the project root and configure your environment variables:
env
CLOUD_NAME=your-cloud-name
API_KEY=your-api-key
API_SECRET=your-api-secret


### *5. Run the Application*
Start the Flask development server:
bash
python ./app/run.py


The application will be accessible at http://localhost:5000.

## *Docker Instructions*

### *1. Build the Docker Image*
bash
docker build -t wordtopdf .


### *2. Run the Docker Container*
bash
docker run -d -p 5000:5000 --env-file .env --name wordtopdf-container wordtopdf


### *3. Access the Application*
Go to [http://localhost:5000](http://localhost:5000) in your browser.
## *Usage*

### *1. Upload a Word File*
1. Click the "Upload File" button.
2. Select a .docx file to upload.
3. Once the file is uploaded, you’ll receive a status message.

### *2. Convert Word to PDF*
1. After uploading, click the "Convert to PDF" button.
2. The converted PDF will be downloaded automatically.

### *3. Add Password to PDF*
1. Enter a password in the input field.
2. Click the "Add Password" button.
3. A password-protected PDF will be downloaded.

### *4. Extract Metadata*
1. Click the "Get Metadata" button.
2. Metadata details will be displayed on the screen.
## *Project Structure*

WordToPDF/
│
├── app/
│   ├── static/          # Static files (CSS, JS, images, etc.)
│   ├── templates/       # HTML templates
│   ├── run.py           # Main Flask application
│   └── ...              # Other backend modules
│
├── venv/                # Virtual environment (not included in Git)
├── Dockerfile           # Docker configuration
├── .env                 # Environment variables (ignored in Git)
├── requirements.txt     # Python dependencies
├── README.md            # Project documentation
└── ...

## *Endpoints*

### **1. /upload**
- *Method*: POST
- *Description*: Uploads a .docx file to the server.
- *Request Body*: Multipart/form-data with a file key.
- *Response*:
  json
  {
      "message": "File uploaded successfully",
      "file_url": "https://cloudinary.com/example.pdf"
  }
  

### **2. /convert**
- *Method*: POST
- *Description*: Converts a Word file to a PDF.
- *Request Body*: JSON with the file_url of the uploaded file.
- *Response*: Downloads the converted PDF.

### **3. /password-protection**
- *Method*: POST
- *Description*: Adds a password to the converted PDF.
- *Request Body*: JSON with file_url and password.
- *Response*: Downloads the password-protected PDF.

### **4. /metadata**
- *Method*: POST
- *Description*: Extracts metadata from the uploaded/converted PDF.
- *Request Body*: JSON with the file_url.
- *Response*:
  json
  {
      "metadata": {
          "title": "Example PDF",
          "author": "John Doe",
          ...
      }
  }

## Environment Variables

To run this project, you will need to add the following environment variables to your .env file

`CLOUD_NAME`

`API_KEY`

`API_SECRET`

