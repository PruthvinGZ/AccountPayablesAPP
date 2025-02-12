from flask import Flask, render_template, request, send_from_directory
import os
import logging
import subprocess
import sys
import glob
import socket

# Create the Flask app instance once
app = Flask(__name__, template_folder='templates', static_folder='static')

# Configuration constants
UPLOAD_FOLDER = 'uploads'
PROCESSED_FOLDER = 'processed'
LOG_FOLDER = 'logs'
SCRIPT_PATH = 'Payable_Account_Automation.py'

# Configure app
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['PROCESSED_FOLDER'] = PROCESSED_FOLDER
app.config['LOG_FOLDER'] = LOG_FOLDER

# Ensure directories exist
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(PROCESSED_FOLDER, exist_ok=True)
os.makedirs(LOG_FOLDER, exist_ok=True)

# Configure logging
logging.basicConfig(
    filename=os.path.join(LOG_FOLDER, 'app.log'),
    level=logging.ERROR,
    format='%(asctime)s - %(message)s'
)

def is_port_available(port):
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        try:
            s.bind(('127.0.0.1', port))
            return True
        except OSError:
            return False

def find_available_port(start_port=5000, max_port=5020):
    for port in range(start_port, max_port):
        if is_port_available(port):
            return port
    return None

# Routes
@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    try:
        if 'file' not in request.files:
            return 'No file part', 400

        file = request.files['file']
        file_type = request.form['file_type']

        if file.filename == '':
            return 'No selected file', 400

        filepath = os.path.join(app.config['UPLOAD_FOLDER'], f"{file_type}.xlsx")
        file.save(filepath)
        logging.info(f"Uploaded file: {file.filename} as {file_type}.xlsx")

        return f"File uploaded successfully: {filepath}", 200
    except Exception as e:
        logging.error(f"Error uploading file: {str(e)}")
        return f"Error uploading file: {str(e)}", 500

@app.route('/process', methods=['POST'])
def process_files():
    try:
        required_files = [
            os.path.join(app.config['UPLOAD_FOLDER'], "account_payables.xlsx"),
            os.path.join(app.config['UPLOAD_FOLDER'], "bank_balance.xlsx"),
            os.path.join(app.config['UPLOAD_FOLDER'], "cash_management.xlsx")
        ]

        if not all(os.path.exists(file) for file in required_files):
            return 'Missing one or more required files.', 400

        result = subprocess.run(
            [sys.executable, SCRIPT_PATH],
            capture_output=True,
            text=True
        )

        if result.returncode != 0:
            logging.error(f"Script error: {result.stderr}")
            return f'Processing error: {result.stderr}', 500

        logging.info("Processing completed successfully.")
        return 'Processing completed, final report ready.', 200
    except Exception as e:
        logging.error(f"Processing failed: {str(e)}")
        return f'Processing failed: {str(e)}', 500

@app.route('/download/final_report')
def download_final_report():
    report_files = glob.glob(os.path.join(app.config['PROCESSED_FOLDER'], "Payables_Summary_*.xlsx"))
    
    if not report_files:
        return 'Final report not found.', 404

    latest_report = max(report_files, key=os.path.getctime)
    return send_from_directory(
        app.config['PROCESSED_FOLDER'], 
        os.path.basename(latest_report), 
        as_attachment=True
    )

@app.route('/open-folder')
def open_folder():
    try:
        folder_path = os.path.abspath(app.config['PROCESSED_FOLDER'])
        if sys.platform == 'win32':
            os.startfile(folder_path)
        return 'Folder opened successfully', 200
    except Exception as e:
        logging.error(f"Folder error: {str(e)}")
        return f'Folder error: {str(e)}', 500

@app.route('/health')
def health_check():
    return "Application is running", 200

# Main server startup
if __name__ == '__main__':
    port = find_available_port()
    if not port:
        logging.error("No available ports in range 5000-5020")
        sys.exit(1)
    
    try:
        with open(os.path.join(LOG_FOLDER, 'server_port.txt'), 'w') as f:
            f.write(str(port))
    except Exception as e:
        logging.error(f"Port file error: {str(e)}")
        sys.exit(1)
    
    try:
        print(f"Starting server on port {port}")
        app.run(host='127.0.0.1', port=port, debug=False)
    except Exception as e:
        logging.error(f"Server start failed: {str(e)}")
        sys.exit(1)