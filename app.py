import os
import time
import uuid
import threading
try:
    import pythoncom
    import win32com.client
    HAS_WIN32COM = True
except ImportError:
    HAS_WIN32COM = False
from flask import Flask, request, jsonify, send_from_directory
from werkzeug.utils import secure_filename
from pdf2docx import Converter

app = Flask(__name__, static_folder='static', static_url_path='/static')

# Configuration
UPLOAD_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'uploads')
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 25 * 1024 * 1024  # 25MB Limit
ALLOWED_EXTENSIONS = {'pdf', 'docx', 'doc'}

os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# Background file cleanup loop (runs every 10 minutes, deletes files > 1 hour old)
def cleanup_old_files(folder_path, max_age_seconds=3600):
    while True:
        try:
            now = time.time()
            for filename in os.listdir(folder_path):
                file_path = os.path.join(folder_path, filename)
                if os.path.isfile(file_path):
                    # Delete files older than 1 hour
                    if now - os.path.getmtime(file_path) > max_age_seconds:
                        try:
                            os.remove(file_path)
                            print(f"Cleanup: Deleted expired file {filename}")
                        except Exception as e:
                            print(f"Cleanup: Error deleting {filename}: {e}")
        except Exception as e:
            print(f"Cleanup: Error running cleanup loop: {e}")
        time.sleep(600)  # Sleep 10 minutes

# Start the cleanup thread as a daemon
cleanup_thread = threading.Thread(target=cleanup_old_files, args=(UPLOAD_FOLDER,), daemon=True)
cleanup_thread.start()

def ensure_word_registry_settings():
    if os.name != 'nt':
        return
    try:
        import winreg
        registry_path = r"Software\Microsoft\Office\16.0\Word\Options"
        key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, registry_path)
        winreg.SetValueEx(key, "DidShowPDFConversionWarning", 0, winreg.REG_DWORD, 1)
        winreg.SetValueEx(key, "DisableConvertPdfWarning", 0, winreg.REG_DWORD, 1)
        winreg.CloseKey(key)
    except Exception as e:
        print(f"Error setting registry options: {e}")

ensure_word_registry_settings()

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def convert_pdf_to_docx(pdf_path, docx_path):
    """Converts a PDF file to DOCX using Word COM on Windows or pdf2docx on Linux."""
    if HAS_WIN32COM and os.name == 'nt':
        pythoncom.CoInitialize()
        word = None
        doc = None
        try:
            import win32com.client.gencache
            word = win32com.client.gencache.EnsureDispatch("Word.Application")
            word.Visible = False
            word.DisplayAlerts = 0  # wdAlertsNone
            doc = word.Documents.Open(os.path.abspath(pdf_path), False, True)
            doc.SaveAs(os.path.abspath(docx_path), 16)  # 16 = wdFormatXMLDocument (docx)
        finally:
            if doc:
                try: doc.Close(SaveChanges=0)
                except Exception: pass
            if word:
                try: word.Quit()
                except Exception: pass
            pythoncom.CoUninitialize()
    else:
        # Fallback to python-based converter on Linux
        cv = Converter(pdf_path)
        cv.convert(docx_path, start=0, end=None)
        cv.close()

def convert_docx_to_pdf_libreoffice(docx_path, pdf_path):
    """Converts a DOCX file to PDF using LibreOffice headless command line on Linux."""
    import subprocess
    import shutil
    
    soffice_path = shutil.which('libreoffice') or shutil.which('soffice')
    if not soffice_path:
        for path in ['/usr/bin/libreoffice', '/usr/bin/soffice', '/usr/lib/libreoffice/program/soffice']:
            if os.path.exists(path):
                soffice_path = path
                break
                
    if not soffice_path:
        raise RuntimeError("LibreOffice/soffice executable not found in PATH or standard directories.")
        
    out_dir = os.path.dirname(os.path.abspath(pdf_path))
    cmd = [
        soffice_path,
        '--headless',
        '--convert-to', 'pdf',
        '--outdir', out_dir,
        os.path.abspath(docx_path)
    ]
    
    print(f"Running LibreOffice conversion: {' '.join(cmd)}")
    result = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
    if result.returncode != 0:
        raise RuntimeError(f"LibreOffice conversion failed: {result.stderr}")

def convert_docx_to_pdf(docx_path, pdf_path):
    """Converts a DOCX/DOC file to PDF using MS Word on Windows or LibreOffice on Linux."""
    if HAS_WIN32COM and os.name == 'nt':
        pythoncom.CoInitialize()
        word = None
        doc = None
        try:
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False
            doc = word.Documents.Open(os.path.abspath(docx_path))
            doc.SaveAs(os.path.abspath(pdf_path), FileFormat=17)
        finally:
            if doc:
                try: doc.Close(SaveChanges=0)
                except Exception: pass
            if word:
                try: word.Quit()
                except Exception: pass
            pythoncom.CoUninitialize()
    else:
        # Fallback to LibreOffice headless conversion on Linux
        convert_docx_to_pdf_libreoffice(docx_path, pdf_path)

@app.route('/')
def index():
    return send_from_directory('static', 'index.html')

@app.route('/api/convert', methods=['POST'])
def convert_file():
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded.'}), 400
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No file selected.'}), 400
    
    if not allowed_file(file.filename):
        return jsonify({'error': 'Unsupported file format. Please upload PDF, DOCX, or DOC.'}), 400
    
    # Save the input file with a unique ID to avoid name collisions
    file_id = str(uuid.uuid4())
    orig_filename = secure_filename(file.filename)
    base_name, ext = os.path.splitext(orig_filename.lower())
    
    input_filename = f"{file_id}{ext}"
    input_path = os.path.join(app.config['UPLOAD_FOLDER'], input_filename)
    file.save(input_path)
    
    try:
        if ext == '.pdf':
            output_filename = f"{file_id}.docx"
            output_path = os.path.join(app.config['UPLOAD_FOLDER'], output_filename)
            convert_pdf_to_docx(input_path, output_path)
            display_name = f"{base_name}.docx"
        else:
            output_filename = f"{file_id}.pdf"
            output_path = os.path.join(app.config['UPLOAD_FOLDER'], output_filename)
            convert_docx_to_pdf(input_path, output_path)
            display_name = f"{base_name}.pdf"
            
        import urllib.parse
        encoded_name = urllib.parse.quote(display_name)
        return jsonify({
            'success': True,
            'download_url': f'/api/download/{output_filename}?name={encoded_name}',
            'filename': display_name
        })
        
    except Exception as e:
        print(f"Error during conversion of {orig_filename}: {e}")
        return jsonify({'error': f'Conversion failed: {str(e)}'}), 500

@app.route('/api/download/<filename>', methods=['GET'])
def download_file(filename):
    # Ensure client cannot escape the uploads directory
    safe_filename = secure_filename(filename)
    file_path = os.path.join(app.config['UPLOAD_FOLDER'], safe_filename)
    if not os.path.exists(file_path):
        return jsonify({'error': 'File not found or has expired.'}), 404
        
    # Get the original filename from request arguments to return a user-friendly name
    display_name = request.args.get('name', safe_filename)
    safe_display_name = secure_filename(display_name)
    
    return send_from_directory(app.config['UPLOAD_FOLDER'], safe_filename, as_attachment=True, download_name=safe_display_name)

@app.errorhandler(413)
def request_entity_too_large(error):
    return jsonify({'error': 'File is too large. Sane limit is 25MB.'}), 413

if __name__ == '__main__':
    app.run(debug=True, port=5000)
