import os
import datetime
import traceback
from flask import Flask, request, send_file, jsonify, render_template

# Custom modules
from document_processor import DocumentProcessor
from utils.file_validator import allowed_file, get_safe_filename
from utils.batch_processor import BatchProcessor
from utils.db_manager import DatabaseManager # import db manager
from config import (
    UPLOAD_FOLDER, 
    OUTPUT_FOLDER, 
    SECRET_KEY, 
    MAX_FILE_SIZE,
    TEMPLATE_DOCX
)

# Initialize Flask App
app = Flask(__name__)
app.config['SECRET_KEY'] = SECRET_KEY
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['OUTPUT_FOLDER'] = OUTPUT_FOLDER
app.config['MAX_CONTENT_LENGTH'] = MAX_FILE_SIZE

# Initialize Global Processors
processor = DocumentProcessor(TEMPLATE_DOCX)
batch_processor = BatchProcessor(OUTPUT_FOLDER)

# Initialize Database
db = DatabaseManager() 

# ==========================================
# UI FRONTEND ROUTES
# ==========================================
@app.route('/', methods=['GET'])
def index():
    """Serves the main upload page."""
    return render_template('upload.html')

@app.route('/result/<filename>', methods=['GET'])
def show_result(filename):
    """Serves the success page after the API finishes."""
    count = request.args.get('count', 1, type=int)
    return render_template(
        'result.html',
        success=True,
        filename=filename,
        original_filename=filename if count == 1 else f"Batch Output ({count} files merged)",
        file_count=count
    )

@app.route('/api/v1/process', methods=['POST'])
def process_documents():
    """API Endpoint: Process text/files, merge them, and log to MySQL."""
    job_id = None
    try:
        text_input = request.form.get('text_input', '').strip()
        files = request.files.getlist('files[]')

        processing_queue = []
        temp_files = []

        has_files = any(f and f.filename != '' for f in files)

        if not text_input and not has_files:
            return jsonify({"success": False, "error": "Please provide text_input or files[]."}), 400

        # ── Handle Raw Text Input ──
        if text_input:
            from docx import Document
            doc = Document()
            for line in text_input.split('\n'):
                line = line.strip()
                if line: doc.add_paragraph(line)
                    
            timestamp = datetime.datetime.now().timestamp()
            text_docx_name = f"Text_Input_{timestamp}.docx"
            text_docx_path = os.path.join(app.config['UPLOAD_FOLDER'], text_docx_name)
            
            doc.save(text_docx_path)
            processing_queue.append(text_docx_path)
            temp_files.append(text_docx_path)

        # ── Handle File Uploads ──
        if has_files:
            for file in files:
                if file and file.filename != '':
                    if not allowed_file(file.filename):
                        return jsonify({"success": False, "error": f"File type not allowed: {file.filename}"}), 400
                    
                    filename = get_safe_filename(file.filename)
                    filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
                    file.save(filepath)
                    
                    processing_queue.append(filepath)
                    temp_files.append(filepath)

        if not processing_queue:
            return jsonify({"success": False, "error": "No valid content to process."}), 400

        original_names = [f.filename for f in files if f and f.filename != '']
        
        # If the user also typed text, add a label for it
        if text_input:
            original_names.insert(0, "Raw Text Input")

        # ==========================================
        # MYSQL LOGGING: START THE JOB
        # ==========================================
        file_count = len(processing_queue)
        
        # UPDATED: We now pass both the count AND the list of names
        job_id = db.create_job(file_count, original_names) 

        # ── Process and Merge Queue ──
        merge_result = batch_processor.process_and_merge(processing_queue, processor)

        # ── Cleanup Temporary Uploads ──
        for p in temp_files:
            try:
                if os.path.exists(p): os.remove(p)
            except Exception: pass

        # ── Return API Response & UPDATE MYSQL ──
        if merge_result.get('success'):
            output_filename = merge_result['output_file']
            
            # Log success to DB
            db.complete_job(job_id, output_filename)
            
            base_url = request.host_url.rstrip('/')
            
            return jsonify({
                "success": True,
                "message": "Processing complete.",
                "data": {
                    "job_id": job_id,
                    "filename": output_filename,
                    "file_count": file_count,
                    "download_url_docx": f"{base_url}/api/v1/download/{output_filename}",
                    "download_url_pdf": f"{base_url}/api/v1/download_pdf/{output_filename}"
                }
            }), 200
        else:
            # Log failure to DB
            db.fail_job(job_id, merge_result.get('error', 'Unknown merge error'))
            return jsonify({"success": False, "error": merge_result.get('error')}), 500

    except Exception as e:
        traceback.print_exc()
        # Log catastrophic failure to DB
        if job_id:
            db.fail_job(job_id, str(e))
        return jsonify({"success": False, "error": str(e)}), 500


@app.route('/api/v1/download/<filename>', methods=['GET'])
def download_file(filename):
    """API Endpoint: Download the converted DOCX file."""
    try:
        file_path = os.path.join(app.config['OUTPUT_FOLDER'], filename)
        if os.path.exists(file_path):
            return send_file(file_path, as_attachment=True, download_name=filename)
        else:
            return jsonify({"success": False, "error": "File not found."}), 404
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500


@app.route('/api/v1/download_pdf/<filename>', methods=['GET'])
def download_pdf(filename):
    """API Endpoint: Convert DOCX to PDF on the fly and download."""
    try:
        docx_path = os.path.join(app.config['OUTPUT_FOLDER'], filename)
        if not os.path.exists(docx_path):
            return jsonify({"success": False, "error": "Source DOCX document not found."}), 404
        
        pdf_filename = filename.replace('.docx', '.pdf')
        pdf_path = os.path.join(app.config['OUTPUT_FOLDER'], pdf_filename)
        
        try:
            from utils.pdf_converter import convert_to_pdf
            convert_to_pdf(docx_path, pdf_path)
            return send_file(pdf_path, as_attachment=True, download_name=pdf_filename)
        except Exception as e:
            traceback.print_exc()
            return jsonify({"success": False, "error": f"PDF conversion failed: {str(e)}"}), 500
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500


@app.errorhandler(413)
def request_entity_too_large(error):
    max_mb = app.config['MAX_CONTENT_LENGTH'] // (1024 * 1024)
    return jsonify({"success": False, "error": f"File too large. Maximum size is {max_mb} MB."}), 413


if __name__ == '__main__':
    print("Starting Document Processing API...")
    app.run(debug=True, host='0.0.0.0', port=5000)