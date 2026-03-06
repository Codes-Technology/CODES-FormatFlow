"""
app.py 
"""

import os
import datetime
import traceback
from flask import Flask, request, send_file, jsonify, render_template, url_for

from document_processor import DocumentProcessor
from utils.file_processor import FileProcessor
from utils.db_manager import DatabaseManager
from config import UPLOAD_FOLDER, OUTPUT_FOLDER, SECRET_KEY, MAX_FILE_SIZE, TEMPLATE_DOCX

app = Flask(__name__)
app.config.update(
    SECRET_KEY=SECRET_KEY,
    UPLOAD_FOLDER=UPLOAD_FOLDER,
    OUTPUT_FOLDER=OUTPUT_FOLDER,
    MAX_CONTENT_LENGTH=MAX_FILE_SIZE
)

processor       = DocumentProcessor(TEMPLATE_DOCX)
file_processor = FileProcessor(OUTPUT_FOLDER)
db              = DatabaseManager()


@app.route('/', methods=['GET'])
def index():
    return render_template('upload.html')


@app.route('/api/v1/process', methods=['POST'])
def process_documents():
    job_id = None
    temp_files = []
    try:
        text_input    = request.form.get('text_input', '').strip()
        files         = request.files.getlist('files[]')
        include_cover = request.form.get('include_cover') == 'true'
        include_toc   = request.form.get('include_toc')   == 'true'

        processing_queue = []
        has_files        = any(f and f.filename for f in files)

        if not text_input and not has_files:
            return jsonify({'success': False, 'error': 'No input provided.'}), 400

        # Rich-text editor
        if text_input:
            doc_obj = processor.html_to_docx(text_input)   # public method
            name = f"editor_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
            path = os.path.join(app.config['UPLOAD_FOLDER'], name)
            doc_obj.save(path)
            processing_queue.append(path)
            temp_files.append(path)

        # File uploads
        for file in files:
            if file and file.filename:
                safe = os.path.basename(file.filename)
                in_path  = os.path.join(app.config['UPLOAD_FOLDER'], safe)
                out_name = f'proc_{safe}.docx'
                out_path = os.path.join(app.config['UPLOAD_FOLDER'], out_name)
                file.save(in_path)

                res = processor.universal_extract(in_path, out_path)
                if res['success']:
                    processing_queue.append(out_path)
                    temp_files.extend([in_path, out_path])
                else:
                    print(f'[App] Extraction failed for {safe}: {res.get("error")}')

        if not processing_queue:
            return jsonify({'success': False, 'error': 'No files processed.'}), 500

        original_names = [f.filename for f in files if f and f.filename]
        if text_input:
            original_names.insert(0, 'Rich Text Input')
        job_id = db.create_job(len(processing_queue), original_names)

        ts = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
        final_name = f'Document_{ts}.docx'
        final_path = os.path.join(app.config['OUTPUT_FOLDER'], final_name)

        ok = file_processor.process_file(
            processing_queue, final_path,
            include_cover=include_cover,
            include_toc=include_toc
        )

        for p in temp_files:
            try:
                if os.path.exists(p): os.remove(p)
            except OSError:
                pass

        if ok:
            db.complete_job(job_id, final_name)
            return jsonify({'success': True, 'data': {
                'filename': final_name,
                'redirect_url': url_for('result_page', filename=final_name)}})
        else:
            db.fail_job(job_id, 'Merge failed.')
            return jsonify({'success': False, 'error': 'Merge failed.'}), 500

    except Exception as e:
        print(f'[App] Critical: {e}')
        traceback.print_exc()
        if job_id:
            db.fail_job(job_id, str(e))
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/result/<filename>')
def result_page(filename):
    return render_template('result.html',
        success=True,
        filename=filename )


@app.route('/download/<filename>')
def download_file(filename):
    path = os.path.join(app.config['OUTPUT_FOLDER'], os.path.basename(filename))
    if os.path.exists(path):
        return send_file(path, as_attachment=True)
    return 'File not found', 404


@app.route('/download-pdf/<filename>')
def download_pdf(filename):
    """Windows-only PDF conversion. COM libraries imported here, not at module level."""
    try:
        import pythoncom
        from docx2pdf import convert
    except ImportError:
        return 'PDF conversion requires Windows + Microsoft Word.', 501

    pythoncom.CoInitialize()
    try:
        docx_path    = os.path.join(app.config['OUTPUT_FOLDER'], os.path.basename(filename))
        pdf_filename = filename.replace('.docx', '.pdf')
        pdf_path     = os.path.join(app.config['OUTPUT_FOLDER'], pdf_filename)
        if not os.path.exists(docx_path):
            return f'File not found: {filename}', 404
        convert(docx_path, pdf_path)
        return send_file(pdf_path, as_attachment=True)
    except Exception as e:
        traceback.print_exc()
        return f'PDF error: {e}', 500
    finally:
        try: pythoncom.CoUninitialize()
        except Exception: pass


if __name__ == '__main__':
    os.makedirs(UPLOAD_FOLDER, exist_ok=True)
    os.makedirs(OUTPUT_FOLDER, exist_ok=True)
    app.run(debug=True, host='0.0.0.0', port=5000)