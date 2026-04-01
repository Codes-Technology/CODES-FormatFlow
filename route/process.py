import os
import re
import tempfile
import time
import requests
from datetime import datetime
from werkzeug.utils import secure_filename
from flask import Blueprint, request, jsonify, send_file
from utils.decorators import require_auth
from utils.db_manager import db, ProcessingJob, ProcessingJobHistory, JobType, JobStatus
from document_processor import DocumentProcessor
from config import TEMPLATE_DOCX
from io import BytesIO
import pythoncom
from docx2pdf import convert
# Helper Functions
def clean_html(text):
    """Remove HTML tags from a string."""
    if not text: return ""
    return re.sub(r'<[^>]+>', '', text)

def normalize_filename(s):
    """Sanitize a string for use as a filename."""
    if not s: return "document"
    # Remove special chars, replace spaces with underscores
    s = re.sub(r'[^\w\s-]', '', s).strip()
    return re.sub(r'[-\s]+', '_', s)
import logging

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

process_bp = Blueprint('process', __name__)


# ==============================
# SMART TITLE GENERATOR
# ==============================
def generate_title(text):
    """
    Generate a smart summary title using Groq API.
    Fallback to first sentence if API fails.
    """
    try:
        api_key = os.getenv("GROQ_API")
        if not api_key:
            logger.warning("GROQ_API key not found, skipping smart title")
            raise ValueError("Missing API Key")

        clean_prompt = clean_html(text)[:1000]
        if not clean_prompt:
            return "New Chat"

        response = requests.post(
            "https://api.groq.com/openai/v1/chat/completions",
            headers={"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"},
            json={
                "model": "llama-3.3-70b-versatile",
                "messages": [
                    {"role": "system", "content": "Generate a catchy 3-5 word title for the following text. Respond ONLY with the plain text title, no punctuation or quotes."},
                    {"role": "user", "content": clean_prompt}
                ],
                "max_tokens": 20
            },
            timeout=5
        )

        if response.status_code == 200:
            title = response.json().get('choices', [{}])[0].get('message', {}).get('content', '').strip()
            if title:
                return normalize_filename(title).replace("_", " ")[:60]

    except Exception as e:
        logger.error(f"Groq Title API failed: {e}")

    return get_first_sentence(text)


# ==============================
# UNIQUE FILENAME GENERATOR
# ==============================
def generate_unique_filename(base_name, current_user):
    clean_name = normalize_filename(base_name)
    timestamp = datetime.now().strftime('%Y-%m-%d_%H-%M')
    filename_base = f"{clean_name}_{timestamp}"
    filename = f"{filename_base}.docx"

    counter = 1
    while ProcessingJobHistory.query.join(ProcessingJob).filter(
        ProcessingJob.UserId == current_user.Id,
        ProcessingJobHistory.OutputFileName == filename
    ).first():
        filename = f"{filename_base}_{counter}.docx"
        counter += 1

    return filename


# ==============================
# PROCESS TEXT
# ==============================
@process_bp.route('/process-text', methods=['POST'])
@require_auth
def process_text(current_user):
    try:
        start_time = time.time()
        data = request.get_json()
        text_input = data.get('text', '')

        user_font = data.get('fontFamily', 'Calibri')
        user_size = int(data.get('fontSize', 11))
        include_cover = data.get('includeCover', False)
        include_toc = data.get('includeTOC', False)

        processor = DocumentProcessor(
            template_path=TEMPLATE_DOCX,
            font_family=user_font,
            font_size=user_size,
            include_cover=include_cover,
            include_toc=include_toc
        )

        doc_obj = processor.html_to_docx(text_input)

        output_buffer = BytesIO()
        doc_obj.save(output_buffer)
        output_buffer.seek(0)
        final_file_data = output_buffer.read()

        full_text_summary = clean_html(text_input)
        job_id = data.get('jobId', None)

        # Job handling
        if job_id:
            job = ProcessingJob.query.filter_by(Id=job_id, UserId=current_user.Id).first()
            if not job:
                return jsonify({'success': False, 'error': 'Conversation not found'}), 404
        else:
            job_name = generate_title(text_input)
            job = ProcessingJob(JobName=job_name, UserId=current_user.Id)
            db.session.add(job)
            db.session.flush()

        filename = generate_unique_filename(job.JobName, current_user)

        history = ProcessingJobHistory(
            ProcessJobId=job.Id,
            JobType=JobType.TEXT,
            Summary=full_text_summary,
            UploadFileData=text_input.encode('utf-8'),
            Status=JobStatus.SUCCESS,
            OutputFileData=final_file_data,
            OutputFileName=filename,
            FontFamily=user_font,
            FontSize=user_size
        )

        db.session.add(history)
        db.session.commit()

        processing_time = round(time.time() - start_time, 2)
        processing_count = history.ProcessingCount

        return jsonify({
            'success': True, 
            'historyId': history.Id, 
            'jobId': job.Id,
            'jobName': job.JobName,
            'processingTime': processing_time,
            'time': datetime.now().strftime('%d %b %Y, %I:%M %p'),
            'processingCount': processing_count
        })

    except Exception as e:
        db.session.rollback()
        import traceback
        traceback.print_exc()
        return jsonify({'success': False, 'error': str(e)}), 500


# ==============================
# PROCESS FILE
# ==============================
@process_bp.route('/process-file', methods=['POST'])
@require_auth
def process_file(current_user):
    try:
        start_time = time.time()
        if 'file' not in request.files:
            return jsonify({'success': False, 'error': 'No file uploaded'})

        file = request.files['file']
        if file.filename == '':
            return jsonify({'success': False, 'error': 'No file selected'})

        user_font = request.form.get('fontFamily', 'Calibri')
        user_size = int(request.form.get('fontSize', 11))
        include_cover = request.form.get('includeCover') == 'true'
        include_toc = request.form.get('includeTOC') == 'true'
        job_id = request.form.get('jobId', None)

        with tempfile.TemporaryDirectory() as tmpdir:
            input_path = os.path.join(tmpdir, secure_filename(file.filename))
            output_path = os.path.join(tmpdir, "output.docx")

            file.save(input_path)

            processor = DocumentProcessor(
                template_path=TEMPLATE_DOCX,
                font_family=user_font,
                font_size=user_size,
                include_cover=include_cover,
                include_toc=include_toc
            )

            result = processor.universal_extract(input_path, output_path)

            if not result.get('success', False):
                return jsonify({'success': False, 'error': result.get('error')}), 500

            with open(output_path, 'rb') as f:
                final_file_data = f.read()

            with open(input_path, 'rb') as f:
                original_file_data = f.read()

        # Job handling
        if job_id and job_id != 'null':
            job = ProcessingJob.query.filter_by(Id=job_id, UserId=current_user.Id).first()
            if not job:
                return jsonify({'success': False, 'error': 'Conversation not found'}), 404
        else:
            base_name = os.path.splitext(file.filename)[0]
            job = ProcessingJob(JobName=base_name, UserId=current_user.Id)
            db.session.add(job)
            db.session.flush()

        base_name = os.path.splitext(file.filename)[0]
        filename = generate_unique_filename(base_name, current_user)

        history = ProcessingJobHistory(
            ProcessJobId=job.Id,
            JobType=JobType.FILE,
            Summary=file.filename,
            UploadFileName=file.filename,
            UploadFileData=original_file_data,
            Status=JobStatus.SUCCESS,
            OutputFileData=final_file_data,
            OutputFileName=filename,
            FontFamily=user_font,
            FontSize=user_size
        )

        db.session.add(history)
        db.session.commit()

        processing_time = round(time.time() - start_time, 2)
        processing_count = history.ProcessingCount

        return jsonify({
            'success': True, 
            'historyId': history.Id, 
            'jobId': job.Id,
            'jobName': job.JobName,
            'processingTime': processing_time,
            'time': datetime.now().strftime('%d %b %Y, %I:%M %p'),
            'processingCount': processing_count
        })

    except Exception as e:
        db.session.rollback()
        import traceback
        traceback.print_exc()
        return jsonify({'success': False, 'error': str(e)}), 500


# ==============================
# UPDATE TITLE (UNCHANGED)
# ==============================
@process_bp.route('/update-title/<int:job_id>', methods=['POST'])
@require_auth
def update_title(current_user, job_id):
    try:
        data = request.json
        new_title = data.get('title', '').strip()
        if not new_title:
            return jsonify({'success': False, 'error': 'Title is required'}), 400

        job = ProcessingJob.query.filter_by(Id=job_id, UserId=current_user.Id).first()
        if not job:
            return jsonify({'success': False, 'error': 'Job not found'}), 404

        job.JobName = new_title
        db.session.commit()
        return jsonify({'success': True})

    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


# ==============================
# DELETE
# ==============================
@process_bp.route('/delete/<int:job_id>', methods=['DELETE'])
@require_auth
def delete_conversation(current_user, job_id):
    try:
        job = ProcessingJob.query.filter_by(Id=job_id, UserId=current_user.Id).first()
        if not job:
            return jsonify({'success': False, 'error': 'Job not found'}), 404

        ProcessingJobHistory.query.filter_by(ProcessJobId=job.Id).delete()
        db.session.delete(job)
        db.session.commit()

        return jsonify({'success': True})

    except Exception as e:
        db.session.rollback()
        return jsonify({'success': False, 'error': str(e)}), 500


# ==============================
# SIDEBAR
# ==============================
@process_bp.route('/conversations', methods=['GET'])
@require_auth
def get_conversations(current_user):
    try:
        jobs = ProcessingJob.query.filter_by(UserId=current_user.Id)\
            .order_by(ProcessingJob.CreatedDate.desc()).all()

        return jsonify([{
            "id": job.Id,
            "title": job.JobName,
            "lastUpdate": job.CreatedDate.isoformat() if job.CreatedDate else None,
            "messageCount": ProcessingJobHistory.query.filter_by(ProcessJobId=job.Id).count(),
            "IsFavorite": job.IsFavorite
        } for job in jobs]), 200
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'success': False, 'error': str(e)}), 500

@process_bp.route('/favorite/<int:job_id>', methods=['POST'])
@require_auth
def toggle_favorite(current_user, job_id):
    job = ProcessingJob.query.filter_by(
        Id=job_id, 
        UserId=current_user.Id
    ).first_or_404()
    
    job.IsFavorite = not job.IsFavorite   # ✅ FIXED
    
    db.session.commit()
    
    return jsonify({
        'success': True, 
        'isFavorite': job.IsFavorite
    })

# ==============================
# EDIT & RE-PROCESS TEXT
# ==============================
@process_bp.route('/edit-text/<int:history_id>', methods=['POST'])
@require_auth
def edit_text(current_user, history_id):
    try:
        start_time = time.time()
        data = request.json
        new_text = data.get('text', '')
        
        # 1. Fetch the existing history record
        original_history = ProcessingJobHistory.query.get_or_404(history_id)
        job = ProcessingJob.query.get(original_history.ProcessJobId)

        if job.UserId != current_user.Id:
            return jsonify({'success': False, 'error': 'Unauthorized'}), 403

        # 2. Re-run the DocumentProcessor with new text
        processor = DocumentProcessor(
            template_path=TEMPLATE_DOCX,
            font_family=original_history.FontFamily,
            font_size=original_history.FontSize,
            include_cover=False, # Or pull from job settings
            include_toc=False
        )

        doc_obj = processor.html_to_docx(new_text)
        
        output_buffer = BytesIO()
        doc_obj.save(output_buffer)
        output_buffer.seek(0)
        final_file_data = output_buffer.read()

        # 3. Create a NEW history entry (to keep the "chat" thread intact)
        new_count = original_history.ProcessingCount + 1
        
        # NEW: Generate a descriptive filename for the edited version
        new_filename = generate_unique_filename(job.JobName, current_user)

        new_history = ProcessingJobHistory(
            ProcessJobId=job.Id,
            JobType=JobType.TEXT,
            Summary=re.sub('<[^<]+?>', '', new_text)[:100],
            UploadFileData=new_text.encode('utf-8'),
            Status=JobStatus.SUCCESS,
            OutputFileData=final_file_data,
            OutputFileName=new_filename,
            FontFamily=original_history.FontFamily,
            FontSize=original_history.FontSize,
            ProcessingCount=new_count
        )

        db.session.add(new_history)
        db.session.commit()

        processing_time = round(time.time() - start_time, 2)

        return jsonify({
            'success': True, 
            'message': 'Text updated and re-processed',
            'historyId': new_history.Id,
            'processingTime': processing_time,
            'time': datetime.now().strftime('%d %b %Y, %I:%M %p'),
            'processingCount': new_count
        })

    except Exception as e:
        db.session.rollback()
        return jsonify({'success': False, 'error': str(e)}), 500

# ==============================
# LOAD CHAT
# ==============================
@process_bp.route('/conversation/<int:job_id>', methods=['GET'])
@require_auth
def get_conversation(current_user, job_id):
    """Load full conversation history"""
    try:
        job = ProcessingJob.query.filter_by(Id=job_id, UserId=current_user.Id).first()
        if not job:
            return jsonify({'error': 'Unauthorized or not found'}), 404
        
        histories = ProcessingJobHistory.query.filter_by(ProcessJobId=job_id)\
            .order_by(ProcessingJobHistory.CreatedDate.asc()).all()
        
        # ✅ Include file data for rendering
        return jsonify([h.to_dict(include_file_data=True) for h in histories]), 200
        
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500

# ==============================
# DOWNLOAD
# ==============================
@process_bp.route('/download/<int:history_id>', methods=['GET'])
@require_auth
def download_file(current_user, history_id):
    try:
        history = ProcessingJobHistory.query.get_or_404(history_id)
        job = ProcessingJob.query.get(history.ProcessJobId)

        if job.UserId != current_user.Id:
            return jsonify({'success': False, 'error': 'Unauthorized access'}), 403

        requested_format = request.args.get('format', 'docx').lower()

        # NEW: Construct a descriptive download name based on the CURRENT job title
        job_name_clean = normalize_filename(job.JobName)
        # We reuse the same timestamping/count logic if needed, or just the current title
        # history.OutputFileName contains a timestamp by default; we replace the prefix
        old_ext = os.path.splitext(history.OutputFileName)[1]
        timestamp = ""
        # Extract timestamp if possible (e.g., Title_2026-04-01_11-43.docx)
        match = re.search(r'_\d{4}-\d{2}-\d{2}_\d{2}-\d{2}', history.OutputFileName)
        if match:
             timestamp = match.group(0)
             
        final_download_name = f"{job_name_clean}{timestamp}{old_ext}"

        if requested_format == 'pdf':
            pythoncom.CoInitialize()
            docx_path = None
            pdf_path = None

            try:
                docx_fd, docx_path = tempfile.mkstemp(suffix='.docx')
                with os.fdopen(docx_fd, 'wb') as f:
                    f.write(history.OutputFileData)

                pdf_fd, pdf_path = tempfile.mkstemp(suffix='.pdf')
                os.close(pdf_fd)

                convert(docx_path, pdf_path)

                with open(pdf_path, 'rb') as f:
                    pdf_data = f.read()

                return send_file(
                    BytesIO(pdf_data),
                    mimetype='application/pdf',
                    as_attachment=True,
                    download_name=final_download_name.replace(old_ext, '.pdf')
                )

            finally:
                if docx_path and os.path.exists(docx_path): os.remove(docx_path)
                if pdf_path and os.path.exists(pdf_path): os.remove(pdf_path)
                pythoncom.CoUninitialize()

        else:
            return send_file(
                BytesIO(history.OutputFileData),
                mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                as_attachment=True,
                download_name=final_download_name
            )

    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'success': False, 'error': str(e)}), 500