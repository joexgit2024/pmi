from flask import Blueprint, render_template, request, jsonify, flash, redirect, url_for, current_app
from werkzeug.utils import secure_filename
from app.models import FileUpload
from app.utils.file_utils import allowed_file, validate_excel_file
from app import db
import os
from datetime import datetime

bp = Blueprint('upload', __name__)

@bp.route('/', methods=['GET', 'POST'])
def index():
    """File upload dashboard and handler."""
    
    if request.method == 'POST':
        return handle_file_upload()
    
    # GET request - show the upload page
    # Get recent uploads
    recent_uploads = FileUpload.query.order_by(FileUpload.upload_date.desc()).limit(10).all()
    
    # Get statistics
    stats = {
        'total_files': FileUpload.query.count(),
        'registration_files': FileUpload.query.filter_by(file_type='registration').count(),
        'charity_files': FileUpload.query.filter_by(file_type='charity').count(),
        'processing': FileUpload.query.filter_by(status='processing').count(),
        'errors': FileUpload.query.filter_by(status='error').count()
    }
    
    return render_template('upload.html', 
                         recent_uploads=recent_uploads, 
                         upload_stats=stats)

def handle_file_upload():
    """Handle file upload via drag & drop or form."""
    
    try:
        print(f"DEBUG: Request files: {request.files}")
        print(f"DEBUG: Request form: {request.form}")
        
        # Check if file was uploaded
        if 'file' not in request.files:
            print("DEBUG: No file in request")
            return jsonify({'error': 'No file provided'}), 400
        
        file = request.files['file']
        file_type = request.form.get('file_type')
        
        print(f"DEBUG: File: {file.filename}, Type: {file_type}")
        
        if file.filename == '':
            return jsonify({'error': 'No file selected'}), 400
        
        if not file_type or file_type not in ['registration', 'charity']:
            return jsonify({'error': 'Invalid file type'}), 400
        
        if not allowed_file(file.filename):
            return jsonify({'error': 'Invalid file format. Only Excel (.xlsx, .xls) and CSV (.csv) files are allowed.'}), 400
        
        # Validate file size
        file.seek(0, os.SEEK_END)
        file_size = file.tell()
        file.seek(0)
        
        if file_size > current_app.config['MAX_CONTENT_LENGTH']:
            return jsonify({'error': 'File too large. Maximum size is 50MB.'}), 400
        
        # Secure filename
        original_filename = file.filename
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        filename = f"{timestamp}_{secure_filename(original_filename)}"
        
        # Ensure upload directory exists
        upload_folder = current_app.config['UPLOAD_FOLDER']
        os.makedirs(upload_folder, exist_ok=True)
        
        # Save file
        file_path = os.path.join(upload_folder, filename)
        file.save(file_path)
        
        # Validate file (for now, just check if it's readable)
        try:
            validation_result = validate_excel_file(file_path)
            if not validation_result['valid']:
                os.remove(file_path)  # Remove invalid file
                return jsonify({'error': validation_result['message']}), 400
        except Exception as e:
            if os.path.exists(file_path):
                os.remove(file_path)  # Remove file on validation error
            return jsonify({'error': f'File validation failed: {str(e)}'}), 400
        
        # Create database record
        file_upload = FileUpload(
            filename=filename,
            original_filename=original_filename,
            file_type=file_type,
            file_path=file_path,
            file_size=file_size,
            status='uploaded',
            rows_count=validation_result.get('rows', 0)
        )
        
        db.session.add(file_upload)
        db.session.commit()
        
        return jsonify({
            'success': True,
            'message': 'File uploaded successfully!',
            'file_id': file_upload.id,
            'filename': original_filename,
            'file_type': file_type,
            'rows_count': validation_result.get('rows', 0)
        })
        
    except Exception as e:
        current_app.logger.error(f"Upload error: {str(e)}")
        return jsonify({'error': 'Upload failed. Please try again.'}), 500

@bp.route('/progress/<int:file_id>')
def upload_progress(file_id):
    """Get upload and processing progress."""
    
    file_upload = FileUpload.query.get_or_404(file_id)
    
    return jsonify({
        'file_id': file_id,
        'filename': file_upload.original_filename,
        'status': file_upload.status,
        'rows_count': file_upload.rows_count,
        'error_message': file_upload.error_message,
        'upload_date': file_upload.upload_date.isoformat() if file_upload.upload_date else None
    })

@bp.route('/files')
def list_files():
    """Get list of uploaded files."""
    
    page = request.args.get('page', 1, type=int)
    per_page = request.args.get('per_page', 20, type=int)
    file_type = request.args.get('file_type')
    status = request.args.get('status')
    
    query = FileUpload.query
    
    if file_type:
        query = query.filter_by(file_type=file_type)
    
    if status:
        query = query.filter_by(status=status)
    
    files = query.order_by(FileUpload.upload_date.desc()).paginate(
        page=page, per_page=per_page, error_out=False
    )
    
    return jsonify({
        'files': [f.to_dict() for f in files.items],
        'total': files.total,
        'pages': files.pages,
        'current_page': files.page,
        'has_next': files.has_next,
        'has_prev': files.has_prev
    })

@bp.route('/delete/<int:file_id>', methods=['DELETE'])
def delete_file(file_id):
    """Delete an uploaded file."""
    
    try:
        file_upload = FileUpload.query.get_or_404(file_id)
        
        # Remove physical file
        if os.path.exists(file_upload.file_path):
            os.remove(file_upload.file_path)
        
        # Remove database record
        db.session.delete(file_upload)
        db.session.commit()
        
        return jsonify({'success': True, 'message': 'File deleted successfully.'})
        
    except Exception as e:
        current_app.logger.error(f"Delete file error: {str(e)}")
        return jsonify({'error': 'Failed to delete file.'}), 500