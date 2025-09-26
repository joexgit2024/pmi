from flask import Blueprint, render_template, jsonify, send_file
from app.models.file_upload import FileUpload
from app.models.registration import Registration
from app.models.charity import Charity
from app.models.email_tracking import EmailTracking
from app.models.matching import MatchingResult
from app import db
from sqlalchemy import func, desc
from datetime import datetime, timedelta
import os

bp = Blueprint('main', __name__)

@bp.route('/')
def index():
    """Main dashboard page."""
    
    # Get system statistics
    stats = get_system_stats()
    
    # Get recent activity
    recent_activity = get_recent_activity()
    
    # Get latest files
    latest_files = FileUpload.query.order_by(desc(FileUpload.upload_date)).limit(5).all()
    
    return render_template('dashboard.html', 
                         stats=stats, 
                         recent_activity=recent_activity,
                         latest_files=latest_files)

# Removed duplicate status route - using the one at the end of file

def get_system_stats():
    """Get comprehensive system statistics."""
    
    # File upload statistics
    total_files = FileUpload.query.count()
    registration_files = FileUpload.query.filter_by(file_type='registration').count()
    charity_files = FileUpload.query.filter_by(file_type='charity').count()
    
    # Registration statistics
    total_registrations = Registration.query.count()
    registrations_with_linkedin = Registration.query.filter(
        Registration.linkedin_url.isnot(None),
        Registration.linkedin_url != ''
    ).count()
    
    # Charity statistics
    total_charities = Charity.query.count()
    available_positions = db.session.query(func.sum(Charity.volunteers_needed)).scalar() or 0
    assigned_positions = db.session.query(func.sum(Charity.volunteers_assigned)).scalar() or 0
    
    # Matching statistics
    total_matches = MatchingResult.query.count()
    avg_match_score = db.session.query(func.avg(MatchingResult.match_score)).scalar() or 0
    
    # Email statistics
    total_emails = EmailTracking.query.count()
    emails_sent = EmailTracking.query.filter_by(status='sent').count()
    emails_drafted = EmailTracking.query.filter_by(status='drafted').count()
    
    # Latest activity dates
    latest_registration_file = FileUpload.query.filter_by(
        file_type='registration'
    ).order_by(desc(FileUpload.upload_date)).first()
    
    latest_charity_file = FileUpload.query.filter_by(
        file_type='charity'
    ).order_by(desc(FileUpload.upload_date)).first()
    
    return {
        'files': {
            'total': total_files,
            'registration_files': registration_files,
            'charity_files': charity_files,
            'latest_registration': latest_registration_file.upload_date.strftime('%Y-%m-%d') if latest_registration_file else None,
            'latest_charity': latest_charity_file.upload_date.strftime('%Y-%m-%d') if latest_charity_file else None
        },
        'registrations': {
            'total': total_registrations,
            'with_linkedin': registrations_with_linkedin,
            'linkedin_percentage': round((registrations_with_linkedin / total_registrations * 100), 1) if total_registrations > 0 else 0
        },
        'charities': {
            'total': total_charities,
            'available_positions': int(available_positions) if available_positions else 0,
            'assigned_positions': int(assigned_positions) if assigned_positions else 0,
            'utilization_percentage': round((assigned_positions / available_positions * 100), 1) if available_positions > 0 else 0
        },
        'matching': {
            'total_matches': total_matches,
            'avg_match_score': round(float(avg_match_score), 2) if avg_match_score else 0
        },
        'emails': {
            'total': total_emails,
            'sent': emails_sent,
            'drafted': emails_drafted,
            'pending': max(0, emails_drafted - emails_sent)
        }
    }

def get_recent_activity(limit=10):
    """Get recent system activity."""
    
    activities = []
    
    # Recent file uploads
    recent_uploads = FileUpload.query.order_by(desc(FileUpload.upload_date)).limit(5).all()
    for upload in recent_uploads:
        activities.append({
            'type': 'file_upload',
            'message': f"Uploaded {upload.file_type} file: {upload.original_filename}",
            'timestamp': upload.upload_date,
            'icon': 'upload',
            'color': 'primary'
        })
    
    # Recent matching results
    recent_matches = MatchingResult.query.order_by(desc(MatchingResult.created_at)).limit(3).all()
    for match in recent_matches:
        activities.append({
            'type': 'matching',
            'message': f"Matched {match.registration.full_name} to {match.charity.organization}",
            'timestamp': match.created_at,
            'icon': 'link',
            'color': 'success'
        })
    
    # Recent emails
    recent_emails = EmailTracking.query.order_by(desc(EmailTracking.created_at)).limit(3).all()
    for email in recent_emails:
        activities.append({
            'type': 'email',
            'message': f"Generated email for {email.recipient_name}",
            'timestamp': email.created_at,
            'icon': 'mail',
            'color': 'info'
        })
    
    # Sort by timestamp and limit
    activities.sort(key=lambda x: x['timestamp'], reverse=True)
    return activities[:limit]

@bp.route('/status')
def status():
    """System status page."""
    try:
        # Basic system health checks
        status_info = {
            'database': 'Connected',
            'file_system': 'Available',
            'upload_folder': os.path.exists('uploads'),
            'total_files': FileUpload.query.count(),
            'total_registrations': Registration.query.count(),
            'total_charities': Charity.query.count()
        }
        
        return render_template('status.html', status=status_info)
        
    except Exception as e:
        return render_template('status.html', status={'error': str(e)})

@bp.route('/download/<filename>')
def download_file(filename):
    """Download exported files."""
    try:
        # This is a placeholder - in production, implement proper file serving
        download_folder = os.path.join(os.getcwd(), 'downloads')
        file_path = os.path.join(download_folder, filename)
        
        if os.path.exists(file_path):
            return send_file(file_path, as_attachment=True)
        else:
            return jsonify({'error': 'File not found'}), 404
            
    except Exception as e:
        return jsonify({'error': str(e)}), 500