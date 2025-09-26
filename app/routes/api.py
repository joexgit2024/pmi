from flask import Blueprint, jsonify, request
from app.models.file_upload import FileUpload
from app.models.registration import Registration
from app.models.charity import Charity
from app.models.matching import MatchingResult
from app.models.email_tracking import EmailTracking
from app import db
from sqlalchemy import desc, func
from datetime import datetime
import os

api = Blueprint('api', __name__, url_prefix='/api')
bp = api  # Alias for consistent import

@api.route('/dashboard/stats')
def dashboard_stats():
    """Get dashboard statistics."""
    try:
        total_registrations = Registration.query.count()
        # Successful matches: treat any MatchingResult as a match (adjust if status column differs)
        successful_matches = MatchingResult.query.count()
        emails_sent = EmailTracking.query.filter_by(status='sent').count()
        pending_emails = EmailTracking.query.filter(EmailTracking.status.in_(['pending','drafted'])).count()

        # Last analysis run time inferred from latest MatchingBatch or output file timestamps
        from app.models import MatchingBatch
        latest_batch = MatchingBatch.query.order_by(desc(MatchingBatch.created_at)).first()
        last_analysis_time = None
        if latest_batch:
            last_analysis_time = latest_batch.created_at.isoformat()
        else:
            # Fallback: check Output/Matching_Summary.csv mtime
            output_csv = os.path.join(os.getcwd(), 'Output', 'Matching_Summary.csv')
            if os.path.exists(output_csv):
                last_analysis_time = datetime.fromtimestamp(os.path.getmtime(output_csv)).isoformat()

        stats = {
            'total_registrations': total_registrations,
            'successful_matches': successful_matches,
            'emails_sent': emails_sent,
            'pending_emails': pending_emails,
            'last_analysis_time': last_analysis_time
        }
        
        return jsonify({
            'success': True,
            'stats': stats
        })
        
    except Exception as e:
        return jsonify({
            'success': False,
            'message': f'Error fetching dashboard stats: {str(e)}'
        }), 500

@api.route('/activity')
def recent_activity():
    """Return recent activity (files, matches, emails) as JSON."""
    try:
        activities = []

        # File uploads
        uploads = FileUpload.query.order_by(desc(FileUpload.upload_date)).limit(5).all()
        for u in uploads:
            activities.append({
                'type': 'file_upload',
                'timestamp': u.upload_date.isoformat() if u.upload_date else None,
                'message': f"Uploaded {u.file_type} file: {u.original_filename}",
                'status': 'success'
            })

        # Matches
        matches = MatchingResult.query.order_by(desc(MatchingResult.created_at)).limit(5).all()
        for m in matches:
            try:
                activities.append({
                    'type': 'matching',
                    'timestamp': m.created_at.isoformat() if m.created_at else None,
                    'message': f"Matched {m.registration.first_name} {m.registration.last_name} to {m.charity.organization}",
                    'status': 'success'
                })
            except Exception:
                pass

        # Emails
        emails = EmailTracking.query.order_by(desc(EmailTracking.created_at)).limit(5).all()
        for em in emails:
            activities.append({
                'type': 'email',
                'timestamp': em.created_at.isoformat() if em.created_at else None,
                'message': f"Email for {em.recipient_name or em.recipient_email}",
                'status': em.status
            })

        # Sort combined
        activities.sort(key=lambda x: x['timestamp'] or '', reverse=True)
        return jsonify({'success': True, 'activities': activities[:15]})
    except Exception as e:
        return jsonify({'success': False, 'message': f'Error fetching activity: {str(e)}'}), 500

@api.route('/system/refresh', methods=['POST'])
def system_refresh():
    """Refresh system data."""
    try:
        # This could trigger various refresh operations
        # For now, just return success
        return jsonify({
            'success': True,
            'message': 'System refreshed successfully'
        })
        
    except Exception as e:
        return jsonify({
            'success': False,
            'message': f'Error refreshing system: {str(e)}'
        }), 500

@api.route('/matching/run', methods=['POST'])
def run_matching():
    """API endpoint to run matching process."""
    try:
        from app.services.matching_service import MatchingService
        
        matching_service = MatchingService()
        result = matching_service.run_matching()
        
        return jsonify(result)
        
    except Exception as e:
        return jsonify({
            'success': False,
            'message': f'Error running matching: {str(e)}'
        }), 500

@api.route('/email/generate', methods=['POST'])
def generate_emails():
    """API endpoint to generate emails."""
    try:
        from app.services.email_service import EmailService
        
        email_service = EmailService()
        result = email_service.generate_email_drafts()
        
        return jsonify(result)
        
    except Exception as e:
        return jsonify({
            'success': False,
            'message': f'Error generating emails: {str(e)}'
        }), 500

@api.route('/files')
def list_files():
    """Get list of uploaded files."""
    try:
        files = FileUpload.query.order_by(FileUpload.created_at.desc()).all()
        
        files_data = []
        for file in files:
            files_data.append({
                'id': file.id,
                'original_filename': file.original_filename,
                'file_type': file.file_type,
                'status': file.status,
                'created_at': file.created_at.isoformat(),
                'file_size': file.file_size
            })
        
        return jsonify({
            'success': True,
            'files': files_data
        })
        
    except Exception as e:
        return jsonify({
            'success': False,
            'message': f'Error listing files: {str(e)}'
        }), 500

@api.route('/registrations')
def list_registrations():
    """Get list of registrations."""
    try:
        page = request.args.get('page', 1, type=int)
        per_page = request.args.get('per_page', 50, type=int)
        
        registrations = Registration.query.paginate(
            page=page, per_page=per_page, error_out=False
        )
        
        registrations_data = []
        for reg in registrations.items:
            registrations_data.append({
                'id': reg.id,
                'name': reg.name,
                'email': reg.email,
                'linkedin_url': reg.linkedin_url,
                'created_at': reg.created_at.isoformat() if reg.created_at else None
            })
        
        return jsonify({
            'success': True,
            'registrations': registrations_data,
            'total': registrations.total,
            'pages': registrations.pages,
            'current_page': registrations.page
        })
        
    except Exception as e:
        return jsonify({
            'success': False,
            'message': f'Error listing registrations: {str(e)}'
        }), 500

@api.route('/charities')
def list_charities():
    """Get list of charities."""
    try:
        charities = Charity.query.all()
        
        charities_data = []
        for charity in charities:
            charities_data.append({
                'id': charity.id,
                'charity_name': charity.charity_name,
                'problem_statement': charity.problem_statement,
                'contact_email': charity.contact_email,
                'created_at': charity.created_at.isoformat() if charity.created_at else None
            })
        
        return jsonify({
            'success': True,
            'charities': charities_data
        })
        
    except Exception as e:
        return jsonify({
            'success': False,
            'message': f'Error listing charities: {str(e)}'
        }), 500

@api.route('/matching-results')
def list_matching_results():
    """Get list of matching results."""
    try:
        page = request.args.get('page', 1, type=int)
        per_page = request.args.get('per_page', 50, type=int)
        
        matches = MatchingResult.query.paginate(
            page=page, per_page=per_page, error_out=False
        )
        
        matches_data = []
        for match in matches.items:
            matches_data.append({
                'id': match.id,
                'registration_id': match.registration_id,
                'charity_id': match.charity_id,
                'confidence_score': match.confidence_score,
                'status': match.status,
                'created_at': match.created_at.isoformat() if match.created_at else None
            })
        
        return jsonify({
            'success': True,
            'matches': matches_data,
            'total': matches.total,
            'pages': matches.pages,
            'current_page': matches.page
        })
        
    except Exception as e:
        return jsonify({
            'success': False,
            'message': f'Error listing matching results: {str(e)}'
        }), 500

@api.route('/emails')
def list_emails():
    """Get list of email tracking records."""
    try:
        page = request.args.get('page', 1, type=int)
        per_page = request.args.get('per_page', 50, type=int)
        
        emails = EmailTracking.query.paginate(
            page=page, per_page=per_page, error_out=False
        )
        
        emails_data = []
        for email in emails.items:
            emails_data.append({
                'id': email.id,
                'recipient_email': email.recipient_email,
                'subject': email.subject,
                'status': email.status,
                'created_at': email.created_at.isoformat() if email.created_at else None,
                'sent_at': email.sent_at.isoformat() if email.sent_at else None
            })
        
        return jsonify({
            'success': True,
            'emails': emails_data,
            'total': emails.total,
            'pages': emails.pages,
            'current_page': emails.page
        })
        
    except Exception as e:
        return jsonify({
            'success': False,
            'message': f'Error listing emails: {str(e)}'
        }), 500

@api.route('/health')
def health_check():
    """Health check endpoint."""
    try:
        # Basic health checks
        db_status = 'ok'
        try:
            db.session.execute('SELECT 1')
        except:
            db_status = 'error'
        
        return jsonify({
            'success': True,
            'status': 'healthy',
            'database': db_status,
            'timestamp': db.func.now()
        })
        
    except Exception as e:
        return jsonify({
            'success': False,
            'status': 'unhealthy',
            'message': str(e)
        }), 500