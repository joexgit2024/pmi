from flask import Blueprint, render_template, request, jsonify, flash, redirect, url_for
from app.models.email_tracking import EmailTracking
from app.models.matching import MatchingResult
from app.models.registration import Registration
from app import db

email = Blueprint('email', __name__, url_prefix='/email')
bp = email  # Alias for consistent import

@email.route('/')
def index():
    """Display email management page."""
    try:
        # Get email statistics
        total_emails = EmailTracking.query.count()
        sent_emails = EmailTracking.query.filter_by(status='sent').count()
        pending_emails = EmailTracking.query.filter_by(status='pending').count()
        failed_emails = EmailTracking.query.filter_by(status='failed').count()
        
        # Get recent emails
        recent_emails = EmailTracking.query.order_by(
            EmailTracking.created_at.desc()
        ).limit(10).all()
        
        # Get ready to send count (successful matches without emails)
        ready_to_send = MatchingResult.query.filter(
            MatchingResult.status == 'matched',
            ~MatchingResult.registration_id.in_(
                db.session.query(EmailTracking.registration_id)
                .filter(EmailTracking.status == 'sent')
            )
        ).count()
        
        stats = {
            'total_emails': total_emails,
            'sent_emails': sent_emails,
            'pending_emails': pending_emails,
            'failed_emails': failed_emails,
            'ready_to_send': ready_to_send
        }
        
        return render_template('email.html', 
                             stats=stats, 
                             recent_emails=recent_emails)
        
    except Exception as e:
        flash(f'Error loading email page: {str(e)}', 'error')
        return render_template('email.html', 
                             stats={}, 
                             recent_emails=[])

@email.route('/generate', methods=['POST'])
def generate_emails():
    """Generate email drafts for matched participants."""
    try:
        # Import email service
        from app.services.email_service import EmailService
        
        email_service = EmailService()
        result = email_service.generate_email_drafts()
        
        if result['success']:
            flash(f"Email generation completed! {result['generated_count']} emails generated.", 'success')
        else:
            flash(f"Email generation failed: {result['message']}", 'error')
        
        return jsonify(result)
        
    except Exception as e:
        error_result = {
            'success': False,
            'message': f'Error generating emails: {str(e)}'
        }
        return jsonify(error_result), 500

@email.route('/send', methods=['POST'])
def send_emails():
    """Send pending email drafts."""
    try:
        email_ids = request.json.get('email_ids', [])
        
        if not email_ids:
            return jsonify({
                'success': False,
                'message': 'No emails selected'
            }), 400
        
        # Import email service
        from app.services.email_service import EmailService
        
        email_service = EmailService()
        result = email_service.send_emails(email_ids)
        
        return jsonify(result)
        
    except Exception as e:
        return jsonify({
            'success': False,
            'message': f'Error sending emails: {str(e)}'
        }), 500

@email.route('/drafts')
def drafts():
    """Display email drafts."""
    try:
        page = request.args.get('page', 1, type=int)
        per_page = 20
        status_filter = request.args.get('status', 'all')
        
        query = EmailTracking.query
        
        if status_filter != 'all':
            query = query.filter_by(status=status_filter)
        
        emails = query.order_by(
            EmailTracking.created_at.desc()
        ).paginate(
            page=page, per_page=per_page, error_out=False
        )
        
        return render_template('email_drafts.html', 
                             emails=emails, 
                             status_filter=status_filter)
        
    except Exception as e:
        flash(f'Error loading email drafts: {str(e)}', 'error')
        return redirect(url_for('email.index'))

@email.route('/preview/<int:email_id>')
def preview_email(email_id):
    """Preview email content."""
    try:
        email_tracking = EmailTracking.query.get_or_404(email_id)
        
        return render_template('email_preview.html', email=email_tracking)
        
    except Exception as e:
        flash(f'Error loading email preview: {str(e)}', 'error')
        return redirect(url_for('email.drafts'))

@email.route('/edit/<int:email_id>', methods=['GET', 'POST'])
def edit_email(email_id):
    """Edit email content."""
    try:
        email_tracking = EmailTracking.query.get_or_404(email_id)
        
        if request.method == 'POST':
            # Update email content
            email_tracking.subject = request.form.get('subject', '')
            email_tracking.body = request.form.get('body', '')
            
            db.session.commit()
            
            flash('Email updated successfully', 'success')
            return redirect(url_for('email.preview', email_id=email_id))
        
        return render_template('email_edit.html', email=email_tracking)
        
    except Exception as e:
        flash(f'Error editing email: {str(e)}', 'error')
        return redirect(url_for('email.drafts'))

@email.route('/delete/<int:email_id>', methods=['POST'])
def delete_email(email_id):
    """Delete email draft."""
    try:
        email_tracking = EmailTracking.query.get_or_404(email_id)
        
        if email_tracking.status == 'sent':
            return jsonify({
                'success': False,
                'message': 'Cannot delete sent emails'
            }), 400
        
        db.session.delete(email_tracking)
        db.session.commit()
        
        return jsonify({
            'success': True,
            'message': 'Email deleted successfully'
        })
        
    except Exception as e:
        db.session.rollback()
        return jsonify({
            'success': False,
            'message': f'Error deleting email: {str(e)}'
        }), 500

@email.route('/bulk-send', methods=['POST'])
def bulk_send():
    """Send multiple emails at once."""
    try:
        email_ids = request.json.get('email_ids', [])
        
        if not email_ids:
            return jsonify({
                'success': False,
                'message': 'No emails selected'
            }), 400
        
        # Import email service
        from app.services.email_service import EmailService
        
        email_service = EmailService()
        result = email_service.bulk_send_emails(email_ids)
        
        return jsonify(result)
        
    except Exception as e:
        return jsonify({
            'success': False,
            'message': f'Error sending emails: {str(e)}'
        }), 500

@email.route('/export')
def export_emails():
    """Export email tracking data."""
    try:
        from app.services.export_service import ExportService
        
        export_service = ExportService()
        file_path = export_service.export_email_data()
        
        if file_path:
            return jsonify({
                'success': True,
                'download_url': url_for('main.download_file', filename=file_path)
            })
        else:
            return jsonify({
                'success': False,
                'message': 'Failed to export email data'
            }), 500
            
    except Exception as e:
        return jsonify({
            'success': False,
            'message': f'Error exporting email data: {str(e)}'
        }), 500