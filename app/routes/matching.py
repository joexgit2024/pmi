from flask import Blueprint, render_template, request, jsonify, flash, redirect, url_for
from app.models.matching import MatchingResult
from app.models.registration import Registration
from app.models.charity import Charity
from app import db

matching = Blueprint('matching', __name__, url_prefix='/matching')
bp = matching  # Alias for consistent import

@matching.route('/')
def index():
    """Display matching management page."""
    try:
        # Get matching statistics
        total_matches = MatchingResult.query.count()
        successful_matches = MatchingResult.query.filter_by(status='matched').count()
        pending_matches = MatchingResult.query.filter_by(status='pending').count()
        failed_matches = MatchingResult.query.filter_by(status='failed').count()
        
        # Get recent matches
        recent_matches = MatchingResult.query.order_by(
            MatchingResult.created_at.desc()
        ).limit(10).all()
        
        # Get unmatched registrations
        unmatched_registrations = Registration.query.filter(
            ~Registration.id.in_(
                db.session.query(MatchingResult.registration_id)
                .filter(MatchingResult.status == 'matched')
            )
        ).count()
        
        # Get available charities
        available_charities = Charity.query.count()
        
        stats = {
            'total_matches': total_matches,
            'successful_matches': successful_matches,
            'pending_matches': pending_matches,
            'failed_matches': failed_matches,
            'unmatched_registrations': unmatched_registrations,
            'available_charities': available_charities
        }
        
        return render_template('matching.html', 
                             stats=stats, 
                             recent_matches=recent_matches)
        
    except Exception as e:
        flash(f'Error loading matching page: {str(e)}', 'error')
        return render_template('matching.html', 
                             stats={}, 
                             recent_matches=[])

@matching.route('/run', methods=['POST'])
def run_matching():
    """Run the charity matching process."""
    try:
        # Get request data if any (for future config options)
        request_data = request.get_json() if request.is_json else {}
        
        # Import matching service
        from app.services.matching_service import MatchingService
        
        matching_service = MatchingService()
        
        # Extract configuration options (for future use)
        use_flexible = request_data.get('allowMultiple', False)
        
        result = matching_service.run_matching(use_flexible=use_flexible)
        
        if result['success']:
            flash(f"Matching completed successfully! {result['matched_count']} matches created.", 'success')
        else:
            flash(f"Matching failed: {result['message']}", 'error')
        
        return jsonify(result)
        
    except Exception as e:
        error_result = {
            'success': False,
            'message': f'Error running matching: {str(e)}'
        }
        return jsonify(error_result), 500

@matching.route('/results')
def results():
    """Display matching results."""
    try:
        page = request.args.get('page', 1, type=int)
        per_page = 20
        
        matches = MatchingResult.query.order_by(
            MatchingResult.created_at.desc()
        ).paginate(
            page=page, per_page=per_page, error_out=False
        )
        
        return render_template('matching_results.html', matches=matches)
        
    except Exception as e:
        flash(f'Error loading matching results: {str(e)}', 'error')
        return redirect(url_for('matching.index'))

@matching.route('/export')
def export_results():
    """Export matching results to Excel."""
    try:
        from app.services.export_service import ExportService
        
        export_service = ExportService()
        file_path = export_service.export_matching_results()
        
        if file_path:
            return jsonify({
                'success': True,
                'download_url': url_for('main.download_file', filename=file_path)
            })
        else:
            return jsonify({
                'success': False,
                'message': 'Failed to export results'
            }), 500
            
    except Exception as e:
        return jsonify({
            'success': False,
            'message': f'Error exporting results: {str(e)}'
        }), 500

@matching.route('/match/<int:match_id>/approve', methods=['POST'])
def approve_match(match_id):
    """Approve a specific match."""
    try:
        match = MatchingResult.query.get_or_404(match_id)
        match.status = 'approved'
        db.session.commit()
        
        return jsonify({
            'success': True,
            'message': 'Match approved successfully'
        })
        
    except Exception as e:
        db.session.rollback()
        return jsonify({
            'success': False,
            'message': f'Error approving match: {str(e)}'
        }), 500

@matching.route('/match/<int:match_id>/reject', methods=['POST'])
def reject_match(match_id):
    """Reject a specific match."""
    try:
        match = MatchingResult.query.get_or_404(match_id)
        match.status = 'rejected'
        db.session.commit()
        
        return jsonify({
            'success': True,
            'message': 'Match rejected successfully'
        })
        
    except Exception as e:
        db.session.rollback()
        return jsonify({
            'success': False,
            'message': f'Error rejecting match: {str(e)}'
        }), 500