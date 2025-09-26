from app import db
from datetime import datetime

class MatchingResult(db.Model):
    """Model for storing PMP-Charity matching results."""
    
    __tablename__ = 'matching_results'
    
    id = db.Column(db.Integer, primary_key=True)
    registration_id = db.Column(db.Integer, db.ForeignKey('registrations.id'), nullable=False)
    charity_id = db.Column(db.Integer, db.ForeignKey('charities.id'), nullable=False)
    
    # Matching Scores
    match_score = db.Column(db.Float, nullable=False)
    linkedin_quality = db.Column(db.Float)
    experience_match = db.Column(db.Float)
    skills_match = db.Column(db.Float)
    interest_match = db.Column(db.Float)
    
    # Matching Context
    batch_id = db.Column(db.String(50), nullable=False)
    matching_algorithm = db.Column(db.String(50), default='enhanced_v1')
    assignment_rank = db.Column(db.Integer)  # 1st choice, 2nd choice, etc.
    
    # Status
    status = db.Column(db.String(20), default='proposed')  # proposed, confirmed, rejected
    notes = db.Column(db.Text)
    
    # Metadata
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)
    
    def __repr__(self):
        return f'<MatchingResult {self.registration.full_name} -> {self.charity.organization}>'
    
    def to_dict(self):
        return {
            'id': self.id,
            'registration_id': self.registration_id,
            'charity_id': self.charity_id,
            'pmp_name': self.registration.full_name if self.registration else None,
            'charity_name': self.charity.organization if self.charity else None,
            'match_score': self.match_score,
            'linkedin_quality': self.linkedin_quality,
            'experience_match': self.experience_match,
            'skills_match': self.skills_match,
            'interest_match': self.interest_match,
            'batch_id': self.batch_id,
            'assignment_rank': self.assignment_rank,
            'status': self.status,
            'created_at': self.created_at.isoformat() if self.created_at else None
        }


class MatchingBatch(db.Model):
    """Model for tracking matching batch operations."""
    
    __tablename__ = 'matching_batches'
    
    id = db.Column(db.Integer, primary_key=True)
    batch_id = db.Column(db.String(50), unique=True, nullable=False)
    
    # Batch Information
    registration_file_id = db.Column(db.Integer, db.ForeignKey('file_uploads.id'))
    charity_file_id = db.Column(db.Integer, db.ForeignKey('file_uploads.id'))
    algorithm_version = db.Column(db.String(50), default='enhanced_v1')
    matching_type = db.Column(db.String(20), default='standard')  # standard, flexible
    
    # Results
    total_registrations = db.Column(db.Integer)
    total_charities = db.Column(db.Integer)
    total_matches = db.Column(db.Integer)
    avg_match_score = db.Column(db.Float)
    
    # Status
    status = db.Column(db.String(20), default='running')  # running, completed, failed
    progress_percentage = db.Column(db.Integer, default=0)
    error_message = db.Column(db.Text)
    
    # Output Files
    excel_report_path = db.Column(db.String(500))
    csv_summary_path = db.Column(db.String(500))
    
    # Metadata
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    completed_at = db.Column(db.DateTime)
    
    def __repr__(self):
        return f'<MatchingBatch {self.batch_id}>'
    
    def to_dict(self):
        return {
            'id': self.id,
            'batch_id': self.batch_id,
            'algorithm_version': self.algorithm_version,
            'matching_type': self.matching_type,
            'total_registrations': self.total_registrations,
            'total_charities': self.total_charities,
            'total_matches': self.total_matches,
            'avg_match_score': self.avg_match_score,
            'status': self.status,
            'progress_percentage': self.progress_percentage,
            'created_at': self.created_at.isoformat() if self.created_at else None,
            'completed_at': self.completed_at.isoformat() if self.completed_at else None
        }