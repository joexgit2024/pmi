from app import db
from datetime import datetime

class Registration(db.Model):
    """Model for PMP professional registrations."""
    
    __tablename__ = 'registrations'
    
    id = db.Column(db.Integer, primary_key=True)
    file_upload_id = db.Column(db.Integer, db.ForeignKey('file_uploads.id'), nullable=False)
    
    # Personal Information
    first_name = db.Column(db.String(100), nullable=False)
    last_name = db.Column(db.String(100), nullable=False)
    email = db.Column(db.String(255), nullable=False)
    phone = db.Column(db.String(50))
    preferred_email = db.Column(db.String(255))
    
    # Professional Information
    pmi_id = db.Column(db.String(20))
    job_title = db.Column(db.String(255))
    company = db.Column(db.String(255))
    linkedin_url = db.Column(db.String(500))
    experience_years = db.Column(db.String(50))
    
    # PMDoS Information
    areas_of_interest = db.Column(db.Text)
    first_time_participant = db.Column(db.Boolean)
    dietary_requirements = db.Column(db.Text)
    background_description = db.Column(db.Text)
    
    # Analysis Results
    linkedin_quality_score = db.Column(db.Float)
    profile_completeness_score = db.Column(db.Float)
    overall_score = db.Column(db.Float)
    
    # Metadata
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)
    
    # Relationships
    matching_results = db.relationship('MatchingResult', backref='registration', lazy='dynamic')
    email_tracking = db.relationship('EmailTracking', backref='registration', lazy='dynamic')
    
    def __repr__(self):
        return f'<Registration {self.first_name} {self.last_name}>'
    
    @property
    def full_name(self):
        return f"{self.first_name} {self.last_name}".strip()
    
    @property
    def primary_email(self):
        return self.preferred_email or self.email
    
    def to_dict(self):
        return {
            'id': self.id,
            'full_name': self.full_name,
            'email': self.primary_email,
            'pmi_id': self.pmi_id,
            'job_title': self.job_title,
            'company': self.company,
            'linkedin_url': self.linkedin_url,
            'experience_years': self.experience_years,
            'areas_of_interest': self.areas_of_interest,
            'linkedin_quality_score': self.linkedin_quality_score,
            'profile_completeness_score': self.profile_completeness_score,
            'overall_score': self.overall_score,
            'created_at': self.created_at.isoformat() if self.created_at else None
        }