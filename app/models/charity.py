from app import db
from datetime import datetime

class Charity(db.Model):
    """Model for charity organizations and their projects."""
    
    __tablename__ = 'charities'
    
    id = db.Column(db.Integer, primary_key=True)
    file_upload_id = db.Column(db.Integer, db.ForeignKey('file_uploads.id'), nullable=False)
    
    # Organization Information
    organization = db.Column(db.String(255), nullable=False)
    contact_person = db.Column(db.String(255))
    contact_email = db.Column(db.String(255))
    contact_phone = db.Column(db.String(50))
    
    # Project Information
    initiative = db.Column(db.String(500), nullable=False)
    description = db.Column(db.Text)
    project_objectives = db.Column(db.Text)
    target_beneficiaries = db.Column(db.Text)
    
    # Requirements
    skills_required = db.Column(db.Text)
    experience_level = db.Column(db.String(100))
    time_commitment = db.Column(db.String(100))
    location = db.Column(db.String(255))
    
    # Analysis
    priority_level = db.Column(db.String(20))  # High, Medium, Low
    complexity = db.Column(db.String(20))  # High, Medium, Low
    category = db.Column(db.String(100))
    
    # Capacity
    volunteers_needed = db.Column(db.Integer, default=2)
    volunteers_assigned = db.Column(db.Integer, default=0)
    
    # Metadata
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)
    
    # Relationships
    matching_results = db.relationship('MatchingResult', backref='charity', lazy='dynamic')
    
    def __repr__(self):
        return f'<Charity {self.organization}>'
    
    @property
    def is_fully_assigned(self):
        return self.volunteers_assigned >= self.volunteers_needed
    
    @property
    def capacity_remaining(self):
        return max(0, self.volunteers_needed - self.volunteers_assigned)
    
    def to_dict(self):
        return {
            'id': self.id,
            'organization': self.organization,
            'initiative': self.initiative,
            'description': self.description,
            'project_objectives': self.project_objectives,
            'skills_required': self.skills_required,
            'experience_level': self.experience_level,
            'priority_level': self.priority_level,
            'complexity': self.complexity,
            'category': self.category,
            'volunteers_needed': self.volunteers_needed,
            'volunteers_assigned': self.volunteers_assigned,
            'capacity_remaining': self.capacity_remaining,
            'created_at': self.created_at.isoformat() if self.created_at else None
        }