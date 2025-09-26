from app import db
from datetime import datetime

class EmailTracking(db.Model):
    """Model for tracking email acknowledgments sent to registrants."""
    
    __tablename__ = 'email_tracking'
    
    id = db.Column(db.Integer, primary_key=True)
    registration_id = db.Column(db.Integer, db.ForeignKey('registrations.id'), nullable=False)
    
    # Email Details
    email_address = db.Column(db.String(255), nullable=False)
    recipient_name = db.Column(db.String(255), nullable=False)
    
    # Batch Information
    batch_id = db.Column(db.String(50), nullable=False)
    batch_date = db.Column(db.Date, nullable=False)
    
    # File Information
    draft_file_path = db.Column(db.String(500))
    draft_filename = db.Column(db.String(255))
    
    # Status
    status = db.Column(db.String(20), default='drafted')  # drafted, sent, failed, bounced
    sent_at = db.Column(db.DateTime)
    error_message = db.Column(db.Text)
    
    # Metadata
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)
    
    def __repr__(self):
        return f'<EmailTracking {self.recipient_name} - {self.email_address}>'
    
    def to_dict(self):
        return {
            'id': self.id,
            'registration_id': self.registration_id,
            'email_address': self.email_address,
            'recipient_name': self.recipient_name,
            'batch_id': self.batch_id,
            'batch_date': self.batch_date.isoformat() if self.batch_date else None,
            'draft_filename': self.draft_filename,
            'status': self.status,
            'sent_at': self.sent_at.isoformat() if self.sent_at else None,
            'created_at': self.created_at.isoformat() if self.created_at else None
        }


class EmailBatch(db.Model):
    """Model for tracking email batch operations."""
    
    __tablename__ = 'email_batches'
    
    id = db.Column(db.Integer, primary_key=True)
    batch_id = db.Column(db.String(50), unique=True, nullable=False)
    
    # Batch Information
    registration_file_id = db.Column(db.Integer, db.ForeignKey('file_uploads.id'))
    batch_date = db.Column(db.Date, nullable=False)
    batch_folder = db.Column(db.String(500))
    
    # Email Count
    total_emails = db.Column(db.Integer, default=0)
    emails_drafted = db.Column(db.Integer, default=0)
    emails_sent = db.Column(db.Integer, default=0)
    emails_failed = db.Column(db.Integer, default=0)
    
    # Processing Status
    status = db.Column(db.String(20), default='processing')  # processing, completed, failed
    progress_percentage = db.Column(db.Integer, default=0)
    error_message = db.Column(db.Text)
    
    # Template Information
    template_used = db.Column(db.String(255))
    template_version = db.Column(db.String(20))
    
    # Metadata
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    completed_at = db.Column(db.DateTime)
    
    def __repr__(self):
        return f'<EmailBatch {self.batch_id}>'
    
    @property
    def completion_percentage(self):
        if self.total_emails == 0:
            return 0
        return int((self.emails_drafted / self.total_emails) * 100)
    
    def to_dict(self):
        return {
            'id': self.id,
            'batch_id': self.batch_id,
            'batch_date': self.batch_date.isoformat() if self.batch_date else None,
            'total_emails': self.total_emails,
            'emails_drafted': self.emails_drafted,
            'emails_sent': self.emails_sent,
            'emails_failed': self.emails_failed,
            'status': self.status,
            'progress_percentage': self.progress_percentage,
            'completion_percentage': self.completion_percentage,
            'created_at': self.created_at.isoformat() if self.created_at else None,
            'completed_at': self.completed_at.isoformat() if self.completed_at else None
        }