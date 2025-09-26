from app import db
from datetime import datetime

class FileUpload(db.Model):
    """Model for tracking uploaded files."""
    
    __tablename__ = 'file_uploads'
    
    id = db.Column(db.Integer, primary_key=True)
    filename = db.Column(db.String(255), nullable=False)
    original_filename = db.Column(db.String(255), nullable=False)
    file_type = db.Column(db.String(20), nullable=False)  # 'registration' or 'charity'
    file_path = db.Column(db.String(500), nullable=False)
    file_size = db.Column(db.Integer)
    upload_date = db.Column(db.DateTime, default=datetime.utcnow)
    status = db.Column(db.String(20), default='uploaded')  # uploaded, processing, processed, error
    rows_count = db.Column(db.Integer)
    error_message = db.Column(db.Text)
    
    # Relationships
    registrations = db.relationship('Registration', backref='file_upload', lazy='dynamic')
    charities = db.relationship('Charity', backref='file_upload', lazy='dynamic')
    
    def __repr__(self):
        return f'<FileUpload {self.filename}>'
    
    def to_dict(self):
        return {
            'id': self.id,
            'filename': self.filename,
            'original_filename': self.original_filename,
            'file_type': self.file_type,
            'file_size': self.file_size,
            'upload_date': self.upload_date.isoformat() if self.upload_date else None,
            'status': self.status,
            'rows_count': self.rows_count,
            'error_message': self.error_message
        }