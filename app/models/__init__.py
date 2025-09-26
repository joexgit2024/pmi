# Import all models for easy access
from .file_upload import FileUpload
from .registration import Registration
from .charity import Charity
from .matching import MatchingResult, MatchingBatch
from .email_tracking import EmailTracking, EmailBatch

# Import the database instance
from app import db

__all__ = [
    'db',
    'FileUpload',
    'Registration', 
    'Charity',
    'MatchingResult',
    'MatchingBatch',
    'EmailTracking',
    'EmailBatch'
]