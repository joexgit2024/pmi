import os
import pandas as pd
from datetime import datetime
from werkzeug.utils import secure_filename
from app.models.file_upload import FileUpload
from app.models.registration import Registration
from app.models.charity import Charity
from app.utils.file_utils import validate_file, get_file_type, process_excel_file, process_csv_file
from app import db

class FileService:
    """Service class for handling file operations and processing."""
    
    ALLOWED_EXTENSIONS = {'xlsx', 'xls', 'csv'}
    MAX_FILE_SIZE = 50 * 1024 * 1024  # 50MB
    
    def __init__(self):
        self.upload_folder = os.path.join(os.getcwd(), 'uploads')
        os.makedirs(self.upload_folder, exist_ok=True)
    
    def save_uploaded_file(self, file, file_type='registration'):
        """
        Save uploaded file and create database record.
        
        Args:
            file: Werkzeug FileStorage object
            file_type: Type of file ('registration' or 'charity')
            
        Returns:
            dict: Result with success status and file info
        """
        try:
            # Validate file
            validation_result = validate_file(file)
            if not validation_result['valid']:
                return {
                    'success': False,
                    'message': validation_result['message']
                }
            
            # Generate secure filename
            filename = secure_filename(file.filename)
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            filename = f"{timestamp}_{filename}"
            
            # Save file to disk
            file_path = os.path.join(self.upload_folder, filename)
            file.save(file_path)
            
            # Get file info
            file_info = get_file_type(file_path)
            
            # Create database record
            file_upload = FileUpload(
                original_filename=file.filename,
                stored_filename=filename,
                file_path=file_path,
                file_type=file_type,
                file_size=os.path.getsize(file_path),
                mime_type=file.content_type,
                status='uploaded'
            )
            
            db.session.add(file_upload)
            db.session.commit()
            
            return {
                'success': True,
                'file_id': file_upload.id,
                'filename': filename,
                'original_filename': file.filename,
                'file_type': file_type,
                'file_info': file_info
            }
            
        except Exception as e:
            db.session.rollback()
            return {
                'success': False,
                'message': f"Error saving file: {str(e)}"
            }
    
    def process_file(self, file_id):
        """
        Process uploaded file and extract data.
        
        Args:
            file_id: ID of the uploaded file
            
        Returns:
            dict: Processing result
        """
        try:
            file_upload = FileUpload.query.get(file_id)
            if not file_upload:
                return {
                    'success': False,
                    'message': 'File not found'
                }
            
            # Update status to processing
            file_upload.status = 'processing'
            db.session.commit()
            
            # Process based on file type
            if file_upload.file_type == 'registration':
                result = self._process_registration_file(file_upload)
            elif file_upload.file_type == 'charity':
                result = self._process_charity_file(file_upload)
            else:
                result = {
                    'success': False,
                    'message': 'Unknown file type'
                }
            
            # Update file status
            file_upload.status = 'completed' if result['success'] else 'error'
            if not result['success']:
                file_upload.error_message = result.get('message', 'Processing failed')
            
            db.session.commit()
            
            return result
            
        except Exception as e:
            if 'file_upload' in locals():
                file_upload.status = 'error'
                file_upload.error_message = str(e)
                db.session.commit()
            
            return {
                'success': False,
                'message': f"Error processing file: {str(e)}"
            }
    
    def _process_registration_file(self, file_upload):
        """Process registration file and extract participant data."""
        try:
            # Read file based on extension
            file_path = file_upload.file_path
            
            if file_path.endswith('.csv'):
                df = process_csv_file(file_path)
            else:
                df = process_excel_file(file_path)
            
            if df is None or df.empty:
                return {
                    'success': False,
                    'message': 'File is empty or could not be read'
                }
            
            # Validate required columns
            required_columns = ['name', 'email']
            missing_columns = []
            
            # Try to find columns with similar names
            column_mapping = self._map_columns(df.columns.tolist(), required_columns)
            
            if not column_mapping:
                return {
                    'success': False,
                    'message': f'Required columns not found. Expected: {required_columns}'
                }
            
            # Process each row
            processed_count = 0
            error_count = 0
            
            for index, row in df.iterrows():
                try:
                    # Extract data using column mapping
                    registration_data = {}
                    
                    for required_col, actual_col in column_mapping.items():
                        if actual_col:
                            registration_data[required_col] = str(row[actual_col]).strip() if pd.notna(row[actual_col]) else ''
                    
                    # Skip rows with missing required data
                    if not registration_data.get('name') or not registration_data.get('email'):
                        continue
                    
                    # Check if registration already exists
                    existing = Registration.query.filter_by(
                        email=registration_data['email']
                    ).first()
                    
                    if existing:
                        # Update existing record
                        for key, value in registration_data.items():
                            setattr(existing, key, value)
                        existing.file_upload_id = file_upload.id
                    else:
                        # Create new registration
                        registration = Registration(
                            file_upload_id=file_upload.id,
                            **registration_data
                        )
                        db.session.add(registration)
                    
                    processed_count += 1
                    
                except Exception as row_error:
                    error_count += 1
                    print(f"Error processing row {index}: {row_error}")
                    continue
            
            db.session.commit()
            
            return {
                'success': True,
                'processed_count': processed_count,
                'error_count': error_count,
                'message': f'Processed {processed_count} registrations'
            }
            
        except Exception as e:
            db.session.rollback()
            return {
                'success': False,
                'message': f"Error processing registration file: {str(e)}"
            }
    
    def _process_charity_file(self, file_upload):
        """Process charity file and extract charity data."""
        try:
            # Read file based on extension
            file_path = file_upload.file_path
            
            if file_path.endswith('.csv'):
                df = process_csv_file(file_path)
            else:
                df = process_excel_file(file_path)
            
            if df is None or df.empty:
                return {
                    'success': False,
                    'message': 'File is empty or could not be read'
                }
            
            # Validate required columns
            required_columns = ['charity_name', 'problem_statement']
            column_mapping = self._map_columns(df.columns.tolist(), required_columns)
            
            if not column_mapping:
                return {
                    'success': False,
                    'message': f'Required columns not found. Expected: {required_columns}'
                }
            
            # Process each row
            processed_count = 0
            error_count = 0
            
            for index, row in df.iterrows():
                try:
                    # Extract data using column mapping
                    charity_data = {}
                    
                    for required_col, actual_col in column_mapping.items():
                        if actual_col:
                            charity_data[required_col] = str(row[actual_col]).strip() if pd.notna(row[actual_col]) else ''
                    
                    # Skip rows with missing required data
                    if not charity_data.get('charity_name'):
                        continue
                    
                    # Check if charity already exists
                    existing = Charity.query.filter_by(
                        charity_name=charity_data['charity_name']
                    ).first()
                    
                    if existing:
                        # Update existing record
                        for key, value in charity_data.items():
                            setattr(existing, key, value)
                        existing.file_upload_id = file_upload.id
                    else:
                        # Create new charity
                        charity = Charity(
                            file_upload_id=file_upload.id,
                            **charity_data
                        )
                        db.session.add(charity)
                    
                    processed_count += 1
                    
                except Exception as row_error:
                    error_count += 1
                    print(f"Error processing row {index}: {row_error}")
                    continue
            
            db.session.commit()
            
            return {
                'success': True,
                'processed_count': processed_count,
                'error_count': error_count,
                'message': f'Processed {processed_count} charities'
            }
            
        except Exception as e:
            db.session.rollback()
            return {
                'success': False,
                'message': f"Error processing charity file: {str(e)}"
            }
    
    def _map_columns(self, actual_columns, required_columns):
        """Map actual column names to required column names."""
        column_mapping = {}
        
        # Common column name variations
        name_variations = ['name', 'full_name', 'participant_name', 'attendee_name', 'first_name']
        email_variations = ['email', 'email_address', 'e_mail', 'mail']
        charity_name_variations = ['charity_name', 'charity', 'organization', 'org_name', 'ngo_name']
        problem_statement_variations = ['problem_statement', 'problem', 'description', 'statement', 'issue']
        
        variations_map = {
            'name': name_variations,
            'email': email_variations,
            'charity_name': charity_name_variations,
            'problem_statement': problem_statement_variations
        }
        
        # Convert to lowercase for comparison
        actual_columns_lower = [col.lower().strip() for col in actual_columns]
        
        for required_col in required_columns:
            column_mapping[required_col] = None
            
            if required_col in variations_map:
                for variation in variations_map[required_col]:
                    if variation.lower() in actual_columns_lower:
                        # Find the original column name
                        original_index = actual_columns_lower.index(variation.lower())
                        column_mapping[required_col] = actual_columns[original_index]
                        break
        
        # Return mapping only if all required columns are found
        if all(column_mapping.values()):
            return column_mapping
        
        return None
    
    def get_recent_uploads(self, limit=10):
        """Get recent file uploads."""
        return FileUpload.query.order_by(
            FileUpload.created_at.desc()
        ).limit(limit).all()
    
    def get_upload_statistics(self):
        """Get upload statistics."""
        total_files = FileUpload.query.count()
        successful_uploads = FileUpload.query.filter_by(status='completed').count()
        registration_files = FileUpload.query.filter_by(file_type='registration').count()
        charity_files = FileUpload.query.filter_by(file_type='charity').count()
        
        return {
            'total_files': total_files,
            'successful_uploads': successful_uploads,
            'registration_files': registration_files,
            'charity_files': charity_files
        }
    
    def delete_file(self, file_id):
        """Delete uploaded file and its data."""
        try:
            file_upload = FileUpload.query.get(file_id)
            if not file_upload:
                return {
                    'success': False,
                    'message': 'File not found'
                }
            
            # Delete physical file
            if os.path.exists(file_upload.file_path):
                os.remove(file_upload.file_path)
            
            # Delete associated data
            if file_upload.file_type == 'registration':
                Registration.query.filter_by(file_upload_id=file_id).delete()
            elif file_upload.file_type == 'charity':
                Charity.query.filter_by(file_upload_id=file_id).delete()
            
            # Delete database record
            db.session.delete(file_upload)
            db.session.commit()
            
            return {
                'success': True,
                'message': 'File deleted successfully'
            }
            
        except Exception as e:
            db.session.rollback()
            return {
                'success': False,
                'message': f"Error deleting file: {str(e)}"
            }