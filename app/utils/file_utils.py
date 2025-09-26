import os
import mimetypes
import pandas as pd
from werkzeug.datastructures import FileStorage

def validate_file(file):
    """
    Validate uploaded file for type, size, and content.
    
    Args:
        file: Werkzeug FileStorage object
        
    Returns:
        dict: Validation result with 'valid' boolean and 'message'
    """
    if not file or not file.filename:
        return {
            'valid': False,
            'message': 'No file selected'
        }
    
    # Check file extension
    allowed_extensions = {'xlsx', 'xls', 'csv'}
    file_ext = file.filename.rsplit('.', 1)[1].lower() if '.' in file.filename else ''
    
    if file_ext not in allowed_extensions:
        return {
            'valid': False,
            'message': f'Invalid file type. Allowed types: {", ".join(allowed_extensions)}'
        }
    
    # Check file size (50MB limit)
    max_size = 50 * 1024 * 1024  # 50MB in bytes
    if hasattr(file, 'content_length') and file.content_length:
        if file.content_length > max_size:
            return {
                'valid': False,
                'message': 'File size exceeds 50MB limit'
            }
    
    # Basic MIME type validation
    mime_type, _ = mimetypes.guess_type(file.filename)
    allowed_mimes = [
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',  # xlsx
        'application/vnd.ms-excel',  # xls
        'text/csv',  # csv
        'application/csv'  # csv alternative
    ]
    
    if mime_type and mime_type not in allowed_mimes:
        # Don't fail on MIME type alone, as it can be unreliable
        pass
    
    return {
        'valid': True,
        'message': 'File validation passed'
    }

def get_file_type(file_path):
    """
    Get detailed file type information.
    
    Args:
        file_path: Path to the file
        
    Returns:
        dict: File information including type, size, etc.
    """
    try:
        file_stats = os.stat(file_path)
        file_ext = file_path.rsplit('.', 1)[1].lower() if '.' in file_path else ''
        mime_type, _ = mimetypes.guess_type(file_path)
        
        return {
            'extension': file_ext,
            'mime_type': mime_type,
            'size_bytes': file_stats.st_size,
            'size_mb': round(file_stats.st_size / (1024 * 1024), 2),
            'is_excel': file_ext in ['xlsx', 'xls'],
            'is_csv': file_ext == 'csv'
        }
    except Exception as e:
        return {
            'extension': '',
            'mime_type': None,
            'size_bytes': 0,
            'size_mb': 0,
            'is_excel': False,
            'is_csv': False,
            'error': str(e)
        }

def process_excel_file(file_path):
    """
    Process Excel file and return DataFrame.
    
    Args:
        file_path: Path to the Excel file
        
    Returns:
        pandas.DataFrame or None
    """
    try:
        # Try to read Excel file
        # First, try to detect if there are multiple sheets
        xl_file = pd.ExcelFile(file_path)
        sheet_names = xl_file.sheet_names
        
        if len(sheet_names) == 1:
            # Single sheet
            df = pd.read_excel(file_path, sheet_name=0)
        else:
            # Multiple sheets - try to find the data sheet
            data_sheet = None
            
            # Common data sheet names
            data_sheet_names = ['data', 'registrations', 'participants', 'charities', 'sheet1']
            
            for sheet_name in sheet_names:
                if sheet_name.lower() in data_sheet_names:
                    data_sheet = sheet_name
                    break
            
            # If no common name found, use the first sheet
            if not data_sheet:
                data_sheet = sheet_names[0]
            
            df = pd.read_excel(file_path, sheet_name=data_sheet)
        
        # Clean up the DataFrame
        df = clean_dataframe(df)
        
        return df
        
    except Exception as e:
        print(f"Error processing Excel file {file_path}: {str(e)}")
        return None

def process_csv_file(file_path):
    """
    Process CSV file and return DataFrame.
    
    Args:
        file_path: Path to the CSV file
        
    Returns:
        pandas.DataFrame or None
    """
    try:
        # Try different encodings
        encodings = ['utf-8', 'latin-1', 'cp1252', 'iso-8859-1']
        
        for encoding in encodings:
            try:
                # Try different delimiters
                for delimiter in [',', ';', '\t']:
                    try:
                        df = pd.read_csv(file_path, encoding=encoding, delimiter=delimiter)
                        
                        # Check if we got meaningful data
                        if len(df.columns) > 1 and len(df) > 0:
                            # Clean up the DataFrame
                            df = clean_dataframe(df)
                            return df
                    except:
                        continue
            except:
                continue
        
        # If all else fails, try with default settings
        df = pd.read_csv(file_path)
        df = clean_dataframe(df)
        return df
        
    except Exception as e:
        print(f"Error processing CSV file {file_path}: {str(e)}")
        return None

def clean_dataframe(df):
    """
    Clean and standardize DataFrame.
    
    Args:
        df: pandas.DataFrame
        
    Returns:
        pandas.DataFrame: Cleaned DataFrame
    """
    if df is None or df.empty:
        return df
    
    try:
        # Remove completely empty rows and columns
        df = df.dropna(how='all')
        df = df.dropna(axis=1, how='all')
        
        # Clean column names
        df.columns = df.columns.astype(str)
        df.columns = [col.strip() for col in df.columns]
        
        # Remove unnamed columns (often from Excel index columns)
        unnamed_cols = [col for col in df.columns if col.lower().startswith('unnamed')]
        df = df.drop(columns=unnamed_cols)
        
        # Reset index
        df = df.reset_index(drop=True)
        
        return df
        
    except Exception as e:
        print(f"Error cleaning DataFrame: {str(e)}")
        return df

def sanitize_filename(filename):
    """
    Sanitize filename for safe storage.
    
    Args:
        filename: Original filename
        
    Returns:
        str: Sanitized filename
    """
    # Remove or replace dangerous characters
    import string
    import re
    
    # Keep only alphanumeric, dots, hyphens, and underscores
    valid_chars = f"-_.() {string.ascii_letters}{string.digits}"
    sanitized = ''.join(c for c in filename if c in valid_chars)
    
    # Remove multiple consecutive spaces/dots
    sanitized = re.sub(r'[.\s]+', '.', sanitized)
    sanitized = re.sub(r'[-_\s]+', '_', sanitized)
    
    # Ensure it's not too long
    if len(sanitized) > 100:
        name, ext = os.path.splitext(sanitized)
        sanitized = name[:96] + ext
    
    return sanitized

def format_file_size(size_bytes):
    """
    Format file size in human readable format.
    
    Args:
        size_bytes: File size in bytes
        
    Returns:
        str: Formatted size string
    """
    if size_bytes == 0:
        return "0 B"
    
    size_names = ["B", "KB", "MB", "GB", "TB"]
    import math
    i = int(math.floor(math.log(size_bytes, 1024)))
    p = math.pow(1024, i)
    s = round(size_bytes / p, 2)
    
    return f"{s} {size_names[i]}"

def validate_email(email):
    """
    Basic email validation.
    
    Args:
        email: Email address string
        
    Returns:
        bool: True if email appears valid
    """
    import re
    
    if not email or not isinstance(email, str):
        return False
    
    # Basic email pattern
    pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    return bool(re.match(pattern, email.strip()))

def validate_linkedin_url(url):
    """
    Validate LinkedIn URL format.
    
    Args:
        url: LinkedIn URL string
        
    Returns:
        bool: True if URL appears to be a valid LinkedIn profile
    """
    if not url or not isinstance(url, str):
        return False
    
    url = url.strip().lower()
    
    # Common LinkedIn URL patterns
    linkedin_patterns = [
        r'linkedin\.com/in/',
        r'www\.linkedin\.com/in/',
        r'https://linkedin\.com/in/',
        r'https://www\.linkedin\.com/in/',
        r'http://linkedin\.com/in/',
        r'http://www\.linkedin\.com/in/'
    ]
    
    import re
    for pattern in linkedin_patterns:
        if re.search(pattern, url):
            return True
    
    return False

def extract_name_components(full_name):
    """
    Extract first and last name from full name.
    
    Args:
        full_name: Full name string
        
    Returns:
        dict: Dictionary with 'first_name' and 'last_name'
    """
    if not full_name or not isinstance(full_name, str):
        return {'first_name': '', 'last_name': ''}
    
    name_parts = full_name.strip().split()
    
    if len(name_parts) == 0:
        return {'first_name': '', 'last_name': ''}
    elif len(name_parts) == 1:
        return {'first_name': name_parts[0], 'last_name': ''}
    else:
        return {
            'first_name': name_parts[0],
            'last_name': ' '.join(name_parts[1:])
        }

def create_backup_filename(original_filename):
    """
    Create a backup filename with timestamp.
    
    Args:
        original_filename: Original filename
        
    Returns:
        str: Backup filename with timestamp
    """
    from datetime import datetime
    
    name, ext = os.path.splitext(original_filename)
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    
    return f"{name}_backup_{timestamp}{ext}"

def ensure_directory_exists(directory_path):
    """
    Ensure directory exists, create if it doesn't.
    
    Args:
        directory_path: Path to directory
        
    Returns:
        bool: True if directory exists or was created successfully
    """
    try:
        os.makedirs(directory_path, exist_ok=True)
        return True
    except Exception as e:
        print(f"Error creating directory {directory_path}: {str(e)}")
        return False

def allowed_file(filename):
    """
    Check if file has allowed extension.
    
    Args:
        filename: Name of the file
        
    Returns:
        bool: True if file extension is allowed
    """
    allowed_extensions = {'xlsx', 'xls', 'csv'}
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in allowed_extensions

def validate_excel_file(file_path):
    """
    Validate file structure and content (Excel or CSV).
    
    Args:
        file_path: Path to the file
        
    Returns:
        dict: Validation result with 'valid' boolean and 'message'
    """
    try:
        # Determine file type and read accordingly
        if file_path.lower().endswith('.csv'):
            df = process_csv_file(file_path)
        else:
            df = process_excel_file(file_path)
        
        if df is None:
            return {
                'valid': False,
                'message': 'Could not read file'
            }
        
        if df.empty:
            return {
                'valid': False,
                'message': 'File is empty'
            }
        
        # Check if there are any columns
        if len(df.columns) == 0:
            return {
                'valid': False,
                'message': 'File has no columns'
            }
        
        return {
            'valid': True,
            'message': 'File validation passed',
            'rows': len(df),
            'columns': len(df.columns)
        }
        
    except Exception as e:
        return {
            'valid': False,
            'message': f'File validation error: {str(e)}'
        }