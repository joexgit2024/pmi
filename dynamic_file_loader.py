"""
Dynamic File Loader - Helper functions to find latest files automatically
"""

import os
import glob
from datetime import datetime


def find_latest_registration_file(input_dir="input"):
    """
    Find the latest PMDoS registration file in the input directory.
    Looks for files containing 'PMDoS' and 'Registration' and 'Responses'.
    Returns the most recently modified file.
    """
    # Pattern to match PMDoS registration files
    patterns = [
        "*PMDoS*Registration*Responses*.xlsx",
        "*Registration*Responses*.xlsx", 
        "*PMI Sydney*Registration*.xlsx"
    ]
    
    files_found = []
    
    for pattern in patterns:
        search_path = os.path.join(input_dir, pattern)
        matching_files = glob.glob(search_path)
        files_found.extend(matching_files)
    
    if not files_found:
        return None
        
    # Remove duplicates and get the most recently modified file
    files_found = list(set(files_found))
    latest_file = max(files_found, key=os.path.getmtime)
    
    return latest_file


def find_latest_charity_file(input_dir="input"):
    """
    Find the latest charity information file in the input directory.
    Looks for files containing 'Charities' and 'Responses'.
    Returns the most recently modified file.
    """
    patterns = [
        "*Charities*Information*Responses*.xlsx",
        "*Charities*Project*Responses*.xlsx",
        "*Charity*Information*.xlsx"
    ]
    
    files_found = []
    
    for pattern in patterns:
        search_path = os.path.join(input_dir, pattern)
        matching_files = glob.glob(search_path)
        files_found.extend(matching_files)
    
    if not files_found:
        return None
        
    # Remove duplicates and get the most recently modified file
    files_found = list(set(files_found))
    latest_file = max(files_found, key=os.path.getmtime)
    
    return latest_file


def get_latest_input_files(input_dir="input"):
    """
    Get both latest registration and charity files.
    Returns tuple (registration_file, charity_file) or (None, None) if not found.
    """
    registration_file = find_latest_registration_file(input_dir)
    charity_file = find_latest_charity_file(input_dir)
    
    return registration_file, charity_file


def validate_dynamic_input_files(input_dir="input"):
    """
    Validate that we can find the required input files dynamically.
    Returns True if both files found, False otherwise.
    """
    reg_file, charity_file = get_latest_input_files(input_dir)
    
    if not reg_file:
        print(f"ERROR: No PMDoS registration file found in {input_dir}/")
        print("Looking for files containing: 'PMDoS', 'Registration', 'Responses'")
        return False
        
    if not charity_file:
        print(f"ERROR: No charity information file found in {input_dir}/")
        print("Looking for files containing: 'Charities', 'Information', 'Responses'")
        return False
    
    print(f"✓ Found registration file: {os.path.basename(reg_file)}")
    print(f"✓ Found charity file: {os.path.basename(charity_file)}")
    
    return True


if __name__ == "__main__":
    # Test the dynamic file finding
    print("Testing dynamic file detection...")
    reg_file, charity_file = get_latest_input_files()
    
    if reg_file:
        print(f"Registration file: {reg_file}")
        mod_time = datetime.fromtimestamp(os.path.getmtime(reg_file))
        print(f"Modified: {mod_time}")
    else:
        print("No registration file found")
        
    if charity_file:
        print(f"Charity file: {charity_file}")
        mod_time = datetime.fromtimestamp(os.path.getmtime(charity_file))
        print(f"Modified: {mod_time}")
    else:
        print("No charity file found")