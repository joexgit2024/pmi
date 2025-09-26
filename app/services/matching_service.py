"""
Matching Service for PMI Web Application
=======================================

This service integrates the run_complete_analysis.py functionality
into the web application, providing charity-PMP matching capabilities.
"""

import os
import sys
import subprocess
from datetime import datetime
import pandas as pd
from pathlib import Path

# Add the project root to Python path to import our analysis modules
project_root = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, project_root)

from app import db
from app.models import MatchingBatch, MatchingResult, Registration, Charity

# Import the dynamic file loader from the root directory
try:
    from dynamic_file_loader import get_latest_input_files
except ImportError:
    # Alternative import path in case the above doesn't work
    import importlib.util
    spec = importlib.util.spec_from_file_location("dynamic_file_loader", 
                                                os.path.join(project_root, "dynamic_file_loader.py"))
    dynamic_file_loader = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(dynamic_file_loader)
    get_latest_input_files = dynamic_file_loader.get_latest_input_files


class MatchingService:
    """Service for running PMP-Charity matching analysis."""
    
    def __init__(self):
        self.project_root = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
        self.output_dir = os.path.join(self.project_root, "Output")
        
        # Ensure Output directory exists
        os.makedirs(self.output_dir, exist_ok=True)
    
    def run_matching(self, use_flexible=False):
        """
        Run the complete PMP-Charity matching analysis.
        
        Args:
            use_flexible (bool): If True, uses flexible assignment (all PMPs assigned)
                               If False, uses standard assignment (2 PMPs per charity)
        
        Returns:
            dict: Result of the matching process
        """
        try:
            # Step 1: Validate input files
            validation_result = self._validate_input_files()
            if not validation_result['success']:
                return validation_result
            
            # Step 2: Run the complete analysis
            analysis_result = self._run_complete_analysis(use_flexible)
            if not analysis_result['success']:
                return analysis_result
            
            # Step 3: Import results to database
            import_result = self._import_results_to_database()
            if not import_result['success']:
                return import_result
            
            # Step 4: Return success with details
            return {
                'success': True,
                'message': 'Matching completed successfully',
                'matched_count': import_result['matched_count'],
                'total_pmps': analysis_result['total_pmps'],
                'total_charities': analysis_result['total_charities'],
                'output_files': analysis_result['output_files'],
                'batch_id': import_result['batch_id']
            }
            
        except Exception as e:
            return {
                'success': False,
                'message': f'Matching failed: {str(e)}'
            }
    
    def _validate_input_files(self):
        """Validate that required input files exist."""
        try:
            reg_file, charity_file = get_latest_input_files()
            
            if not reg_file:
                return {
                    'success': False,
                    'message': 'No PMP registration file found in input/ directory'
                }
            
            if not charity_file:
                return {
                    'success': False,
                    'message': 'No charity information file found in input/ directory'
                }
            
            if not os.path.exists(reg_file):
                return {
                    'success': False,
                    'message': f'Registration file not found: {reg_file}'
                }
            
            if not os.path.exists(charity_file):
                return {
                    'success': False,
                    'message': f'Charity file not found: {charity_file}'
                }
            
            return {
                'success': True,
                'registration_file': reg_file,
                'charity_file': charity_file
            }
            
        except Exception as e:
            return {
                'success': False,
                'message': f'Error validating input files: {str(e)}'
            }
    
    def _run_complete_analysis(self, use_flexible=False):
        """Run the complete analysis using run_complete_analysis.py"""
        try:
            # Change to project directory
            original_cwd = os.getcwd()
            os.chdir(self.project_root)
            
            # Prepare command
            cmd = [sys.executable, 'run_complete_analysis.py']
            if use_flexible:
                cmd.append('--flexible')
            
            # Run the analysis
            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                cwd=self.project_root
            )
            
            # Restore original directory
            os.chdir(original_cwd)
            
            if result.returncode != 0:
                # Sanitize potential unicode issues for web JSON response
                def _sanitize(text):
                    try:
                        return text.encode('utf-8', errors='ignore').decode('utf-8')
                    except Exception:
                        return text[:1000]
                return {
                    'success': False,
                    'message': 'Analysis failed',
                    'stderr': _sanitize(result.stderr),
                    'stdout': _sanitize(result.stdout)
                }
            
            # Parse the output for statistics
            output_lines = result.stdout.split('\n')
            total_pmps = 0
            total_charities = 0
            
            for line in output_lines:
                if 'Total PMPs:' in line:
                    try:
                        total_pmps = int(line.split('Total PMPs:')[1].strip())
                    except:
                        pass
                elif 'Total Charity Projects:' in line:
                    try:
                        total_charities = int(line.split('Total Charity Projects:')[1].strip())
                    except:
                        pass
            
            # Check for expected output files
            expected_files = [
                'LinkedIn_Analysis_Report.xlsx',
                'PMI_PMP_Charity_Matching_Results_Enhanced.xlsx',
                'Matching_Summary.csv',
                'Analysis_Summary.txt'
            ]
            
            output_files = []
            for file in expected_files:
                file_path = os.path.join(self.output_dir, file)
                if os.path.exists(file_path):
                    output_files.append(file)
            
            return {
                'success': True,
                'total_pmps': total_pmps,
                'total_charities': total_charities,
                'output_files': output_files,
                'output': result.stdout
            }
            
        except Exception as e:
            # Restore original directory on error
            try:
                os.chdir(original_cwd)
            except:
                pass
            
            return {
                'success': False,
                'message': f'Error running analysis: {str(e)}'
            }
    
    def _import_results_to_database(self):
        """Import matching results to database."""
        try:
            # Path to the enhanced matching results file
            results_file = os.path.join(self.output_dir, 'PMI_PMP_Charity_Matching_Results_Enhanced.xlsx')
            
            if not os.path.exists(results_file):
                return {
                    'success': False,
                    'message': 'Enhanced matching results file not found'
                }
            
            # Read the matching results
            matching_df = pd.read_excel(results_file, sheet_name='Enhanced_Matching_Summary')
            
            # Create a new matching batch
            batch_id_str = f"web_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
            batch = MatchingBatch(
                batch_id=batch_id_str,
                algorithm_version="enhanced_v2",
                matching_type="standard",
                total_matches=len(matching_df),
                status="running",
                excel_report_path="Output/PMI_PMP_Charity_Matching_Results_Enhanced.xlsx",
                csv_summary_path="Output/Matching_Summary.csv"
            )
            
            db.session.add(batch)
            db.session.flush()  # Get the batch ID
            
            # Import matching results
            matched_count = 0
            
            for _, row in matching_df.iterrows():
                try:
                    # Parse name from PMP_Name column
                    pmp_full_name = str(row.get('PMP_Name', '')).strip()
                    name_parts = pmp_full_name.split(' ', 1)
                    first_name = name_parts[0] if len(name_parts) > 0 else ''
                    last_name = name_parts[1] if len(name_parts) > 1 else ''
                    
                    # Use LinkedIn URL as unique identifier since email might not be in summary
                    linkedin_url = row.get('LinkedIn_URL', '')
                    
                    registration = Registration.query.filter_by(linkedin_url=linkedin_url).first()
                    if not registration and linkedin_url:
                        registration = Registration(
                            first_name=first_name,
                            last_name=last_name,
                            email=f"{first_name.lower()}.{last_name.lower()}@temp.com",  # Temporary email
                            linkedin_url=linkedin_url,
                            job_title=str(row.get('PMP_Job_Title', '')),
                            company=str(row.get('PMP_Company', '')),
                            experience_years=str(row.get('PMP_Experience', '')),
                            areas_of_interest=str(row.get('PMP_Top_Skills', '')),
                            linkedin_quality_score=float(row.get('LinkedIn_Quality', 0)),
                            profile_completeness_score=float(row.get('Profile_Completeness', 0)),
                            overall_score=float(row.get('Overall_PMP_Rating', 0)),
                            file_upload_id=1,  # Default file upload ID
                            created_at=datetime.now()
                        )
                        db.session.add(registration)
                        db.session.flush()
                    
                    # Find or create charity record
                    charity_org = row.get('Charity_Organization', '')
                    charity = Charity.query.filter_by(organization=charity_org).first()
                    if not charity:
                        charity = Charity(
                            organization=charity_org,
                            initiative=str(row.get('Charity_Initiative', '')),
                            description=str(row.get('Project_Description', '')),
                            priority_level=str(row.get('Project_Priority', '')),
                            complexity=str(row.get('Project_Complexity', '')),
                            skills_required=str(row.get('Required_Skills', '')),
                            file_upload_id=1,  # Default file upload ID
                            created_at=datetime.now()
                        )
                        db.session.add(charity)
                        db.session.flush()
                    
                    # Create matching result
                    matching_result = MatchingResult(
                        batch_id=batch.id,
                        registration_id=registration.id,
                        charity_id=charity.id,
                        match_score=float(row.get('Match_Score', 0)),
                        linkedin_quality=float(row.get('LinkedIn_Quality', 0)),
                        skills_match=float(row.get('Match_Score', 0)) / 100.0,  # Normalize to 0-1 range
                        matching_algorithm='enhanced_v2',
                        assignment_rank=1 if 'PMP 1' in str(row.get('PMP_Role', '')) else 2,
                        created_at=datetime.now()
                    )
                    
                    db.session.add(matching_result)
                    matched_count += 1
                    
                except Exception as e:
                    print(f"Error importing row: {e}")
                    continue
            
            # Update batch with final count and completion
            batch.total_matches = matched_count
            batch.status = "completed"
            batch.completed_at = datetime.now()
            batch.progress_percentage = 100
            
            # Commit all changes
            db.session.commit()
            
            return {
                'success': True,
                'matched_count': matched_count,
                'batch_id': batch.id
            }
            
        except Exception as e:
            db.session.rollback()
            return {
                'success': False,
                'message': f'Error importing results to database: {str(e)}'
            }
    
    def get_latest_results(self):
        """Get the latest matching results."""
        try:
            latest_batch = MatchingBatch.query.order_by(MatchingBatch.created_at.desc()).first()
            
            if not latest_batch:
                return {
                    'success': False,
                    'message': 'No matching batches found'
                }
            
            # Get matching results for this batch
            results = MatchingResult.query.filter_by(batch_id=latest_batch.id).all()
            
            results_data = []
            for result in results:
                results_data.append({
                    'pmp_name': f"{result.registration.first_name} {result.registration.last_name}",
                    'pmp_email': result.registration.email,
                    'charity_name': result.charity.name,
                    'match_score': result.match_score,
                    'created_at': result.created_at.isoformat()
                })
            
            return {
                'success': True,
                'batch': {
                    'id': latest_batch.id,
                    'name': latest_batch.name,
                    'created_at': latest_batch.created_at.isoformat(),
                    'total_matches': latest_batch.total_matches
                },
                'results': results_data
            }
            
        except Exception as e:
            return {
                'success': False,
                'message': f'Error getting latest results: {str(e)}'
            }
    
    def get_output_files(self):
        """Get list of generated output files."""
        try:
            output_files = []
            
            if os.path.exists(self.output_dir):
                for file in os.listdir(self.output_dir):
                    if file.endswith(('.xlsx', '.csv', '.txt')):
                        file_path = os.path.join(self.output_dir, file)
                        file_stat = os.stat(file_path)
                        
                        output_files.append({
                            'name': file,
                            'size': file_stat.st_size,
                            'modified': datetime.fromtimestamp(file_stat.st_mtime).isoformat(),
                            'path': file_path
                        })
            
            # Sort by modification time (newest first)
            output_files.sort(key=lambda x: x['modified'], reverse=True)
            
            return {
                'success': True,
                'files': output_files
            }
            
        except Exception as e:
            return {
                'success': False,
                'message': f'Error getting output files: {str(e)}'
            }