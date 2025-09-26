"""
Email Service for PMI Web Application
====================================

This service integrates the email_manager.py functionality
into the web application, providing email generation and management capabilities.
"""

import os
import sys
import subprocess
from datetime import datetime
import pandas as pd
from pathlib import Path

# Add the project root to Python path to import our email modules
project_root = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, project_root)

from app import db
from app.models import EmailBatch, Registration

# Import email tracking system from the root directory
try:
    from email_tracking_system import EmailTracker
    from enhanced_email_generator import generate_incremental_emails
except ImportError:
    # Alternative import path in case the above doesn't work
    import importlib.util
    
    # EmailTracker
    spec = importlib.util.spec_from_file_location("email_tracking_system", 
                                                os.path.join(project_root, "email_tracking_system.py"))
    email_tracking_system = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(email_tracking_system)
    EmailTracker = email_tracking_system.EmailTracker
    
    # Enhanced email generator
    spec = importlib.util.spec_from_file_location("enhanced_email_generator", 
                                                os.path.join(project_root, "enhanced_email_generator.py"))
    enhanced_email_generator = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(enhanced_email_generator)
    generate_incremental_emails = enhanced_email_generator.generate_incremental_emails


class EmailService:
    """Service for managing email generation and tracking."""
    
    def __init__(self):
        self.project_root = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
        self.tracker = EmailTracker()
    
    def generate_email_drafts(self):
        """
        Generate email drafts for new registrations only.
        
        Returns:
            dict: Result of the email generation process
        """
        try:
            # Change to project directory
            original_cwd = os.getcwd()
            os.chdir(self.project_root)
            
            # Use the enhanced email generator
            success = generate_incremental_emails()
            
            # Restore original directory
            os.chdir(original_cwd)
            
            if success:
                # Get the latest statistics
                stats = self._get_email_statistics()
                
                return {
                    'success': True,
                    'message': 'Email drafts generated successfully',
                    'new_emails_count': stats.get('new_emails_count', 0),
                    'total_emails_sent': stats.get('total_emails_sent', 0),
                    'latest_batch_id': stats.get('latest_batch_id', None)
                }
            else:
                return {
                    'success': False,
                    'message': 'No new registrations found or email generation failed'
                }
                
        except Exception as e:
            # Restore original directory on error
            try:
                os.chdir(original_cwd)
            except:
                pass
            
            return {
                'success': False,
                'message': f'Error generating email drafts: {str(e)}'
            }
    
    def get_email_status(self):
        """Get current email tracking status."""
        try:
            # Reload tracking data
            self.tracker = EmailTracker()
            
            if self.tracker.tracking_data["metadata"]["total_emails_sent"] == 0:
                return {
                    'success': True,
                    'initialized': False,
                    'message': 'No tracking data found. Run email generation to initialize.',
                    'total_emails_sent': 0,
                    'total_batches': 0,
                    'new_registrations': 0
                }
            
            # Get new registrations count
            try:
                from dynamic_file_loader import get_latest_input_files
                reg_file, _ = get_latest_input_files()
                
                if reg_file:
                    df = pd.read_excel(reg_file)
                    new_registrations = self.tracker.identify_new_registrations(df)
                    new_count = len(new_registrations)
                    total_registrations = len(df)
                else:
                    new_count = 0
                    total_registrations = 0
                    
            except Exception as e:
                new_count = 0
                total_registrations = 0
            
            metadata = self.tracker.tracking_data["metadata"]
            
            # Get recent batches
            recent_batches = []
            for batch in metadata['batches'][-5:]:  # Last 5 batches
                recent_batches.append({
                    'batch_id': batch['batch_id'],
                    'date': batch['date'],
                    'count': batch['count'],
                    'folder': batch['folder']
                })
            
            return {
                'success': True,
                'initialized': True,
                'total_emails_sent': metadata['total_emails_sent'],
                'total_batches': len(metadata['batches']),
                'last_updated': metadata.get('last_updated', 'Unknown'),
                'new_registrations': new_count,
                'total_registrations': total_registrations,
                'recent_batches': recent_batches
            }
            
        except Exception as e:
            return {
                'success': False,
                'message': f'Error getting email status: {str(e)}'
            }
    
    def _get_email_statistics(self):
        """Get email statistics after generation."""
        try:
            # Reload tracker to get latest data
            self.tracker = EmailTracker()
            
            metadata = self.tracker.tracking_data["metadata"]
            
            # Get the latest batch to determine new emails count
            latest_batch = None
            if metadata['batches']:
                latest_batch = metadata['batches'][-1]
            
            return {
                'total_emails_sent': metadata['total_emails_sent'],
                'new_emails_count': latest_batch['count'] if latest_batch else 0,
                'latest_batch_id': latest_batch['batch_id'] if latest_batch else None,
                'total_batches': len(metadata['batches'])
            }
            
        except Exception as e:
            return {
                'total_emails_sent': 0,
                'new_emails_count': 0,
                'latest_batch_id': None,
                'total_batches': 0
            }
    
    def send_emails(self, email_ids):
        """
        Send specific emails (placeholder - would integrate with actual email service).
        
        Args:
            email_ids (list): List of email IDs to send
            
        Returns:
            dict: Result of the send operation
        """
        try:
            # This would integrate with an actual email service like SendGrid, SES, etc.
            # For now, we'll just mark them as sent in our tracking
            
            return {
                'success': True,
                'message': f'Would send {len(email_ids)} emails (placeholder implementation)',
                'sent_count': len(email_ids)
            }
            
        except Exception as e:
            return {
                'success': False,
                'message': f'Error sending emails: {str(e)}'
            }
    
    def bulk_send_emails(self, email_ids):
        """
        Send multiple emails in bulk.
        
        Args:
            email_ids (list): List of email IDs to send
            
        Returns:
            dict: Result of the bulk send operation
        """
        try:
            # This would implement bulk sending logic
            # For now, it's the same as individual sending
            
            return self.send_emails(email_ids)
            
        except Exception as e:
            return {
                'success': False,
                'message': f'Error bulk sending emails: {str(e)}'
            }
    
    def get_email_drafts(self):
        """Get list of generated email draft files."""
        try:
            email_drafts = []
            
            # Check original email_drafts folder
            original_folder = os.path.join(self.project_root, "email_drafts")
            if os.path.exists(original_folder):
                for file in os.listdir(original_folder):
                    if file.endswith("_email_draft.txt"):
                        file_path = os.path.join(original_folder, file)
                        file_stat = os.stat(file_path)
                        
                        email_drafts.append({
                            'name': file,
                            'folder': 'email_drafts',
                            'size': file_stat.st_size,
                            'modified': datetime.fromtimestamp(file_stat.st_mtime).isoformat(),
                            'path': file_path
                        })
            
            # Check new email drafts folders
            new_drafts_folder = os.path.join(self.project_root, "new_email_drafts")
            if os.path.exists(new_drafts_folder):
                for date_folder in os.listdir(new_drafts_folder):
                    date_path = os.path.join(new_drafts_folder, date_folder)
                    if os.path.isdir(date_path):
                        for file in os.listdir(date_path):
                            if file.endswith("_email_draft.txt"):
                                file_path = os.path.join(date_path, file)
                                file_stat = os.stat(file_path)
                                
                                email_drafts.append({
                                    'name': file,
                                    'folder': f'new_email_drafts/{date_folder}',
                                    'size': file_stat.st_size,
                                    'modified': datetime.fromtimestamp(file_stat.st_mtime).isoformat(),
                                    'path': file_path
                                })
            
            # Sort by modification time (newest first)
            email_drafts.sort(key=lambda x: x['modified'], reverse=True)
            
            return {
                'success': True,
                'drafts': email_drafts
            }
            
        except Exception as e:
            return {
                'success': False,
                'message': f'Error getting email drafts: {str(e)}'
            }
    
    def get_tracking_report(self):
        """Generate detailed tracking report."""
        try:
            # Reload tracker
            self.tracker = EmailTracker()
            
            if self.tracker.tracking_data["metadata"]["total_emails_sent"] == 0:
                return {
                    'success': False,
                    'message': 'No tracking data available'
                }
            
            # Generate report using existing method
            report_content = self.tracker.get_summary_report()
            
            # Get detailed batch information
            batches = []
            for batch in self.tracker.tracking_data["metadata"]["batches"]:
                batch_emails = [email for email, data in self.tracker.tracking_data["sent_emails"].items() 
                               if data.get("batch_id") == batch["batch_id"]]
                
                batches.append({
                    'batch_id': batch['batch_id'],
                    'date': batch['date'],
                    'count': batch['count'],
                    'folder': batch['folder'],
                    'emails': batch_emails[:5]  # First 5 emails for preview
                })
            
            return {
                'success': True,
                'report_content': report_content,
                'batches': batches,
                'total_emails': self.tracker.tracking_data["metadata"]["total_emails_sent"],
                'total_batches': len(self.tracker.tracking_data["metadata"]["batches"])
            }
            
        except Exception as e:
            return {
                'success': False,
                'message': f'Error generating tracking report: {str(e)}'
            }