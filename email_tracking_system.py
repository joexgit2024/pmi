"""
Email Tracking System for PMI PMDoS Registration Acknowledgments
===============================================================

This module manages tracking of sent acknowledgment emails to avoid duplicates
and enables incremental email processing for new registrations.

Features:
- Tracks all sent emails with registration details
- Identifies new registrations that need acknowledgment
- Generates emails only for new responses
- Organizes new email drafts by date
- Maintains persistent record across multiple runs
"""

import json
import pandas as pd
import os
from datetime import datetime
from pathlib import Path


class EmailTracker:
    def __init__(self, tracking_file="email_tracking.json"):
        self.tracking_file = tracking_file
        self.tracking_data = self.load_tracking_data()
    
    def load_tracking_data(self):
        """Load existing tracking data or create new structure"""
        if os.path.exists(self.tracking_file):
            with open(self.tracking_file, 'r', encoding='utf-8') as f:
                return json.load(f)
        else:
            return {
                "metadata": {
                    "created": datetime.now().isoformat(),
                    "last_updated": None,
                    "total_emails_sent": 0,
                    "batches": []
                },
                "sent_emails": {}  # Key: email_address, Value: tracking info
            }
    
    def save_tracking_data(self):
        """Save tracking data to file"""
        self.tracking_data["metadata"]["last_updated"] = datetime.now().isoformat()
        with open(self.tracking_file, 'w', encoding='utf-8') as f:
            json.dump(self.tracking_data, f, indent=2, ensure_ascii=False)
    
    def initialize_from_existing_drafts(self, email_drafts_folder="email_drafts"):
        """Initialize tracking from existing email drafts (for first-time setup)"""
        if self.tracking_data["metadata"]["total_emails_sent"] > 0:
            print("Tracking data already exists. Skipping initialization.")
            return
        
        print("Initializing tracking data from existing email drafts...")
        
        # Get all existing email draft files
        draft_files = list(Path(email_drafts_folder).glob("*_email_draft.txt"))
        
        # Load registration data to get email addresses
        from dynamic_file_loader import get_latest_input_files
        reg_file, _ = get_latest_input_files()
        df = pd.read_excel(reg_file)
        
        initialized_count = 0
        for draft_file in draft_files:
            # Extract name from filename (e.g., "01_Maria_Mainhardt_email_draft.txt")
            filename = draft_file.name
            name_part = filename.split('_', 1)[1].replace('_email_draft.txt', '').replace('_', ' ')
            
            # Find matching record in registration data
            matching_row = None
            for _, row in df.iterrows():
                full_name = f"{row.get('First Name', '')} {row.get('Last Name', '')}".strip()
                if self._normalize_name(full_name) == self._normalize_name(name_part):
                    matching_row = row
                    break
            
            if matching_row is not None:
                email = matching_row.get('Email address', '') or matching_row.get('Preferred Email Address', '')
                if email:
                    self.tracking_data["sent_emails"][email.lower()] = {
                        "name": f"{matching_row.get('First Name', '')} {matching_row.get('Last Name', '')}".strip(),
                        "email": email,
                        "sent_date": "2025-09-26",  # Approximate date for existing emails
                        "draft_file": str(draft_file),
                        "batch_id": "initial_batch_29",
                        "registration_timestamp": str(matching_row.get('Timestamp', '')),
                        "pmi_id": str(matching_row.get('PMI ID Number', ''))
                    }
                    initialized_count += 1
        
        # Update metadata
        self.tracking_data["metadata"]["total_emails_sent"] = initialized_count
        self.tracking_data["metadata"]["batches"].append({
            "batch_id": "initial_batch_29",
            "date": "2025-09-26",
            "count": initialized_count,
            "folder": email_drafts_folder
        })
        
        self.save_tracking_data()
        print(f"Initialized tracking for {initialized_count} existing email drafts")
    
    def _normalize_name(self, name):
        """Normalize name for comparison"""
        return name.lower().strip().replace('  ', ' ')
    
    def identify_new_registrations(self, registration_df):
        """Identify registrations that haven't received acknowledgment emails yet"""
        new_registrations = []
        
        for _, row in registration_df.iterrows():
            email = (row.get('Email address', '') or row.get('Preferred Email Address', '')).lower()
            
            if email and email not in self.tracking_data["sent_emails"]:
                new_registrations.append(row)
        
        return pd.DataFrame(new_registrations)
    
    def generate_batch_id(self):
        """Generate a new batch ID"""
        today = datetime.now().strftime("%Y%m%d")
        existing_batches_today = [b for b in self.tracking_data["metadata"]["batches"] 
                                 if b["batch_id"].startswith(f"batch_{today}")]
        batch_number = len(existing_batches_today) + 1
        return f"batch_{today}_{batch_number:02d}"
    
    def record_sent_emails(self, sent_emails_df, batch_folder):
        """Record newly sent emails in tracking system"""
        batch_id = self.generate_batch_id()
        today = datetime.now().strftime("%Y-%m-%d")
        
        for _, row in sent_emails_df.iterrows():
            email = (row.get('Email address', '') or row.get('Preferred Email Address', '')).lower()
            name = f"{row.get('First Name', '')} {row.get('Last Name', '')}".strip()
            
            self.tracking_data["sent_emails"][email] = {
                "name": name,
                "email": row.get('Email address', '') or row.get('Preferred Email Address', ''),
                "sent_date": today,
                "draft_file": f"{batch_folder}/{self._generate_filename(name)}",
                "batch_id": batch_id,
                "registration_timestamp": str(row.get('Timestamp', '')),
                "pmi_id": str(row.get('PMI ID Number', ''))
            }
        
        # Update metadata
        self.tracking_data["metadata"]["total_emails_sent"] += len(sent_emails_df)
        self.tracking_data["metadata"]["batches"].append({
            "batch_id": batch_id,
            "date": today,
            "count": len(sent_emails_df),
            "folder": batch_folder
        })
        
        self.save_tracking_data()
        return batch_id
    
    def _generate_filename(self, name):
        """Generate filename for email draft"""
        current_total = self.tracking_data["metadata"]["total_emails_sent"]
        safe_name = name.replace(' ', '_').replace('.', '').replace(',', '')
        return f"{current_total + 1:02d}_{safe_name}_email_draft.txt"
    
    def get_summary_report(self):
        """Generate summary report of email tracking"""
        total_sent = self.tracking_data["metadata"]["total_emails_sent"]
        batches = self.tracking_data["metadata"]["batches"]
        
        report = f"""
EMAIL TRACKING SUMMARY REPORT
============================
Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}

OVERALL STATISTICS:
- Total emails sent: {total_sent}
- Total batches: {len(batches)}
- First batch: {batches[0]['date'] if batches else 'None'}
- Last batch: {batches[-1]['date'] if batches else 'None'}

BATCH DETAILS:
"""
        
        for batch in batches:
            report += f"- {batch['batch_id']}: {batch['count']} emails on {batch['date']} (folder: {batch['folder']})\n"
        
        return report
    
    def get_sent_emails_list(self):
        """Get list of all sent email addresses"""
        return list(self.tracking_data["sent_emails"].keys())


def create_incremental_email_drafts():
    """Main function to create email drafts for new registrations only"""
    print("=== PMI EMAIL TRACKING & INCREMENTAL DRAFT GENERATION ===")
    
    # Initialize email tracker
    tracker = EmailTracker()
    
    # Initialize from existing drafts if first time
    tracker.initialize_from_existing_drafts()
    
    # Load latest registration data
    from dynamic_file_loader import get_latest_input_files
    reg_file, _ = get_latest_input_files()
    print(f"Loading registration data from: {reg_file}")
    
    df = pd.read_excel(reg_file)
    print(f"Total registrations in file: {len(df)}")
    
    # Identify new registrations
    new_registrations = tracker.identify_new_registrations(df)
    print(f"New registrations needing emails: {len(new_registrations)}")
    
    if len(new_registrations) == 0:
        print("✅ No new registrations found. All current registrations have been sent acknowledgment emails.")
        return
    
    # Create date-based folder for new drafts
    today = datetime.now().strftime("%Y-%m-%d")
    new_folder = f"new_email_drafts/{today}"
    os.makedirs(new_folder, exist_ok=True)
    
    # Generate email drafts for new registrations
    print(f"Generating {len(new_registrations)} new email drafts in folder: {new_folder}")
    
    # Load email template
    with open('revised_acknowledgment_email.txt', 'r', encoding='utf-8') as file:
        email_template = file.read()
    
    current_number = tracker.tracking_data["metadata"]["total_emails_sent"] + 1
    
    for _, row in new_registrations.iterrows():
        # Extract details
        first_name = row.get('First Name', '')
        last_name = row.get('Last Name', '')
        full_name = f"{first_name} {last_name}".strip()
        
        # Create personalized email (proper mail merge like original)
        personalized_email = email_template.replace('[PMP Professional Name]', full_name)
        
        # Get email address and add To/From fields like original
        email_address = row.get('Email address', '') or row.get('Preferred Email Address', '')
        if email_address:
            personalized_email = personalized_email.replace(
                'Email: pmdos_professionals@pmisydney.org',
                f'To: {email_address}\nFrom: pmdos_professionals@pmisydney.org'
            )
        
        # Generate filename
        safe_name = full_name.replace(' ', '_').replace('.', '').replace(',', '')
        filename = f"{current_number:02d}_{safe_name}_email_draft.txt"
        
        # Write email draft
        with open(f"{new_folder}/{filename}", 'w', encoding='utf-8') as file:
            file.write(personalized_email)
        
        print(f"Created: {filename}")
        current_number += 1
    
    # Record the sent emails in tracking system
    batch_id = tracker.record_sent_emails(new_registrations, new_folder)
    
    # Create summary file
    with open(f"{new_folder}/NEW_EMAILS_SUMMARY.md", 'w', encoding='utf-8') as f:
        f.write(f"# New Email Drafts Summary\n\n")
        f.write(f"**Batch ID:** {batch_id}\n")
        f.write(f"**Date:** {today}\n")
        f.write(f"**Count:** {len(new_registrations)}\n\n")
        f.write(f"## New Registrations:\n")
        for i, (_, row) in enumerate(new_registrations.iterrows(), 1):
            name = f"{row.get('First Name', '')} {row.get('Last Name', '')}".strip()
            email = row.get('Email address', '') or row.get('Preferred Email Address', '')
            f.write(f"{i}. **{name}** - {email}\n")
    
    print(f"\n✅ SUCCESS!")
    print(f"📧 Generated {len(new_registrations)} new email drafts")
    print(f"📁 Saved in folder: {new_folder}")
    print(f"🏷️ Batch ID: {batch_id}")
    print(f"\n📊 {tracker.get_summary_report()}")


if __name__ == "__main__":
    create_incremental_email_drafts()