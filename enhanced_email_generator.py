"""
Enhanced Email Draft Generator with Tracking Integration
=======================================================

This enhanced version integrates with the email tracking system to:
- Only generate drafts for new registrations
- Maintain tracking of sent emails
- Organize new drafts in date-based folders
- Provide comprehensive reporting

Usage:
    python enhanced_email_generator.py

Features:
- Automatic detection of new registrations
- Integration with existing tracking system
- Date-based folder organization
- Comprehensive tracking and reporting
"""

import pandas as pd
import os
from datetime import datetime
from email_tracking_system import EmailTracker

def generate_incremental_emails():
    """Generate email drafts only for new registrations"""
    
    print("=== ENHANCED EMAIL DRAFT GENERATOR WITH TRACKING ===")
    print(f"Started at: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 60)
    
    # Initialize email tracker
    tracker = EmailTracker()
    
    # Initialize from existing drafts if first time
    if tracker.tracking_data["metadata"]["total_emails_sent"] == 0:
        print("🔄 First time setup - initializing from existing email drafts...")
        tracker.initialize_from_existing_drafts()
        print("✅ Initialization complete")
    
    # Load latest registration data using dynamic detection
    from dynamic_file_loader import get_latest_input_files
    
    reg_file, _ = get_latest_input_files()
    if not reg_file:
        print("❌ ERROR: Could not find PMP registration file")
        return False
    
    print(f"📁 Loading registration data from: {os.path.basename(reg_file)}")
    
    try:
        df = pd.read_excel(reg_file)
        print(f"📊 Total registrations in file: {len(df)}")
    except Exception as e:
        print(f"❌ ERROR loading Excel file: {e}")
        return False
    
    # Identify new registrations
    print("🔍 Identifying new registrations...")
    new_registrations = tracker.identify_new_registrations(df)
    
    if len(new_registrations) == 0:
        print("✅ No new registrations found!")
        print("📧 All current registrations have already been sent acknowledgment emails.")
        print("\n📊 Current Status:")
        print(tracker.get_summary_report())
        return True
    
    print(f"🎯 Found {len(new_registrations)} new registrations needing emails:")
    for i, (_, row) in enumerate(new_registrations.iterrows(), 1):
        name = f"{row.get('First Name', '')} {row.get('Last Name', '')}".strip()
        email = row.get('Email address', '') or row.get('Preferred Email Address', '')
        print(f"   {i}. {name} - {email}")
    
    # Create date-based folder for new drafts
    today = datetime.now().strftime("%Y-%m-%d")
    new_folder = f"new_email_drafts/{today}"
    
    try:
        os.makedirs(new_folder, exist_ok=True)
        print(f"📁 Created/Using folder: {new_folder}")
    except Exception as e:
        print(f"❌ ERROR creating folder: {e}")
        return False
    
    # Load email template
    template_file = 'revised_acknowledgment_email.txt'
    if not os.path.exists(template_file):
        print(f"❌ ERROR: Email template file '{template_file}' not found")
        return False
    
    try:
        with open(template_file, 'r', encoding='utf-8') as file:
            email_template = file.read()
        print(f"📧 Loaded email template from: {template_file}")
    except Exception as e:
        print(f"❌ ERROR loading email template: {e}")
        return False
    
    # Generate email drafts
    print(f"✍️  Generating {len(new_registrations)} email drafts...")
    print("=" * 60)
    
    current_number = tracker.tracking_data["metadata"]["total_emails_sent"] + 1
    created_files = []
    
    for i, (_, row) in enumerate(new_registrations.iterrows(), 1):
        # Extract details
        first_name = row.get('First Name', '').strip()
        last_name = row.get('Last Name', '').strip()
        full_name = f"{first_name} {last_name}".strip()
        email_address = row.get('Email address', '') or row.get('Preferred Email Address', '')
        
        if not first_name:
            print(f"⚠️  Warning: No first name for registration {i}, skipping...")
            continue
        
        # Create personalized email (proper mail merge like original)
        personalized_email = email_template.replace('[PMP Professional Name]', full_name)
        
        # Get email address and add To/From fields like original
        email_address = row.get('Email address', '') or row.get('Preferred Email Address', '')
        if email_address:
            personalized_email = personalized_email.replace(
                'Email: pmdos_professionals@pmisydney.org',
                f'To: {email_address}\nFrom: pmdos_professionals@pmisydney.org'
            )
        
        # Generate safe filename
        safe_name = full_name.replace(' ', '_').replace('.', '').replace(',', '').replace('/', '_')
        filename = f"{current_number:02d}_{safe_name}_email_draft.txt"
        
        # Write email draft
        try:
            filepath = os.path.join(new_folder, filename)
            with open(filepath, 'w', encoding='utf-8') as file:
                file.write(personalized_email)
            
            created_files.append(filename)
            print(f"   ✅ {filename} -> {email_address}")
            current_number += 1
            
        except Exception as e:
            print(f"   ❌ ERROR creating {filename}: {e}")
    
    if not created_files:
        print("❌ No email drafts were created")
        return False
    
    # Record the sent emails in tracking system
    print("\\n📝 Updating tracking system...")
    try:
        batch_id = tracker.record_sent_emails(new_registrations, new_folder)
        print(f"✅ Recorded batch: {batch_id}")
    except Exception as e:
        print(f"❌ ERROR updating tracking: {e}")
        return False
    
    # Create summary file
    summary_file = os.path.join(new_folder, "NEW_EMAILS_SUMMARY.md")
    try:
        with open(summary_file, 'w', encoding='utf-8') as f:
            f.write(f"# New Email Drafts Summary\\n\\n")
            f.write(f"**Generated:** {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\\n")
            f.write(f"**Batch ID:** {batch_id}\\n")
            f.write(f"**Date:** {today}\\n")
            f.write(f"**Count:** {len(created_files)}\\n")
            f.write(f"**Folder:** {new_folder}\\n\\n")
            
            f.write(f"## New Registrations:\\n")
            for i, (_, row) in enumerate(new_registrations.iterrows(), 1):
                name = f"{row.get('First Name', '')} {row.get('Last Name', '')}".strip()
                email = row.get('Email address', '') or row.get('Preferred Email Address', '')
                pmi_id = row.get('PMI ID Number', 'N/A')
                f.write(f"{i}. **{name}** - {email} (PMI ID: {pmi_id})\\n")
            
            f.write(f"\\n## Files Generated:\\n")
            for i, filename in enumerate(created_files, 1):
                f.write(f"{i}. `{filename}`\\n")
        
        print(f"📄 Summary saved: {summary_file}")
    except Exception as e:
        print(f"⚠️  Warning: Could not create summary file: {e}")
    
    # Final success report
    print("\\n" + "=" * 60)
    print("🎉 EMAIL GENERATION COMPLETED SUCCESSFULLY!")
    print("=" * 60)
    print(f"📧 Generated: {len(created_files)} email drafts")
    print(f"📁 Location: {new_folder}")
    print(f"🏷️  Batch ID: {batch_id}")
    print(f"📊 Total emails tracked: {tracker.tracking_data['metadata']['total_emails_sent']}")
    
    print("\\n📋 Next Steps:")
    print("1. Review the generated email drafts")
    print("2. Send the emails to the new registrants") 
    print("3. The system will track these as sent for future runs")
    
    print("\\n📊 Overall Tracking Summary:")
    print(tracker.get_summary_report())
    
    return True


def show_tracking_status():
    """Show current tracking status without generating new emails"""
    print("=== CURRENT EMAIL TRACKING STATUS ===")
    
    tracker = EmailTracker()
    
    if tracker.tracking_data["metadata"]["total_emails_sent"] == 0:
        print("📋 No email tracking data found.")
        print("💡 Run the generator first to initialize tracking from existing drafts.")
        return
    
    print(tracker.get_summary_report())
    
    # Show recent activity
    batches = tracker.tracking_data["metadata"]["batches"]
    if batches:
        print("\\n📅 Recent Batches:")
        for batch in batches[-3:]:  # Show last 3 batches
            print(f"   • {batch['batch_id']}: {batch['count']} emails ({batch['date']})")


if __name__ == "__main__":
    import sys
    
    if len(sys.argv) > 1 and sys.argv[1] == "--status":
        show_tracking_status()
    else:
        generate_incremental_emails()