"""
PMI Email Management Dashboard
============================

Simple command-line interface for managing PMI email tracking and generation.

Usage:
    python email_manager.py [command]

Commands:
    status      - Show current tracking status
    generate    - Generate emails for new registrations  
    report      - Generate detailed tracking report
    help        - Show this help message

Examples:
    python email_manager.py status
    python email_manager.py generate
    python email_manager.py report
"""

import sys
import os
from datetime import datetime
from email_tracking_system import EmailTracker
from dynamic_file_loader import get_latest_input_files
import pandas as pd


def show_status():
    """Show current email tracking status"""
    print("📊 PMI EMAIL TRACKING STATUS")
    print("=" * 50)
    
    tracker = EmailTracker()
    
    if tracker.tracking_data["metadata"]["total_emails_sent"] == 0:
        print("❓ No tracking data found")
        print("💡 Run 'python email_manager.py generate' to initialize")
        return
    
    # Basic stats
    metadata = tracker.tracking_data["metadata"]
    print(f"📧 Total emails sent: {metadata['total_emails_sent']}")
    print(f"📦 Total batches: {len(metadata['batches'])}")
    print(f"🗓️ Last updated: {metadata.get('last_updated', 'Unknown')}")
    
    # Recent batches
    print("\\n📅 Recent Batches:")
    for batch in metadata['batches'][-3:]:
        print(f"   • {batch['batch_id']}: {batch['count']} emails on {batch['date']}")
    
    # Check for new registrations
    reg_file, _ = get_latest_input_files()
    if reg_file:
        df = pd.read_excel(reg_file)
        new_registrations = tracker.identify_new_registrations(df)
        
        print(f"\\n🔍 Current Registration File: {os.path.basename(reg_file)}")
        print(f"📝 Total registrations: {len(df)}")
        print(f"🆕 New registrations needing emails: {len(new_registrations)}")
        
        if len(new_registrations) > 0:
            print("\\n🎯 New registrations:")
            for i, (_, row) in enumerate(new_registrations.iterrows(), 1):
                name = f"{row.get('First Name', '')} {row.get('Last Name', '')}".strip()
                email = row.get('Email address', '') or row.get('Preferred Email Address', '')
                print(f"   {i}. {name} - {email}")


def generate_emails():
    """Generate emails for new registrations"""
    print("🚀 GENERATING EMAILS FOR NEW REGISTRATIONS")
    print("=" * 50)
    
    from enhanced_email_generator import generate_incremental_emails
    success = generate_incremental_emails()
    
    if success:
        print("\\n✅ Email generation completed successfully!")
    else:
        print("\\n❌ Email generation failed!")
    
    return success


def generate_report():
    """Generate detailed tracking report"""
    print("📋 DETAILED EMAIL TRACKING REPORT")
    print("=" * 50)
    
    tracker = EmailTracker()
    
    if tracker.tracking_data["metadata"]["total_emails_sent"] == 0:
        print("❓ No tracking data found")
        return
    
    # Full report
    print(tracker.get_summary_report())
    
    # Batch details
    print("\\n📦 DETAILED BATCH INFORMATION:")
    print("-" * 30)
    
    for batch in tracker.tracking_data["metadata"]["batches"]:
        print(f"\\nBatch: {batch['batch_id']}")
        print(f"Date: {batch['date']}")
        print(f"Count: {batch['count']} emails")
        print(f"Folder: {batch['folder']}")
        
        # List emails in this batch
        batch_emails = [email for email, data in tracker.tracking_data["sent_emails"].items() 
                       if data.get("batch_id") == batch["batch_id"]]
        
        if len(batch_emails) <= 5:
            print("Emails:")
            for email in batch_emails:
                name = tracker.tracking_data["sent_emails"][email]["name"]
                print(f"  • {name} - {email}")
        else:
            print(f"Emails: {len(batch_emails)} total (showing first 3)")
            for email in batch_emails[:3]:
                name = tracker.tracking_data["sent_emails"][email]["name"]
                print(f"  • {name} - {email}")
            print(f"  ... and {len(batch_emails) - 3} more")
    
    # Save report to file
    report_file = f"email_tracking_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
    try:
        with open(report_file, 'w', encoding='utf-8') as f:
            f.write("PMI EMAIL TRACKING DETAILED REPORT\\n")
            f.write("=" * 50 + "\\n")
            f.write(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\\n\\n")
            f.write(tracker.get_summary_report())
            
            f.write("\\n\\nDETAILED EMAIL LIST:\\n")
            f.write("-" * 30 + "\\n")
            for email, data in tracker.tracking_data["sent_emails"].items():
                f.write(f"{data['name']} - {email} - {data['sent_date']} - {data['batch_id']}\\n")
        
        print(f"\\n💾 Detailed report saved to: {report_file}")
        
    except Exception as e:
        print(f"\\n⚠️ Could not save report file: {e}")


def show_help():
    """Show help information"""
    print(__doc__)


def main():
    """Main function"""
    
    if len(sys.argv) < 2:
        command = "status"  # Default command
    else:
        command = sys.argv[1].lower()
    
    print(f"🎯 PMI Email Manager - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 60)
    
    if command in ["status", "s"]:
        show_status()
    elif command in ["generate", "gen", "g"]:
        generate_emails()
    elif command in ["report", "r"]:
        generate_report()
    elif command in ["help", "h", "?"]:
        show_help()
    else:
        print(f"❓ Unknown command: {command}")
        print("💡 Available commands: status, generate, report, help")
        print("   Example: python email_manager.py status")
        return False
    
    return True


if __name__ == "__main__":
    main()