"""
PMDoS 2025 Selection Notification Email Generator
===============================================

This script generates personalized email drafts for all three selection outcome groups:
1. Selected and Matched - Full participants with project assignments
2. Selected as Backup - Reserve participants with induction invitation  
3. Not Selected - Respectful notification with future opportunities

Usage:
    python generate_selection_notification_emails.py

Input Files Required:
- input/selected and matched.xlsx
- input/seected as backup.xlsx  
- input/not selected.xlsx

Output:
- selection_notifications/selected_and_matched/ - Individual email drafts
- selection_notifications/selected_as_backup/ - Individual email drafts
- selection_notifications/not_selected/ - Individual email drafts
- Selection summary and tracking files
"""

import pandas as pd
import os
from datetime import datetime


def create_selection_email_drafts():
    """Generate personalized email drafts for all three selection groups"""
    
    # Define file mappings
    selection_groups = {
        'selected_and_matched': {
            'input_file': 'input/selected and matched.xlsx',
            'template_file': 'selection_notifications/selected_and_matched_template.txt',
            'output_dir': 'selection_notifications/selected_and_matched',
            'group_name': 'Selected and Matched'
        },
        'selected_as_backup': {
            'input_file': 'input/seected as backup.xlsx',
            'template_file': 'selection_notifications/selected_as_backup_template.txt', 
            'output_dir': 'selection_notifications/selected_as_backup',
            'group_name': 'Selected as Backup'
        },
        'not_selected': {
            'input_file': 'input/not selected.xlsx',
            'template_file': 'selection_notifications/not_selected_template.txt',
            'output_dir': 'selection_notifications/not_selected',
            'group_name': 'Not Selected'
        }
    }
    
    print("PMDoS 2025 Selection Notification Email Generator")
    print("=" * 55)
    print(f"Processing selection results and generating email drafts...")
    print()
    
    total_emails = 0
    summary_data = []
    
    # Process each selection group
    for group_key, group_info in selection_groups.items():
        
        print(f"Processing: {group_info['group_name']}")
        print("-" * 30)
        
        # Check if input file exists
        if not os.path.exists(group_info['input_file']):
            print(f"⚠️  Input file not found: {group_info['input_file']}")
            print("   Skipping this group...")
            print()
            continue
            
        # Read the Excel file
        try:
            df = pd.read_excel(group_info['input_file'])
            print(f"📊 Found {len(df)} people in {group_info['group_name']} group")
        except Exception as e:
            print(f"❌ Error reading {group_info['input_file']}: {str(e)}")
            continue
            
        # Read the email template
        try:
            with open(group_info['template_file'], 'r', encoding='utf-8') as file:
                email_template = file.read()
        except Exception as e:
            print(f"❌ Error reading template {group_info['template_file']}: {str(e)}")
            continue
            
        # Create output directory
        os.makedirs(group_info['output_dir'], exist_ok=True)
        
        # Generate individual email files
        group_count = 0
        for index, row in df.iterrows():
            try:
                first_name = str(row['First Name']).strip()
                last_name = str(row['Last Name']).strip()
                email_address = str(row['Preferred Email Address']).strip()
                
                # Replace placeholders in template
                personalized_email = email_template.replace('[PMP Professional Name]', f'{first_name} {last_name}')
                personalized_email = personalized_email.replace('[Email Address]', email_address)
                
                # Create safe filename
                safe_name = f"{first_name}_{last_name}".replace(' ', '_').replace('/', '_')
                filename = f"{group_info['output_dir']}/{index+1:02d}_{safe_name}_notification.txt"
                
                # Write individual email file
                with open(filename, 'w', encoding='utf-8') as email_file:
                    email_file.write(personalized_email)
                
                group_count += 1
                print(f"  ✅ Created: {os.path.basename(filename)}")
                
            except Exception as e:
                print(f"  ❌ Error processing row {index}: {str(e)}")
                continue
        
        total_emails += group_count
        summary_data.append({
            'Group': group_info['group_name'],
            'Count': group_count,
            'Input_File': group_info['input_file'],
            'Output_Directory': group_info['output_dir']
        })
        
        print(f"✅ {group_info['group_name']}: {group_count} email drafts created")
        print()
    
    # Create summary file
    create_selection_summary(summary_data, total_emails)
    
    print("=" * 55)
    print(f"🎉 EMAIL GENERATION COMPLETE!")
    print(f"📧 Total email drafts created: {total_emails}")
    print("=" * 55)
    print()
    print("📁 Output Structure:")
    print("   selection_notifications/")
    print("   ├── selected_and_matched/     - Email drafts for successful participants")
    print("   ├── selected_as_backup/       - Email drafts for backup participants") 
    print("   ├── not_selected/             - Email drafts for non-selected applicants")
    print("   └── SELECTION_SUMMARY.md      - Complete summary and tracking")
    print()
    print("🎯 Next Steps:")
    print("   1. Review email drafts in each folder")
    print("   2. Copy-paste into pmdos_professionals@pmisydney.org Outlook")
    print("   3. Use SELECTION_SUMMARY.md for tracking sent emails")
    print("   4. Send induction Google Meet links to selected groups")


def create_selection_summary(summary_data, total_emails):
    """Create a comprehensive summary and tracking file"""
    
    summary_content = f"""# PMDoS 2025 Selection Notification Summary

## 📊 Email Generation Summary

**Generated:** {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}  
**Total Email Drafts:** {total_emails}

---

## 📁 Group Breakdown

"""
    
    for group in summary_data:
        summary_content += f"""### {group['Group']}
- **Count:** {group['Count']} people
- **Input File:** `{group['Input_File']}`
- **Output Directory:** `{group['Output_Directory']}/`
- **Status:** Ready for sending

"""
    
    summary_content += f"""---

## 📧 Email Templates Used

### 1. Selected and Matched Template
- ✅ Congratulations message
- 📅 Mandatory induction: October 15, 2025 (Google Meet)
- 🎯 Project assignment details (to be filled)
- 📋 Next steps and commitments

### 2. Selected as Backup Template  
- ⭐ Backup status explanation
- 📅 Induction invitation: October 15, 2025 (Google Meet)
- 🔄 Activation process (Oct 16-25, 2025)
- 🎯 Priority for future events

### 3. Not Selected Template
- 💝 Respectful notification
- 📊 Context about high competition
- 🔄 Future opportunities with PMI Sydney
- 🤝 Alternative ways to contribute

---

## 🗓️ Key Dates

- **October 1, 2025:** Confirmation deadline for selected participants
- **October 15, 2025:** Mandatory induction session (7:00-8:30 PM AEDT)
- **October 16-25, 2025:** Backup activation period
- **November 6, 2025:** PMDoS 2025 main event

---

## 📋 Tracking Checklist

### Selected and Matched ({[g['Count'] for g in summary_data if g['Group'] == 'Selected and Matched'][0] if any(g['Group'] == 'Selected and Matched' for g in summary_data) else 0} people)
"""
    
    # Add tracking checklist for each group
    for group in summary_data:
        if group['Count'] > 0:
            summary_content += f"\n#### {group['Group']} Tracking\n"
            for i in range(1, group['Count'] + 1):
                summary_content += f"- [ ] {i:02d}_[Name]_notification - Email Sent\n"
    
    summary_content += f"""

---

## ⚡ Usage Instructions

### For Each Email Draft:
1. **Open** the individual `.txt` file from appropriate folder
2. **Copy** entire content (Ctrl+A, Ctrl+C)
3. **Log into** pmdos_professionals@pmisydney.org Outlook
4. **Create new email** and paste content
5. **Add subject line:**
   - Selected: "PMDoS 2025: Congratulations - You're In! Next Steps Inside"
   - Backup: "PMDoS 2025: Selected as Backup - Important Next Steps"
   - Not Selected: "PMDoS 2025: Selection Update and Future Opportunities"
6. **Send** the email
7. **Check** the box in this tracking list

### Google Meet Setup:
- **Create** Google Meet for October 15, 2025, 7:00-8:30 PM AEDT
- **Send** meeting link to Selected and Backup groups 24 hours before
- **Prepare** induction materials and project briefs

---

## 📞 Support

- **Questions:** pmdos_professionals@pmisydney.org
- **Technical Issues:** Check individual draft files for formatting
- **Missing Information:** Review original Excel files for data completeness

---

*This summary file helps track the email notification process for PMDoS 2025 selection results.*
"""
    
    # Write summary file
    with open('selection_notifications/SELECTION_SUMMARY.md', 'w', encoding='utf-8') as f:
        f.write(summary_content)
    
    print("📋 Summary file created: selection_notifications/SELECTION_SUMMARY.md")


if __name__ == "__main__":
    create_selection_email_drafts()