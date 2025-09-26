"""
PMDoS 2025 Project Assignment Email Updater
==========================================

This script updates the "Selected and Matched" email drafts with actual project assignment details
after the matching process is complete.

Usage:
    python update_project_assignments.py

This script will:
1. Read the matching results from Output/PMI_PMP_Charity_Matching_Results_Enhanced.xlsx
2. Update the selected_and_matched email drafts with specific project details
3. Replace placeholder text with actual charity assignments and team information
"""

import pandas as pd
import os
import glob


def update_selected_emails_with_assignments():
    """Update selected and matched emails with actual project assignments"""
    
    print("PMDoS 2025 Project Assignment Email Updater")
    print("=" * 45)
    
    # Check if matching results exist
    matching_files = [
        'Output/PMI_PMP_Charity_Matching_Results_Enhanced.xlsx',
        'Output/PMI_PMP_Charity_Flexible_Matching_Results.xlsx'
    ]
    
    matching_file = None
    for file in matching_files:
        if os.path.exists(file):
            matching_file = file
            break
    
    if not matching_file:
        print("❌ No matching results file found!")
        print("Please run the matching analysis first:")
        print("   python run_complete_analysis.py")
        return
    
    print(f"📊 Reading matching results from: {matching_file}")
    
    try:
        # Read matching results
        if 'Flexible' in matching_file:
            df_matches = pd.read_excel(matching_file, sheet_name='Flexible_Matching')
        else:
            df_matches = pd.read_excel(matching_file, sheet_name='Enhanced_Matching_Summary')
        
        print(f"✅ Found {len(df_matches)} project assignments")
        
        # Group assignments by PMP name for team information
        pmp_assignments = {}
        charity_teams = {}
        
        for _, row in df_matches.iterrows():
            pmp_name = row['PMP_Name']
            charity_org = row['Charity_Organization']
            charity_initiative = row['Charity_Initiative']
            match_score = row['Match_Score']
            
            # Store individual assignment
            pmp_assignments[pmp_name] = {
                'charity_org': charity_org,
                'charity_initiative': charity_initiative,
                'match_score': match_score
            }
            
            # Group by charity for team info
            if charity_org not in charity_teams:
                charity_teams[charity_org] = []
            charity_teams[charity_org].append(pmp_name)
        
        # Update email drafts
        draft_files = glob.glob('selection_notifications/selected_and_matched/*_notification.txt')
        updated_count = 0
        
        for draft_file in draft_files:
            # Extract name from filename
            filename = os.path.basename(draft_file)
            # Remove number prefix and suffix
            name_part = filename.split('_', 1)[1].replace('_notification.txt', '')
            # Convert back to readable name
            readable_name = name_part.replace('_', ' ')
            
            # Find matching assignment
            assigned_pmp = None
            for pmp_name in pmp_assignments.keys():
                if pmp_name.replace(' ', '_').lower() == name_part.lower():
                    assigned_pmp = pmp_name
                    break
            
            if assigned_pmp:
                assignment = pmp_assignments[assigned_pmp]
                charity_org = assignment['charity_org']
                
                # Get team members
                team_members = [name for name in charity_teams[charity_org] if name != assigned_pmp]
                if team_members:
                    team_text = ', '.join(team_members)
                else:
                    team_text = 'Team assignment to be confirmed'
                
                # Read current email draft
                with open(draft_file, 'r', encoding='utf-8') as f:
                    email_content = f.read()
                
                # Update placeholders
                email_content = email_content.replace(
                    '- Charity Organization: [To be filled based on matching results]',
                    f'- Charity Organization: {charity_org}'
                )
                email_content = email_content.replace(
                    '- Project Details: [To be filled based on matching results]',
                    f'- Project Initiative: {assignment["charity_initiative"]}'
                )
                email_content = email_content.replace(
                    '- Team Partners: [To be filled based on matching results]',
                    f'- Team Partners: {team_text}'
                )
                
                # Write updated email
                with open(draft_file, 'w', encoding='utf-8') as f:
                    f.write(email_content)
                
                updated_count += 1
                print(f"✅ Updated: {filename} -> {charity_org}")
            else:
                print(f"⚠️  No assignment found for: {readable_name}")
        
        print(f"\n🎉 Successfully updated {updated_count} email drafts with project assignments!")
        print("\n📧 Next Steps:")
        print("   1. Review updated emails in selection_notifications/selected_and_matched/")
        print("   2. All project assignments are now included in the email drafts")
        print("   3. Ready to send to selected participants")
        
    except Exception as e:
        print(f"❌ Error updating assignments: {str(e)}")


if __name__ == "__main__":
    update_selected_emails_with_assignments()