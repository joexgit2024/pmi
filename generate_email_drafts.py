import pandas as pd
import os
from datetime import datetime

def create_email_drafts():
    # Read the Excel file
    df = pd.read_excel('input/2025 - PMI Sydney Chapter Project Management Day of Service (PMDoS) 2025 Professional Registration (Responses).xlsx')
    
    # Read the email template
    with open('revised_acknowledgment_email.txt', 'r', encoding='utf-8') as file:
        email_template = file.read()
    
    # Create email_drafts directory if it doesn't exist
    os.makedirs('email_drafts', exist_ok=True)
    
    # Create individual email files
    for index, row in df.iterrows():
        first_name = str(row['First Name']).strip()
        last_name = str(row['Last Name']).strip()
        email_address = str(row['Preferred Email Address']).strip()
        
        # Replace placeholder in template
        personalized_email = email_template.replace('[PMP Professional Name]', f'{first_name} {last_name}')
        
        # Update the contact email line to include recipient's email
        personalized_email = personalized_email.replace(
            'Email: pmdos_professionals@pmisydney.org',
            f'To: {email_address}\nFrom: pmdos_professionals@pmisydney.org'
        )
        
        # Create filename (safe for filesystem)
        safe_name = f"{first_name}_{last_name}".replace(' ', '_').replace('/', '_')
        filename = f"email_drafts/{index+1:02d}_{safe_name}_email_draft.txt"
        
        # Write individual email file
        with open(filename, 'w', encoding='utf-8') as email_file:
            email_file.write(personalized_email)
        
        print(f"Created: {filename}")
    
    print(f"\n✅ Successfully created {len(df)} email drafts in 'email_drafts/' folder")
    print(f"📧 Ready for copy-paste into pmdos_professionals@pmisydney.org")

if __name__ == "__main__":
    create_email_drafts()