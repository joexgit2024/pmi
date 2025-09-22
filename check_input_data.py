import pandas as pd

# Load the current file
df = pd.read_excel('input/2025 - PMI Sydney Chapter Project Management Day of Service (PMDoS) 2025 Professional Registration (Responses).xlsx')

print('=== CURRENT INPUT FILE ANALYSIS ===')
print(f'Total rows: {len(df)}')
print(f'Non-empty First Name rows: {df["First Name"].notna().sum()}')
print(f'Non-empty Last Name rows: {df["Last Name"].notna().sum()}')

print('\n=== ALL PARTICIPANTS ===')
for i, row in df.iterrows():
    first_name = row['First Name'] if pd.notna(row['First Name']) else 'MISSING'
    last_name = row['Last Name'] if pd.notna(row['Last Name']) else 'MISSING'
    print(f'{i+1:2d}. {first_name} {last_name}')

print('\n=== CHECKING FOR LINKEDIN URLS ===')
linkedin_count = df['LinkedIn Profile URL'].notna().sum()
print(f'Participants with LinkedIn URLs: {linkedin_count}')

print('\n=== CHECKING FOR EMPTY ROWS ===')
# Check if any rows are completely empty
empty_rows = df.isnull().all(axis=1).sum()
print(f'Completely empty rows: {empty_rows}')

# Check if any critical columns are missing
critical_columns = ['First Name', 'Last Name', 'Email address']
for col in critical_columns:
    missing = df[col].isna().sum()
    print(f'Missing {col}: {missing}')