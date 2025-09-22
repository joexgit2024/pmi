import pandas as pd

# Load charity project data
df = pd.read_excel('input/Charities Project Information 2025 (Responses).xlsx')

print('=== CHARITY PROJECT ANALYSIS ===')
print(f'Total charities: {len(df)}\n')

problem_charities = []

for i, row in df.iterrows():
    org_name = row['Name of the organisation']
    description = str(row['Simple description of the initiative or the project.'])
    outcomes = str(row['What are the key outcomes expected from this initiative or project?'])
    benefits = str(row['How will this initiative benefit your organisation?'])
    expectations = str(row['What outcome(s) do you expect to achieve by participating in this PMDoS event?'])
    
    print(f'{i+1}. {org_name}')
    print(f'   Description: {description}')
    print(f'   Outcomes: {outcomes[:100]}...' if len(outcomes) > 100 else f'   Outcomes: {outcomes}')
    print(f'   Benefits: {benefits[:100]}...' if len(benefits) > 100 else f'   Benefits: {benefits}')
    print(f'   Expectations: {expectations[:100]}...' if len(expectations) > 100 else f'   Expectations: {expectations}')
    
    # Check if description is missing or inadequate
    if (description in ['nan', 'None', ''] or 
        len(description.strip()) < 20 or 
        description.lower().strip() in ['tbd', 'to be determined', 'n/a', 'na']):
        problem_charities.append(org_name)
        print('   ⚠️  WARNING: Inadequate problem statement!')
    
    print('-' * 80)

print(f'\n=== SUMMARY ===')
print(f'Charities with inadequate problem statements: {len(problem_charities)}')
if problem_charities:
    for charity in problem_charities:
        print(f'  - {charity}')