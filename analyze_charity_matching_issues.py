from enhanced_pmp_charity_matching import analyze_charity_requirements, load_and_process_data

# Load data and analyze charity requirements
_, charity_df = load_and_process_data()
projects = analyze_charity_requirements(charity_df)

print('=== CHARITY SKILL REQUIREMENT ANALYSIS ===')
print('How the algorithm assigns skill weights based on problem descriptions:\n')

problematic_charities = []

for project in projects:
    org_name = project['Organization']
    total_skill_weight = sum(project['Required_Skills'].values())
    
    print(f'{org_name}:')
    print(f'  Total skill weight: {total_skill_weight}')
    print(f'  Priority: {project["Priority_Level"]}')
    print(f'  Complexity: {project["Complexity"]}')
    
    # Show top skills identified
    top_skills = sorted(project['Required_Skills'].items(), 
                       key=lambda x: x[1], reverse=True)[:3]
    if any(weight > 0 for _, weight in top_skills):
        print('  Top identified skills:')
        for skill, weight in top_skills:
            if weight > 0:
                print(f'    - {skill}: {weight}')
    else:
        print('  ⚠️  No significant skills identified!')
        problematic_charities.append(org_name)
    
    print()

print('=== SUMMARY ===')
print(f'Charities with low/no skill identification: {len(problematic_charities)}')
for charity in problematic_charities:
    print(f'  - {charity}')

print('\n=== RECOMMENDATION ===')
if problematic_charities:
    print('For charities with inadequate problem statements, the algorithm will:')
    print('1. Assign very low skill weights (or zero)')
    print('2. Rely mainly on general factors like:')
    print('   - PMP experience level')
    print('   - Interest in non-profit work')
    print('   - LinkedIn profile quality')
    print('   - Profile completeness')
    print('3. May result in suboptimal matches')
    print('\nSUGGESTED SOLUTIONS:')
    print('A. Contact these charities for better problem statements')
    print('B. Use default skill weights for common charity needs')
    print('C. Manual assignment based on organization type')