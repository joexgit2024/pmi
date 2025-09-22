"""
Flexible PMP Assignment - All PMPs to Projects
==============================================

This script modifies the matching algorithm to assign ALL PMPs to projects,
allowing some projects to receive 3+ PMPs based on:
- Project complexity and priority
- Match quality scores
- Balanced distribution

Key changes:
- No strict 2-PMP limit per charity
- Priority-based assignment
- Ensures all 22 PMPs are matched
"""

import pandas as pd
from enhanced_pmp_charity_matching import (
    load_and_process_data,
    extract_pmp_skills,
    analyze_charity_requirements,
    calculate_match_score,
    generate_matching_report,
    create_detailed_analysis
)


def calculate_project_capacity_score(charity_project):
    """
    Calculate how many PMPs a project can effectively utilize based on:
    - Complexity (High = more PMPs)
    - Priority (High = more PMPs)  
    - Project scope
    """
    base_capacity = 2  # Minimum PMPs per project
    
    # Complexity bonus
    if charity_project['Complexity'] == 'High':
        complexity_bonus = 2
    elif charity_project['Complexity'] == 'Medium':
        complexity_bonus = 1
    else:
        complexity_bonus = 0
    
    # Priority bonus
    if charity_project['Priority_Level'] == 'High':
        priority_bonus = 1
    else:
        priority_bonus = 0
    
    # Skill diversity bonus (more skills needed = more PMPs)
    skill_count = sum(1 for weight in charity_project['Required_Skills'].values() if weight > 2)
    skill_bonus = min(skill_count // 3, 1)  # Extra PMP for every 3 significant skills
    
    total_capacity = base_capacity + complexity_bonus + priority_bonus + skill_bonus
    return min(total_capacity, 4)  # Cap at 4 PMPs max per project


def create_flexible_matching(pmp_profiles, charity_projects):
    """
    Create flexible matching that assigns ALL PMPs to projects.
    Some projects may get 3+ PMPs based on complexity and priority.
    """
    
    # Calculate all possible match scores
    all_matches = []
    for pmp in pmp_profiles:
        for charity in charity_projects:
            score = calculate_match_score(pmp, charity)
            all_matches.append({
                'PMP_ID': pmp['ID'],
                'PMP_Name': pmp['Name'],
                'Charity_ID': charity['ID'],
                'Organization': charity['Organization'],
                'Initiative': charity['Initiative'],
                'Score': score,
                'PMP_Profile': pmp,
                'Charity_Project': charity
            })
    
    # Sort by score (highest first)
    all_matches.sort(key=lambda x: x['Score'], reverse=True)
    
    # Calculate project capacities
    project_capacities = {}
    for charity in charity_projects:
        capacity = calculate_project_capacity_score(charity)
        project_capacities[charity['ID']] = {
            'max_capacity': capacity,
            'current_assignments': 0,
            'assigned_pmps': []
        }
    
    # Assign PMPs using flexible algorithm
    assigned_pmps = set()
    final_matches = []
    
    print("=== PROJECT CAPACITY ANALYSIS ===")
    for charity in charity_projects:
        capacity = project_capacities[charity['ID']]['max_capacity']
        print(f"{charity['Organization']}: Max capacity {capacity} PMPs")
        print(f"  - Priority: {charity['Priority_Level']}")
        print(f"  - Complexity: {charity['Complexity']}")
        print(f"  - Skill requirements: {sum(1 for w in charity['Required_Skills'].values() if w > 2)} significant skills")
        print()
    
    # First pass: Assign top matches respecting capacities
    for match in all_matches:
        charity_id = match['Charity_ID']
        pmp_id = match['PMP_ID']
        
        if (pmp_id not in assigned_pmps and 
            project_capacities[charity_id]['current_assignments'] < project_capacities[charity_id]['max_capacity']):
            
            # Assign this PMP
            assigned_pmps.add(pmp_id)
            project_capacities[charity_id]['current_assignments'] += 1
            project_capacities[charity_id]['assigned_pmps'].append(match)
            final_matches.append(match)
    
    # Check if all PMPs are assigned
    unassigned_pmps = [p for p in pmp_profiles if p['ID'] not in assigned_pmps]
    
    if unassigned_pmps:
        print(f"=== SECOND PASS: Assigning {len(unassigned_pmps)} remaining PMPs ===")
        
        # Second pass: Assign remaining PMPs to projects with lowest current assignment ratio
        for pmp in unassigned_pmps:
            # Find best available match for this PMP
            best_match = None
            best_score = 0
            
            for charity in charity_projects:
                # Calculate current assignment ratio
                current = project_capacities[charity['ID']]['current_assignments']
                max_cap = project_capacities[charity['ID']]['max_capacity']
                
                # Allow exceeding capacity if needed, but prefer projects with space
                preference_score = (max_cap - current) * 10  # Prefer projects with space
                
                # Find the match score for this PMP-charity combination
                for match in all_matches:
                    if (match['PMP_ID'] == pmp['ID'] and 
                        match['Charity_ID'] == charity['ID']):
                        
                        adjusted_score = match['Score'] + preference_score
                        
                        if adjusted_score > best_score:
                            best_score = adjusted_score
                            best_match = match
                            break
            
            if best_match:
                charity_id = best_match['Charity_ID']
                project_capacities[charity_id]['current_assignments'] += 1
                project_capacities[charity_id]['assigned_pmps'].append(best_match)
                final_matches.append(best_match)
                assigned_pmps.add(pmp['ID'])
                
                print(f"  Assigned {best_match['PMP_Name']} to {best_match['Organization']} (Score: {best_match['Score']:.2f})")
    
    # Create final assignment structure
    assigned_charities = {}
    for charity_id, capacity_info in project_capacities.items():
        if capacity_info['assigned_pmps']:
            assigned_charities[charity_id] = capacity_info['assigned_pmps']
    
    return final_matches, assigned_charities


def generate_flexible_matching_report(final_matches, assigned_charities):
    """
    Generate a report showing the flexible assignment results.
    """
    
    # Create summary DataFrame
    match_data = []
    
    for charity_id, matches in assigned_charities.items():
        charity_info = matches[0]['Charity_Project']
        
        for i, match in enumerate(matches):
            pmp_info = match['PMP_Profile']
            
            # Get top 3 skills for this PMP
            top_skills = sorted(pmp_info['Skills'].items(), key=lambda x: x[1], reverse=True)[:3]
            top_skills_str = ", ".join([f"{skill}: {rating}" for skill, rating in top_skills])
            
            match_data.append({
                'Charity_Organization': charity_info['Organization'],
                'Charity_Initiative': charity_info['Initiative'],
                'Project_Priority': charity_info['Priority_Level'],
                'Project_Complexity': charity_info['Complexity'],
                'Team_Size': len(matches),
                'PMP_Role': f"PMP {i+1}",
                'PMP_Name': pmp_info['Name'],
                'PMP_Experience': pmp_info['Experience'],
                'PMP_Company': pmp_info.get('Company', ''),
                'LinkedIn_Quality': pmp_info.get('LinkedIn_Quality_Score', 0),
                'Match_Score': round(match['Score'], 2),
                'PMP_Top_Skills': top_skills_str,
                'Overall_PMP_Rating': round(pmp_info['Overall_Score'], 2)
            })
    
    return pd.DataFrame(match_data)


def main():
    """
    Main function to run flexible PMP assignment.
    """
    
    print("Flexible PMP Assignment - All PMPs to Projects")
    print("=" * 55)
    
    # Load and process data
    print("Loading data...")
    pmp_df, charity_df = load_and_process_data()
    pmp_profiles = extract_pmp_skills(pmp_df)
    charity_projects = analyze_charity_requirements(charity_df)
    
    print(f"Total PMPs: {len(pmp_profiles)}")
    print(f"Total Projects: {len(charity_projects)}")
    print()
    
    # Create flexible matching
    print("Creating flexible matching...")
    final_matches, assigned_charities = create_flexible_matching(pmp_profiles, charity_projects)
    
    # Generate reports
    print("\nGenerating reports...")
    matching_summary = generate_flexible_matching_report(final_matches, assigned_charities)
    
    # Save results
    output_file = 'PMI_PMP_Charity_Flexible_Matching_Results.xlsx'
    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        
        # Main matching results
        matching_summary.to_excel(writer, sheet_name='Flexible_Matching', index=False)
        
        # Team composition summary
        team_summary = matching_summary.groupby('Charity_Organization').agg({
            'Team_Size': 'first',
            'Project_Priority': 'first',
            'Project_Complexity': 'first',
            'PMP_Name': lambda x: ' | '.join(x),
            'Match_Score': ['mean', 'min', 'max']
        }).round(2)
        
        team_summary.columns = ['Team_Size', 'Priority', 'Complexity', 'Team_Members', 'Avg_Score', 'Min_Score', 'Max_Score']
        team_summary.to_excel(writer, sheet_name='Team_Summary', index=True)
        
        # Project capacity analysis
        capacity_analysis = []
        for charity_id, matches in assigned_charities.items():
            charity_info = matches[0]['Charity_Project']
            max_capacity = calculate_project_capacity_score(charity_info)
            actual_assignments = len(matches)
            
            capacity_analysis.append({
                'Organization': charity_info['Organization'],
                'Max_Recommended_Capacity': max_capacity,
                'Actual_Assignments': actual_assignments,
                'Capacity_Utilization': f"{(actual_assignments/max_capacity)*100:.0f}%",
                'Priority': charity_info['Priority_Level'],
                'Complexity': charity_info['Complexity'],
                'Over_Capacity': 'Yes' if actual_assignments > max_capacity else 'No'
            })
        
        capacity_df = pd.DataFrame(capacity_analysis)
        capacity_df.to_excel(writer, sheet_name='Capacity_Analysis', index=False)
        
        # Format worksheets
        workbook = writer.book
        header_format = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'top',
            'fg_color': '#D7E4BC',
            'border': 1
        })
        
        for sheet_name in writer.sheets:
            worksheet = writer.sheets[sheet_name]
            worksheet.set_row(0, None, header_format)
            worksheet.autofit()
    
    # Print results summary
    print(f"\n=== FLEXIBLE MATCHING RESULTS ===")
    print(f"Total PMPs assigned: {len(final_matches)}")
    print(f"Total projects with assignments: {len(assigned_charities)}")
    print(f"Results saved to: {output_file}")
    
    print(f"\n=== TEAM ASSIGNMENTS ===")
    for charity_id, matches in assigned_charities.items():
        charity_name = matches[0]['Charity_Project']['Organization']
        team_size = len(matches)
        priority = matches[0]['Charity_Project']['Priority_Level']
        complexity = matches[0]['Charity_Project']['Complexity']
        
        print(f"\n{charity_name} ({priority} Priority, {complexity} Complexity):")
        print(f"  Team Size: {team_size} PMPs")
        for i, match in enumerate(matches, 1):
            print(f"    {i}. {match['PMP_Name']} (Score: {match['Score']:.2f})")
    
    # Check if all PMPs are assigned
    assigned_pmp_count = len(set(match['PMP_ID'] for match in final_matches))
    print(f"\n=== VERIFICATION ===")
    print(f"PMPs assigned: {assigned_pmp_count}/{len(pmp_profiles)}")
    if assigned_pmp_count == len(pmp_profiles):
        print("✅ All PMPs successfully assigned!")
    else:
        print("⚠️  Some PMPs may not be assigned - check the results.")


if __name__ == "__main__":
    main()