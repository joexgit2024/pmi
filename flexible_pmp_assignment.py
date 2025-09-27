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
    build_match_score_matrix,
    _normalize_company_name
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
    significant_skills = charity_project['Required_Skills'].values()
    skill_count = sum(1 for weight in significant_skills if weight > 2)
    # Extra PMP for every 3 significant skills
    skill_bonus = min(skill_count // 3, 1)

    total_capacity = (
        base_capacity + complexity_bonus + priority_bonus + skill_bonus
    )
    return min(total_capacity, 4)  # Cap at 4 PMPs max per project


def create_flexible_matching(
    pmp_profiles,
    charity_projects,
    score_matrix=None,
    enforce_unique_company=True
):
    """
    Create flexible matching that assigns ALL PMPs to projects.
    Some projects may get 3+ PMPs based on complexity and priority.
    """

    # Calculate all possible match scores
    if score_matrix is None:
        score_matrix = build_match_score_matrix(pmp_profiles, charity_projects)

    all_matches = []
    for pmp in pmp_profiles:
        for charity in charity_projects:
            score = score_matrix[pmp['ID']][charity['ID']]
            all_matches.append({
                'PMP_ID': pmp['ID'],
                'PMP_Name': pmp['Name'],
                'Charity_ID': charity['ID'],
                'Organization': charity['Organization'],
                'Initiative': charity['Initiative'],
                'Score': score,
                'PMP_Profile': pmp,
                'Charity_Project': charity,
                'Company_Key': _normalize_company_name(
                    pmp.get('Company'),
                    pmp['ID']
                )
            })
    
    # Sort by score (highest first)
    all_matches.sort(key=lambda x: x['Score'], reverse=True)
    
    # Calculate project capacities
    project_capacities = {}
    for charity in charity_projects:
        capacity = calculate_project_capacity_score(charity)
        project_capacities[charity['ID']] = {
            'max_capacity': max(capacity, 2),  # Ensure minimum 2 PMPs
            'min_capacity': 2,  # Minimum 2 PMPs for risk management
            'current_assignments': 0,
            'assigned_pmps': [],
            'companies': set()
        }
    
    # Assign PMPs using flexible algorithm with minimum requirements
    assigned_pmps = set()
    final_matches = []

    def _assign(match, state, assigned_set, output_list):
        state['current_assignments'] += 1
        state['assigned_pmps'].append(match)
        state['companies'].add(match['Company_Key'])
        assigned_set.add(match['PMP_ID'])
        output_list.append(match)
    
    print("=== PROJECT CAPACITY ANALYSIS ===")
    for charity in charity_projects:
        capacity = project_capacities[charity['ID']]['max_capacity']
        min_capacity = project_capacities[charity['ID']]['min_capacity']
        org_name = charity['Organization']
        print(f"{org_name}: Capacity {min_capacity}-{capacity} PMPs "
              f"(min {min_capacity} for risk management)")
        print(f"  - Priority: {charity['Priority_Level']}")
        print(f"  - Complexity: {charity['Complexity']}")
        skill_count = sum(
            1 for weight in charity['Required_Skills'].values() if weight > 2
        )
        print(f"  - Skill requirements: {skill_count} significant skills")
        print()
    
    # Phase 1: Ensure each project gets minimum 2 PMPs
    print("=== PHASE 1: Ensuring minimum 2 PMPs per project ===")
    projects_needing_min = list(charity_projects)
    
    for project in projects_needing_min:
        charity_id = project['ID']
        state = project_capacities[charity_id]
        min_capacity = state['min_capacity']
        project_matches = [
            match for match in all_matches
            if match['Charity_ID'] == charity_id
            and match['PMP_ID'] not in assigned_pmps
        ]

        # First, try to satisfy minimum capacity with unique companies
        for match in project_matches:
            if state['current_assignments'] >= min_capacity:
                break

            if match['PMP_ID'] in assigned_pmps:
                continue

            if (
                enforce_unique_company
                and match['Company_Key'] in state['companies']
            ):
                continue

            pmp_name = match['PMP_Name']
            org_name = project['Organization']
            _assign(match, state, assigned_pmps, final_matches)
            assignment_msg = (
                f"  Assigned {pmp_name} to {org_name}"
                " (min requirement)"
            )
            print(assignment_msg)

        # If still short, allow duplicates to reach minimum
        if state['current_assignments'] < min_capacity:
            for match in project_matches:
                if state['current_assignments'] >= min_capacity:
                    break
                if match['PMP_ID'] in assigned_pmps:
                    continue

                pmp_name = match['PMP_Name']
                org_name = project['Organization']
                _assign(match, state, assigned_pmps, final_matches)
                assignment_msg = (
                    f"  Assigned {pmp_name} to {org_name}"
                    " (min requirement - duplicate company)"
                )
                print(assignment_msg)
    
    # Phase 2: Assign remaining PMPs to projects with available capacity
    print("\n=== PHASE 2: Assigning remaining PMPs based on capacity ===")
    remaining_matches = [
        match for match in all_matches if match['PMP_ID'] not in assigned_pmps
    ]

    deferred_matches = []

    for match in remaining_matches:
        charity_id = match['Charity_ID']
        state = project_capacities[charity_id]

        if match['PMP_ID'] in assigned_pmps:
            continue
        if state['current_assignments'] >= state['max_capacity']:
            continue

        if (
            enforce_unique_company
            and match['Company_Key'] in state['companies']
        ):
            deferred_matches.append(match)
            continue

        _assign(match, state, assigned_pmps, final_matches)
        org_name = match['Organization']
        assignment_msg = (
            f"  Assigned {match['PMP_Name']} to {org_name}"
            " (additional capacity)"
        )
        print(assignment_msg)

    # Process deferred matches allowing duplicates if capacity remains
    for match in deferred_matches:
        charity_id = match['Charity_ID']
        state = project_capacities[charity_id]

        if match['PMP_ID'] in assigned_pmps:
            continue
        if state['current_assignments'] >= state['max_capacity']:
            continue

        _assign(match, state, assigned_pmps, final_matches)
        org_name = match['Organization']
        assignment_msg = (
            f"  Assigned {match['PMP_Name']} to {org_name}"
            " (additional capacity - duplicate company)"
        )
        print(assignment_msg)
    # Check if all PMPs are assigned
    unassigned_pmps = [p for p in pmp_profiles if p['ID'] not in assigned_pmps]
    
    if unassigned_pmps:
        print(
            "=== SECOND PASS: Assigning "
            f"{len(unassigned_pmps)} remaining PMPs ==="
        )

        # Second pass: Assign remaining PMPs to projects with
        # lowest current assignment ratio
        for pmp in unassigned_pmps:
            # Find best available match for this PMP
            best_match = None
            best_score = 0
            
            for charity in charity_projects:
                # Calculate current assignment ratio
                current = project_capacities[charity['ID']][
                    'current_assignments'
                ]
                max_cap = project_capacities[charity['ID']]['max_capacity']
                
                # Allow exceeding capacity if needed, but prefer
                # projects with available space
                preference_score = (max_cap - current) * 10
                
                # Find the match score for this PMP-charity combination
                for match in all_matches:
                    if (
                        match['PMP_ID'] == pmp['ID']
                        and match['Charity_ID'] == charity['ID']
                    ):
                        
                        adjusted_score = match['Score'] + preference_score
                        
                        if adjusted_score > best_score:
                            best_score = adjusted_score
                            best_match = match
                            break
            
            if best_match:
                charity_id = best_match['Charity_ID']
                state = project_capacities[charity_id]
                _assign(best_match, state, assigned_pmps, final_matches)
                pmp_name = best_match['PMP_Name']
                org_name = best_match['Organization']
                print(
                    f"  Assigned {pmp_name} to {org_name}"
                    f" (Score: {best_match['Score']:.2f})"
                )
    
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
            top_skills = sorted(
                pmp_info['Skills'].items(),
                key=lambda item: item[1],
                reverse=True
            )[:3]
            top_skills_str = ", ".join(
                [f"{skill}: {rating}" for skill, rating in top_skills]
            )
            
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
    final_matches, assigned_charities = create_flexible_matching(
        pmp_profiles,
        charity_projects
    )
    
    # Generate reports
    print("\nGenerating reports...")
    matching_summary = generate_flexible_matching_report(
        final_matches,
        assigned_charities
    )
    
    # Save results
    output_file = 'PMI_PMP_Charity_Flexible_Matching_Results.xlsx'
    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        
        # Main matching results
        matching_summary.to_excel(
            writer,
            sheet_name='Flexible_Matching',
            index=False
        )
        
        # Team composition summary
        team_summary = matching_summary.groupby('Charity_Organization').agg({
            'Team_Size': 'first',
            'Project_Priority': 'first',
            'Project_Complexity': 'first',
            'PMP_Name': lambda x: ' | '.join(x),
            'Match_Score': ['mean', 'min', 'max']
        }).round(2)
        
        team_summary.columns = [
            'Team_Size',
            'Priority',
            'Complexity',
            'Team_Members',
            'Avg_Score',
            'Min_Score',
            'Max_Score'
        ]
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
                'Capacity_Utilization': (
                    f"{(actual_assignments / max_capacity) * 100:.0f}%"
                ),
                'Priority': charity_info['Priority_Level'],
                'Complexity': charity_info['Complexity'],
                'Over_Capacity': (
                    'Yes' if actual_assignments > max_capacity else 'No'
                )
            })
        
        capacity_df = pd.DataFrame(capacity_analysis)
        capacity_df.to_excel(
            writer,
            sheet_name='Capacity_Analysis',
            index=False
        )
        
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
    print("\n=== FLEXIBLE MATCHING RESULTS ===")
    print(f"Total PMPs assigned: {len(final_matches)}")
    print(f"Total projects with assignments: {len(assigned_charities)}")
    print(f"Results saved to: {output_file}")
    
    print("\n=== TEAM ASSIGNMENTS ===")
    for charity_id, matches in assigned_charities.items():
        charity_name = matches[0]['Charity_Project']['Organization']
        team_size = len(matches)
        priority = matches[0]['Charity_Project']['Priority_Level']
        complexity = matches[0]['Charity_Project']['Complexity']
        
        print(
            "\n"
            f"{charity_name} ({priority} Priority, {complexity} Complexity):"
        )
        print(f"  Team Size: {team_size} PMPs")
        for i, match in enumerate(matches, 1):
            score_display = f"{match['Score']:.2f}"
            print(
                f"    {i}. {match['PMP_Name']} (Score: {score_display})"
            )
    
    # Check if all PMPs are assigned
    assigned_pmp_count = len(set(match['PMP_ID'] for match in final_matches))
    print("\n=== VERIFICATION ===")
    print(f"PMPs assigned: {assigned_pmp_count}/{len(pmp_profiles)}")
    if assigned_pmp_count == len(pmp_profiles):
        print("✅ All PMPs successfully assigned!")
    else:
        print("⚠️  Some PMPs may not be assigned - check the results.")


if __name__ == "__main__":
    main()
