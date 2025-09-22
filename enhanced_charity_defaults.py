"""
Enhanced Charity Matching with Default Skills for Poor Problem Statements
=========================================================================

This script provides solutions for charities that haven't provided adequate
problem statements or project descriptions.
"""

import pandas as pd
from enhanced_pmp_charity_matching import (
    load_and_process_data, 
    analyze_charity_requirements,
    extract_pmp_skills,
    calculate_match_score,
    create_optimal_matching
)


def identify_problematic_charities(charity_projects, min_skill_threshold=5):
    """
    Identify charities with inadequate problem statements based on 
    low total skill weights.
    """
    problematic = []
    
    for project in charity_projects:
        total_skills = sum(project['Required_Skills'].values())
        if total_skills < min_skill_threshold:
            problematic.append({
                'organization': project['Organization'],
                'total_skills': total_skills,
                'description_length': len(project['Description']),
                'project': project
            })
    
    return problematic


def assign_default_skills_by_organization_type(organization_name, description):
    """
    Assign default skill requirements based on organization type and name
    when problem statement is inadequate.
    """
    org_lower = organization_name.lower()
    desc_lower = description.lower()
    
    # Default skill patterns based on organization type
    default_skills = {
        'Project Management': 3,
        'Strategic Planning': 2,
        'Volunteering for a Non-profit Organisation': 4
    }
    
    # Foundation-specific skills
    if 'foundation' in org_lower:
        default_skills.update({
            'Strategic Planning': 5,
            'Events Planning and Management': 3,
            'Business Analysis': 3
        })
    
    # Environmental organizations
    if any(word in org_lower for word in ['environment', 'green', 'climate', 'sustainability']):
        default_skills.update({
            'Strategic Planning': 6,
            'Business Change Management': 4,
            'Project Management': 4
        })
    
    # Community/social services
    if any(word in org_lower for word in ['community', 'social', 'family', 'youth']):
        default_skills.update({
            'Events Planning and Management': 4,
            'Business Analysis': 3,
            'Strategic Planning': 4
        })
    
    # Technology/digital mentions
    if any(word in desc_lower for word in ['website', 'digital', 'software', 'system', 'technology']):
        default_skills.update({
            'Technology Change Management': 6,
            'Planning & Management of the Implementation of New Software Solutions': 5,
            'Systems Integration (Business and Technical)': 4
        })
    
    # Accounting/finance mentions
    if any(word in desc_lower for word in ['accounting', 'finance', 'budget', 'financial']):
        default_skills.update({
            'Business Analysis': 5,
            'Business Change Management': 4,
            'Planning & Management of the Implementation of New Software Solutions': 3
        })
    
    return default_skills


def enhance_charity_requirements_with_defaults(charity_projects):
    """
    Enhance charity requirements by adding default skills for organizations
    with poor problem statements.
    """
    enhanced_projects = []
    
    for project in charity_projects:
        enhanced_project = project.copy()
        total_skills = sum(project['Required_Skills'].values())
        
        # If skill identification is poor, use defaults
        if total_skills < 5:
            print(f"⚠️  Enhancing {project['Organization']} with default skills...")
            
            default_skills = assign_default_skills_by_organization_type(
                project['Organization'], 
                project['Description']
            )
            
            # Merge with existing skills (take maximum)
            for skill, default_weight in default_skills.items():
                current_weight = enhanced_project['Required_Skills'].get(skill, 0)
                enhanced_project['Required_Skills'][skill] = max(current_weight, default_weight)
            
            # Update priority and complexity
            enhanced_project['Priority_Level'] = 'Medium'  # Default to medium
            if enhanced_project['Complexity'] == 'Low':
                enhanced_project['Complexity'] = 'Medium'
        
        enhanced_projects.append(enhanced_project)
    
    return enhanced_projects


def generate_enhanced_matching_report():
    """
    Generate a matching report that handles charities with poor problem statements.
    """
    print("Enhanced Charity Matching with Default Skills")
    print("=" * 50)
    
    # Load data
    pmp_df, charity_df = load_and_process_data()
    pmp_profiles = extract_pmp_skills(pmp_df)
    charity_projects = analyze_charity_requirements(charity_df)
    
    # Identify problematic charities
    problematic = identify_problematic_charities(charity_projects)
    
    if problematic:
        print(f"\nFound {len(problematic)} charities with inadequate problem statements:")
        for p in problematic:
            print(f"  - {p['organization']} (skill weight: {p['total_skills']})")
        
        # Enhance with default skills
        print("\nEnhancing with default skills...")
        enhanced_projects = enhance_charity_requirements_with_defaults(charity_projects)
        
        # Compare before and after
        print("\n=== BEFORE vs AFTER ENHANCEMENT ===")
        for i, (original, enhanced) in enumerate(zip(charity_projects, enhanced_projects)):
            org_name = original['Organization']
            original_skills = sum(original['Required_Skills'].values())
            enhanced_skills = sum(enhanced['Required_Skills'].values())
            
            if enhanced_skills > original_skills:
                print(f"{org_name}:")
                print(f"  Original skill weight: {original_skills}")
                print(f"  Enhanced skill weight: {enhanced_skills}")
                print(f"  Improvement: +{enhanced_skills - original_skills}")
                print()
        
        # Run matching with enhanced requirements
        print("Running matching with enhanced requirements...")
        final_matches, assigned_charities = create_optimal_matching(pmp_profiles, enhanced_projects)
        
        # Save enhanced results
        from enhanced_pmp_charity_matching import generate_matching_report
        matching_summary = generate_matching_report(final_matches, assigned_charities)
        
        output_file = 'PMI_PMP_Charity_Matching_Enhanced_Defaults.xlsx'
        with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
            matching_summary.to_excel(writer, sheet_name='Enhanced_Matching', index=False)
            
            # Add enhancement details
            enhancement_details = pd.DataFrame([{
                'Organization': p['organization'],
                'Original_Skill_Weight': p['total_skills'],
                'Enhanced_Skill_Weight': sum(enhanced_projects[i]['Required_Skills'].values()),
                'Description_Length': p['description_length'],
                'Enhancement_Applied': 'Yes' if p['total_skills'] < 5 else 'No'
            } for i, p in enumerate(problematic)])
            
            enhancement_details.to_excel(writer, sheet_name='Enhancement_Details', index=False)
        
        print(f"\nEnhanced results saved to: {output_file}")
        
    else:
        print("\nAll charities have adequate problem statements. No enhancement needed.")
    
    return enhanced_projects if problematic else charity_projects


def main():
    """Main function to run enhanced charity matching."""
    enhanced_projects = generate_enhanced_matching_report()
    
    print("\n=== RECOMMENDATIONS FOR FUTURE ===")
    print("1. Contact charities with poor descriptions for better requirements")
    print("2. Use organization type-based defaults as fallback")
    print("3. Consider manual review of enhanced matches")
    print("4. Update default skill patterns based on experience")


if __name__ == "__main__":
    main()