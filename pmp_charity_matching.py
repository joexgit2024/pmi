import pandas as pd
import numpy as np

def load_and_process_data():
    """Load and process both datasets"""
    
    # Read PMP professionals data
    pmp_file = r"c:\PMI\input\2025 - PMI Sydney Chapter Project Management Day of Service (PMDoS) 2025 Professional Registration (Responses).xlsx"
    pmp_df = pd.read_excel(pmp_file)
    
    # Read charity projects data
    charity_file = r"c:\PMI\input\Charities Project Information 2025 (Responses).xlsx"
    charity_df = pd.read_excel(charity_file)
    
    return pmp_df, charity_df

def extract_pmp_skills(pmp_df):
    """Extract and process PMP professional skills"""
    
    # Define skill columns (ratings 1-5)
    skill_columns = [
        'Project Management',
        'Strategic Planning', 
        'Business Change Management',
        'Business Analysis',
        'Portfolio Management',
        'Development of User Requirements',
        'Technology Change Management',
        'Understanding of Agile Principles',
        'Plan and Manage Agile Projects',
        'Planning & Management of the Implementation of New Software Solutions',
        'Volunteering for a Non-profit Organisation',
        'Events Planning and Management',
        'Systems Integration (Business and Technical)'
    ]
    
    # Create PMP profiles
    pmp_profiles = []
    
    for idx, row in pmp_df.iterrows():
        profile = {
            'ID': idx,
            'Name': f"{row['First Name']} {row['Last Name']}",
            'Experience': row['Year(s) as a Project Professional'],
            'Areas_of_Interest': row['Areas of Interest'],
            'Skills': {}
        }
        
        # Extract skill ratings
        for skill in skill_columns:
            try:
                rating = float(row[skill]) if pd.notna(row[skill]) else 0
                profile['Skills'][skill] = rating
            except:
                profile['Skills'][skill] = 0
        
        # Calculate overall skill score
        profile['Overall_Score'] = sum(profile['Skills'].values()) / len(skill_columns)
        
        pmp_profiles.append(profile)
    
    return pmp_profiles

def analyze_charity_requirements(charity_df):
    """Analyze charity project requirements and map to required skills"""
    
    charity_projects = []
    
    for idx, row in charity_df.iterrows():
        
        org_name = row['Name of the organisation']
        initiative_name = row['Name of the initiative? ']
        description = str(row['Simple description of the initiative or the project.'])
        outcomes = str(row['What are the key outcomes expected from this initiative or project?'])
        benefits = str(row['How will this initiative benefit your organisation?'])
        expectations = str(row['What outcome(s) do you expect to achieve by participating in this PMDoS event?'])
        
        project = {
            'ID': idx,
            'Organization': org_name,
            'Initiative': initiative_name,
            'Description': description,
            'Outcomes': outcomes,
            'Benefits': benefits,
            'Expectations': expectations,
            'Required_Skills': analyze_project_skill_requirements(
                org_name, initiative_name, description, outcomes, benefits, expectations
            ),
            'Priority_Level': determine_project_priority(description, outcomes),
            'Complexity': assess_project_complexity(description, outcomes)
        }
        
        charity_projects.append(project)
    
    return charity_projects

def analyze_project_skill_requirements(org_name, initiative, description, outcomes, benefits, expectations):
    """Analyze project text to determine required skill weights"""
    
    # Combine all text for analysis
    full_text = f"{org_name} {initiative} {description} {outcomes} {benefits} {expectations}".lower()
    
    # Define skill keywords and their importance weights
    skill_keywords = {
        'Project Management': ['project plan', 'project management', 'timeline', 'deliverable', 'milestone', 'scope', 'budget'],
        'Strategic Planning': ['strategic', 'strategy', 'planning', 'vision', 'mission', 'long-term', 'roadmap', 'alignment'],
        'Business Change Management': ['change', 'transformation', 'transition', 'migration', 'implementation', 'adoption'],
        'Business Analysis': ['analysis', 'requirements', 'process', 'workflow', 'business', 'assessment'],
        'Portfolio Management': ['portfolio', 'program', 'multiple projects', 'prioritisation', 'resource allocation'],
        'Development of User Requirements': ['requirements', 'user needs', 'stakeholder', 'specification', 'functional'],
        'Technology Change Management': ['technology', 'software', 'system', 'digital', 'IT', 'technical'],
        'Understanding of Agile Principles': ['agile', 'iterative', 'flexible', 'adaptive', 'sprint'],
        'Plan and Manage Agile Projects': ['agile project', 'scrum', 'kanban', 'sprint planning'],
        'Planning & Management of the Implementation of New Software Solutions': ['software implementation', 'system implementation', 'ERP', 'accounting software', 'new software'],
        'Volunteering for a Non-profit Organisation': ['non-profit', 'charity', 'volunteer', 'community', 'foundation', 'NGO'],
        'Events Planning and Management': ['event', 'anniversary', 'fundraising', 'celebration', 'conference'],
        'Systems Integration (Business and Technical)': ['integration', 'system', 'platform', 'interface', 'technical']
    }
    
    skill_weights = {}
    
    for skill, keywords in skill_keywords.items():
        weight = 0
        for keyword in keywords:
            weight += full_text.count(keyword) * 2  # Base weight for keyword presence
        
        # Bonus for exact matches in key fields
        if any(keyword in initiative.lower() if isinstance(initiative, str) else "" for keyword in keywords):
            weight += 5
        
        skill_weights[skill] = min(weight, 10)  # Cap at 10
    
    return skill_weights

def determine_project_priority(description, outcomes):
    """Determine project priority based on description and outcomes"""
    text = f"{description} {outcomes}".lower()
    
    high_priority_indicators = ['urgent', 'critical', '50th anniversary', 'strategic', 'foundation']
    medium_priority_indicators = ['important', 'essential', 'significant']
    
    if any(indicator in text for indicator in high_priority_indicators):
        return 'High'
    elif any(indicator in text for indicator in medium_priority_indicators):
        return 'Medium'
    else:
        return 'Low'

def assess_project_complexity(description, outcomes):
    """Assess project complexity"""
    text = f"{description} {outcomes}".lower()
    
    complexity_indicators = {
        'High': ['comprehensive', 'national', 'multiple', 'complex', 'integration', 'strategic'],
        'Medium': ['implementation', 'development', 'planning', 'management'],
        'Low': ['simple', 'basic', 'guidance', 'advice', 'template']
    }
    
    scores = {}
    for level, indicators in complexity_indicators.items():
        scores[level] = sum(1 for indicator in indicators if indicator in text)
    
    return max(scores, key=scores.get) if any(scores.values()) else 'Medium'

def calculate_match_score(pmp_profile, charity_project):
    """Calculate match score between PMP professional and charity project"""
    
    total_score = 0
    max_possible_score = 0
    
    # Skills matching (70% of total score)
    for skill, required_weight in charity_project['Required_Skills'].items():
        if required_weight > 0:
            pmp_skill_level = pmp_profile['Skills'].get(skill, 0)
            skill_score = (pmp_skill_level / 5.0) * required_weight
            total_score += skill_score
            max_possible_score += required_weight
    
    # Experience bonus (20% of total score)
    experience = pmp_profile['Experience']
    if 'More than 8 Years' in str(experience):
        experience_bonus = 10
    elif '4 - 8 Years' in str(experience):
        experience_bonus = 8
    elif '1 - 3 Years' in str(experience):
        experience_bonus = 5
    else:
        experience_bonus = 2
    
    total_score += experience_bonus
    max_possible_score += 10
    
    # Interest alignment bonus (10% of total score)
    interests = str(pmp_profile['Areas_of_Interest']).lower()
    org_name = charity_project['Organization'].lower()
    description = charity_project['Description'].lower()
    
    interest_bonus = 0
    if 'non-profit' in interests or 'volunteer' in interests:
        interest_bonus += 3
    if any(word in interests for word in ['strategic', 'planning', 'change', 'events']):
        interest_bonus += 2
    
    total_score += interest_bonus
    max_possible_score += 5
    
    # Normalize score to 0-100
    if max_possible_score > 0:
        normalized_score = (total_score / max_possible_score) * 100
    else:
        normalized_score = 0
    
    return normalized_score

def create_optimal_matching(pmp_profiles, charity_projects):
    """Create optimal matching using greedy algorithm with constraints"""
    
    # Calculate all possible match scores
    match_matrix = []
    for pmp in pmp_profiles:
        pmp_matches = []
        for charity in charity_projects:
            score = calculate_match_score(pmp, charity)
            pmp_matches.append({
                'PMP_ID': pmp['ID'],
                'PMP_Name': pmp['Name'],
                'Charity_ID': charity['ID'],
                'Organization': charity['Organization'],
                'Initiative': charity['Initiative'],
                'Score': score,
                'PMP_Profile': pmp,
                'Charity_Project': charity
            })
        match_matrix.append(pmp_matches)
    
    # Flatten and sort by score
    all_matches = []
    for pmp_matches in match_matrix:
        all_matches.extend(pmp_matches)
    
    all_matches.sort(key=lambda x: x['Score'], reverse=True)
    
    # Greedy matching - assign 2 PMPs per charity
    assigned_pmps = set()
    assigned_charities = {}  # charity_id: [pmp1, pmp2]
    final_matches = []
    
    for match in all_matches:
        charity_id = match['Charity_ID']
        pmp_id = match['PMP_ID']
        
        # Check if we can assign this PMP to this charity
        if (pmp_id not in assigned_pmps and 
            charity_id not in assigned_charities):
            # First PMP for this charity
            assigned_charities[charity_id] = [match]
            assigned_pmps.add(pmp_id)
            final_matches.append(match)
            
        elif (pmp_id not in assigned_pmps and 
              charity_id in assigned_charities and 
              len(assigned_charities[charity_id]) < 2):
            # Second PMP for this charity
            assigned_charities[charity_id].append(match)
            assigned_pmps.add(pmp_id)
            final_matches.append(match)
    
    return final_matches, assigned_charities

def generate_matching_report(final_matches, assigned_charities):
    """Generate detailed matching report"""
    
    # Create summary DataFrame
    match_data = []
    
    for charity_id, matches in assigned_charities.items():
        charity_info = matches[0]['Charity_Project']
        
        for i, match in enumerate(matches):
            pmp_info = match['PMP_Profile']
            
            # Get top 3 skills for this PMP
            top_skills = sorted(pmp_info['Skills'].items(), key=lambda x: x[1], reverse=True)[:3]
            top_skills_str = ", ".join([f"{skill}: {rating}" for skill, rating in top_skills])
            
            # Get top required skills for charity
            top_required = sorted(charity_info['Required_Skills'].items(), key=lambda x: x[1], reverse=True)[:3]
            top_required_str = ", ".join([f"{skill}: {weight}" for skill, weight in top_required if weight > 0])
            
            match_data.append({
                'Charity_Organization': charity_info['Organization'],
                'Charity_Initiative': charity_info['Initiative'],
                'Project_Description': charity_info['Description'][:100] + "..." if len(charity_info['Description']) > 100 else charity_info['Description'],
                'Project_Priority': charity_info['Priority_Level'],
                'Project_Complexity': charity_info['Complexity'],
                'PMP_Role': f"PMP {i+1}",
                'PMP_Name': pmp_info['Name'],
                'PMP_Experience': pmp_info['Experience'],
                'Match_Score': round(match['Score'], 2),
                'PMP_Top_Skills': top_skills_str,
                'Required_Skills': top_required_str,
                'Overall_PMP_Rating': round(pmp_info['Overall_Score'], 2)
            })
    
    return pd.DataFrame(match_data)

def create_detailed_analysis(pmp_profiles, charity_projects, final_matches):
    """Create detailed analysis with reasoning"""
    
    analysis_data = []
    
    # Group matches by charity
    charity_matches = {}
    for match in final_matches:
        charity_id = match['Charity_ID']
        if charity_id not in charity_matches:
            charity_matches[charity_id] = []
        charity_matches[charity_id].append(match)
    
    for charity_id, matches in charity_matches.items():
        charity_info = matches[0]['Charity_Project']
        
        # Analyze why these PMPs were selected
        reasons = []
        
        for i, match in enumerate(matches):
            pmp_info = match['PMP_Profile']
            
            # Find strongest skill alignments
            skill_alignments = []
            for skill, required_weight in charity_info['Required_Skills'].items():
                if required_weight > 0:
                    pmp_skill = pmp_info['Skills'].get(skill, 0)
                    if pmp_skill >= 4:  # Strong skill
                        skill_alignments.append(f"{skill} (PMP: {pmp_skill}/5, Required: {required_weight})")
            
            pmp_reason = f"PMP {i+1} ({pmp_info['Name']}) selected because: "
            if skill_alignments:
                pmp_reason += f"Strong skills in {'; '.join(skill_alignments[:2])}. "
            
            if 'More than 8 Years' in str(pmp_info['Experience']):
                pmp_reason += "Extensive experience (8+ years). "
            
            if 'non-profit' in str(pmp_info['Areas_of_Interest']).lower():
                pmp_reason += "Interest in non-profit work. "
            
            reasons.append(pmp_reason)
        
        analysis_data.append({
            'Organization': charity_info['Organization'],
            'Initiative': charity_info['Initiative'],
            'Description': charity_info['Description'],
            'Key_Requirements': ', '.join([skill for skill, weight in charity_info['Required_Skills'].items() if weight > 2]),
            'Assigned_PMPs': ' | '.join([match['PMP_Name'] for match in matches]),
            'Match_Scores': ' | '.join([str(round(match['Score'], 2)) for match in matches]),
            'Selection_Reasoning': ' | '.join(reasons)
        })
    
    return pd.DataFrame(analysis_data)

def main():
    """Main execution function"""
    
    print("Loading and processing data...")
    pmp_df, charity_df = load_and_process_data()
    
    print("Extracting PMP skills...")
    pmp_profiles = extract_pmp_skills(pmp_df)
    
    print("Analyzing charity requirements...")
    charity_projects = analyze_charity_requirements(charity_df)
    
    print("Creating optimal matching...")
    final_matches, assigned_charities = create_optimal_matching(pmp_profiles, charity_projects)
    
    print("Generating reports...")
    
    # Generate matching summary
    matching_summary = generate_matching_report(final_matches, assigned_charities)
    
    # Generate detailed analysis
    detailed_analysis = create_detailed_analysis(pmp_profiles, charity_projects, final_matches)
    
    # Save to Excel
    with pd.ExcelWriter('PMI_PMP_Charity_Matching_Results.xlsx', engine='xlsxwriter') as writer:
        
        # Summary sheet
        matching_summary.to_excel(writer, sheet_name='Matching_Summary', index=False)
        
        # Detailed analysis
        detailed_analysis.to_excel(writer, sheet_name='Detailed_Analysis', index=False)
        
        # Raw data sheets
        pmp_summary = pd.DataFrame([{
            'ID': p['ID'],
            'Name': p['Name'],
            'Experience': p['Experience'],
            'Overall_Score': round(p['Overall_Score'], 2),
            'Areas_of_Interest': p['Areas_of_Interest']
        } for p in pmp_profiles])
        pmp_summary.to_excel(writer, sheet_name='PMP_Profiles', index=False)
        
        charity_summary = pd.DataFrame([{
            'ID': c['ID'],
            'Organization': c['Organization'],
            'Initiative': c['Initiative'],
            'Priority': c['Priority_Level'],
            'Complexity': c['Complexity']
        } for c in charity_projects])
        charity_summary.to_excel(writer, sheet_name='Charity_Projects', index=False)
        
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
    
    print(f"\nMatching completed successfully!")
    print(f"Total PMPs: {len(pmp_profiles)}")
    print(f"Total Charity Projects: {len(charity_projects)}")
    print(f"Total Matches Created: {len(final_matches)}")
    print(f"\nResults saved to: PMI_PMP_Charity_Matching_Results.xlsx")
    
    # Print summary statistics
    print("\n=== MATCHING SUMMARY ===")
    for charity_id, matches in assigned_charities.items():
        charity_name = matches[0]['Charity_Project']['Organization']
        pmp_names = [match['PMP_Name'] for match in matches]
        scores = [round(match['Score'], 2) for match in matches]
        print(f"{charity_name}: {pmp_names} (Scores: {scores})")

if __name__ == "__main__":
    main()
