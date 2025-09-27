# flake8: noqa
"""
Enhanced PMP-Charity Matching Script with LinkedIn Integration
=============================================================

This enhanced version of the original matching script incorporates:
1. LinkedIn profile quality analysis
2. Profile completeness scoring
3. Improved match scoring algorithm
4. Additional reporting on professional presence

Based on the original pmp_charity_matching.py with LinkedIn enhancements.
"""

import pandas as pd
import numpy as np


QUALIFIED_SCORE_THRESHOLD = 65.0
BACKUP_SCORE_THRESHOLD = 50.0


def _normalize_company_name(company_raw, fallback_id):
    """Return a normalized company key for assignment checks."""
    company = str(company_raw or '').strip()
    if not company:
        return f"__no_company__{fallback_id}"
    return company.lower()


def build_match_score_matrix(pmp_profiles, charity_projects):
    """Precompute match scores for every PMP-charity combination."""
    score_matrix = {}
    for pmp in pmp_profiles:
        pmp_scores = {}
        for charity in charity_projects:
            score = calculate_match_score(pmp, charity)
            pmp_scores[charity['ID']] = score
        score_matrix[pmp['ID']] = pmp_scores
    return score_matrix


def categorize_pmp_candidates(pmp_profiles, charity_projects,
                              qualified_threshold=QUALIFIED_SCORE_THRESHOLD,
                              backup_threshold=BACKUP_SCORE_THRESHOLD):
    """
    Split PMP profiles into qualified, backup and non-selected lists
    based on their best available match score.
    """

    score_matrix = build_match_score_matrix(pmp_profiles, charity_projects)
    qualified = []
    backup = []
    non_selected = []
    best_scores = {}

    for pmp in pmp_profiles:
        pmp_scores = score_matrix.get(pmp['ID'], {})
        if pmp_scores:
            best_charity_id, best_score = max(
                pmp_scores.items(), key=lambda item: item[1]
            )
        else:
            best_charity_id, best_score = (None, 0.0)

        best_scores[pmp['ID']] = {
            'best_score': best_score,
            'best_charity_id': best_charity_id
        }

        if best_score >= qualified_threshold:
            qualified.append(pmp)
        elif best_score >= backup_threshold:
            backup.append(pmp)
        else:
            non_selected.append(pmp)

    return qualified, backup, non_selected, best_scores, score_matrix


def load_and_process_data():
    """Load and process both datasets (with dynamic file detection)"""
    from dynamic_file_loader import get_latest_input_files
    
    # Get latest files dynamically
    pmp_file, charity_file = get_latest_input_files()
    
    if not pmp_file:
        raise FileNotFoundError("Could not find PMP registration file in input/ directory")
    if not charity_file:
        raise FileNotFoundError("Could not find charity information file in input/ directory")
    
    print(f"Loading PMP data from: {pmp_file}")
    print(f"Loading charity data from: {charity_file}")
    
    # Read PMP professionals data
    pmp_df = pd.read_excel(pmp_file)
    
    # Read charity projects data
    charity_df = pd.read_excel(charity_file)
    
    return pmp_df, charity_df


def analyze_linkedin_url_quality(linkedin_url):
    """
    Analyze LinkedIn URL quality based on structure only (no scraping).
    Returns score 0-10 based on URL validity and format.
    """
    if pd.isna(linkedin_url) or linkedin_url == '':
        return 0
    
    url = str(linkedin_url).lower().strip()
    score = 0
    
    # Check if it contains linkedin
    if 'linkedin' in url:
        score += 3
    
    # Check for proper linkedin.com domain
    if 'linkedin.com/in/' in url:
        score += 4
    
    # Check for https (security consideration)
    if url.startswith('https://'):
        score += 2
    
    # Check for custom profile name (vs default numbers)
    if '/in/' in url:
        profile_part = url.split('/in/')[-1].rstrip('/')
        if profile_part and not profile_part.isdigit():
            score += 1
    
    return min(score, 10)


def calculate_profile_completeness(row):
    """
    Calculate how complete a PMP profile is based on provided information.
    Returns score 0-10.
    """
    score = 0
    
    # Essential fields (3 points)
    essential_fields = ['First Name', 'Last Name', 'Email address']
    for field in essential_fields:
        if pd.notna(row.get(field, '')) and str(row.get(field, '')).strip():
            score += 1
    
    # Professional fields (3 points)
    professional_fields = ['Current / Latest Job Title', 'Company', 'PMI ID Number']
    for field in professional_fields:
        if pd.notna(row.get(field, '')) and str(row.get(field, '')).strip():
            score += 1
    
    # Experience and interests (2 points)
    experience_fields = ['Year(s) as a Project Professional', 'Areas of Interest']
    for field in experience_fields:
        if pd.notna(row.get(field, '')) and str(row.get(field, '')).strip():
            score += 1
    
    # LinkedIn presence (1 point)
    if pd.notna(row.get('LinkedIn Profile URL', '')) and str(row.get('LinkedIn Profile URL', '')).strip():
        score += 1
    
    # Skills completion (1 point)
    skill_columns = ['Project Management', 'Strategic Planning', 
                     'Business Change Management', 'Business Analysis', 
                     'Portfolio Management']
    filled_skills = sum(1 for skill in skill_columns 
                       if pd.notna(row.get(skill, '')) and str(row.get(skill, '')).strip())
    if filled_skills >= len(skill_columns) // 2:
        score += 1
    
    return score


def extract_pmp_skills(pmp_df):
    """Extract and process PMP professional skills with LinkedIn enhancement"""
    
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
            'LinkedIn_URL': row.get('LinkedIn Profile URL', ''),
            'Company': row.get('Company', ''),
            'Job_Title': row.get('Current / Latest Job Title', ''),
            'Email': row.get('Email address', ''),
            'Skills': {},
            'LinkedIn_Quality_Score': 0,
            'Profile_Completeness_Score': 0
        }
        
        # Analyze LinkedIn URL quality
        linkedin_url = str(row.get('LinkedIn Profile URL', ''))
        profile['LinkedIn_Quality_Score'] = analyze_linkedin_url_quality(linkedin_url)
        
        # Calculate profile completeness
        profile['Profile_Completeness_Score'] = calculate_profile_completeness(row)
        
        # Extract skill ratings
        for skill in skill_columns:
            try:
                rating = float(row[skill]) if pd.notna(row[skill]) else 0
                profile['Skills'][skill] = rating
            except (ValueError, TypeError):
                profile['Skills'][skill] = 0
        
        # Calculate enhanced overall skill score
        base_score = sum(profile['Skills'].values()) / len(skill_columns)
        
        # Add bonuses for LinkedIn presence and profile completeness
        linkedin_bonus = profile['LinkedIn_Quality_Score'] * 0.1  # 10% bonus
        completeness_bonus = profile['Profile_Completeness_Score'] * 0.05  # 5% bonus
        
        profile['Overall_Score'] = base_score + linkedin_bonus + completeness_bonus
        
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
    """Calculate enhanced match score between PMP professional and charity project"""
    
    total_score = 0
    max_possible_score = 0
    
    # Skills matching (60% of total score - reduced from 70% to make room for new factors)
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
    
    # NEW: LinkedIn Quality bonus (5% of total score)
    linkedin_bonus = (pmp_profile['LinkedIn_Quality_Score'] / 10) * 3
    total_score += linkedin_bonus
    max_possible_score += 3
    
    # NEW: Profile Completeness bonus (5% of total score)
    completeness_bonus = (pmp_profile['Profile_Completeness_Score'] / 10) * 2
    total_score += completeness_bonus
    max_possible_score += 2
    
    # Normalize score to 0-100
    if max_possible_score > 0:
        normalized_score = (total_score / max_possible_score) * 100
    else:
        normalized_score = 0
    
    return normalized_score


def create_optimal_matching(pmp_profiles, charity_projects,
                            score_matrix=None,
                            enforce_unique_company=True,
                            max_per_project=2):
    """Create baseline matching ensuring company diversity."""

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
                'Charity_Project': charity
            })

    all_matches.sort(key=lambda x: x['Score'], reverse=True)

    assigned_pmps = set()
    charity_state = {
        charity['ID']: {
            'assignments': [],
            'companies': set()
        }
        for charity in charity_projects
    }

    def _assign(match_item):
        charity_id = match_item['Charity_ID']
        company_key = _normalize_company_name(
            match_item['PMP_Profile'].get('Company'),
            match_item['PMP_ID']
        )
        charity_state[charity_id]['assignments'].append(match_item)
        charity_state[charity_id]['companies'].add(company_key)
        assigned_pmps.add(match_item['PMP_ID'])

    # Pass 1: enforce unique company within each project
    for match in all_matches:
        charity_id = match['Charity_ID']
        pmp_id = match['PMP_ID']
        state = charity_state[charity_id]

        if pmp_id in assigned_pmps:
            continue
        if len(state['assignments']) >= max_per_project:
            continue

        company_key = _normalize_company_name(
            match['PMP_Profile'].get('Company'),
            pmp_id
        )

        if enforce_unique_company and company_key in state['companies']:
            continue

        _assign(match)

    # Pass 2: fill remaining slots even if company duplicates are required
    for match in all_matches:
        charity_id = match['Charity_ID']
        pmp_id = match['PMP_ID']
        state = charity_state[charity_id]

        if pmp_id in assigned_pmps:
            continue
        if len(state['assignments']) >= max_per_project:
            continue

        _assign(match)

    final_matches = []
    assigned_charities = {}
    for charity_id, state in charity_state.items():
        if state['assignments']:
            assigned_charities[charity_id] = state['assignments']
            final_matches.extend(state['assignments'])

    return final_matches, assigned_charities


def generate_matching_report(final_matches, assigned_charities):
    """Generate detailed matching report with LinkedIn information"""
    
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
                'PMP_Company': pmp_info.get('Company', ''),
                'PMP_Job_Title': pmp_info.get('Job_Title', ''),
                'LinkedIn_URL': pmp_info.get('LinkedIn_URL', ''),
                'LinkedIn_Quality': pmp_info.get('LinkedIn_Quality_Score', 0),
                'Profile_Completeness': pmp_info.get('Profile_Completeness_Score', 0),
                'Match_Score': round(match['Score'], 2),
                'PMP_Top_Skills': top_skills_str,
                'Required_Skills': top_required_str,
                'Overall_PMP_Rating': round(pmp_info['Overall_Score'], 2)
            })
    
    return pd.DataFrame(match_data)


def create_detailed_analysis(pmp_profiles, charity_projects, final_matches):
    """Create detailed analysis with reasoning including LinkedIn factors"""
    
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
            
            # Add LinkedIn quality information
            if pmp_info.get('LinkedIn_Quality_Score', 0) >= 7:
                pmp_reason += "High-quality LinkedIn profile. "
            
            if pmp_info.get('Profile_Completeness_Score', 0) >= 8:
                pmp_reason += "Complete professional profile. "
            
            reasons.append(pmp_reason)
        
        analysis_data.append({
            'Organization': charity_info['Organization'],
            'Initiative': charity_info['Initiative'],
            'Description': charity_info['Description'],
            'Key_Requirements': ', '.join([skill for skill, weight in charity_info['Required_Skills'].items() if weight > 2]),
            'Assigned_PMPs': ' | '.join([match['PMP_Name'] for match in matches]),
            'Match_Scores': ' | '.join([str(round(match['Score'], 2)) for match in matches]),
            'LinkedIn_Quality': ' | '.join([str(match['PMP_Profile'].get('LinkedIn_Quality_Score', 0)) for match in matches]),
            'Selection_Reasoning': ' | '.join(reasons)
        })
    
    return pd.DataFrame(analysis_data)


def main():
    """Main execution function with LinkedIn enhancement"""
    
    print("Enhanced PMP-Charity Matching with LinkedIn Analysis")
    print("=" * 60)
    
    print("Loading and processing data...")
    pmp_df, charity_df = load_and_process_data()
    
    print("Extracting enhanced PMP skills...")
    pmp_profiles = extract_pmp_skills(pmp_df)
    
    # Quick LinkedIn analysis summary
    linkedin_urls = [p['LinkedIn_URL'] for p in pmp_profiles if p['LinkedIn_URL']]
    avg_linkedin_quality = sum(p['LinkedIn_Quality_Score'] for p in pmp_profiles) / len(pmp_profiles)
    avg_completeness = sum(p['Profile_Completeness_Score'] for p in pmp_profiles) / len(pmp_profiles)
    
    print(f"LinkedIn Coverage: {len(linkedin_urls)}/{len(pmp_profiles)} PMPs ({len(linkedin_urls)/len(pmp_profiles)*100:.1f}%)")
    print(f"Average LinkedIn Quality: {avg_linkedin_quality:.1f}/10")
    print(f"Average Profile Completeness: {avg_completeness:.1f}/10")
    
    print("Analyzing charity requirements...")
    charity_projects = analyze_charity_requirements(charity_df)
    
    print("Creating optimal matching with LinkedIn enhancement...")
    final_matches, assigned_charities = create_optimal_matching(pmp_profiles, charity_projects)
    
    print("Generating enhanced reports...")
    
    # Generate matching summary
    matching_summary = generate_matching_report(final_matches, assigned_charities)
    
    # Generate detailed analysis
    detailed_analysis = create_detailed_analysis(pmp_profiles, charity_projects, final_matches)
    
    # Save to Excel with LinkedIn information
    output_file = 'PMI_PMP_Charity_Matching_Results_Enhanced.xlsx'
    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        
        # Enhanced summary sheet
        matching_summary.to_excel(writer, sheet_name='Enhanced_Matching_Summary', index=False)
        
        # Detailed analysis with LinkedIn
        detailed_analysis.to_excel(writer, sheet_name='Detailed_Analysis_Enhanced', index=False)
        
        # LinkedIn analysis
        linkedin_analysis = pd.DataFrame([{
            'Name': p['Name'],
            'LinkedIn_URL': p['LinkedIn_URL'],
            'LinkedIn_Quality_Score': p['LinkedIn_Quality_Score'],
            'Profile_Completeness_Score': p['Profile_Completeness_Score'],
            'Company': p.get('Company', ''),
            'Job_Title': p.get('Job_Title', ''),
            'Enhanced_Overall_Score': round(p['Overall_Score'], 2)
        } for p in pmp_profiles])
        linkedin_analysis.to_excel(writer, sheet_name='LinkedIn_Analysis', index=False)
        
        # Raw data sheets
        pmp_summary = pd.DataFrame([{
            'ID': p['ID'],
            'Name': p['Name'],
            'Experience': p['Experience'],
            'LinkedIn_Quality': p['LinkedIn_Quality_Score'],
            'Profile_Completeness': p['Profile_Completeness_Score'],
            'Enhanced_Overall_Score': round(p['Overall_Score'], 2),
            'Areas_of_Interest': p['Areas_of_Interest']
        } for p in pmp_profiles])
        pmp_summary.to_excel(writer, sheet_name='Enhanced_PMP_Profiles', index=False)
        
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
    
    print(f"\nEnhanced matching completed successfully!")
    print(f"Total PMPs: {len(pmp_profiles)}")
    print(f"Total Charity Projects: {len(charity_projects)}")
    print(f"Total Matches Created: {len(final_matches)}")
    print(f"\nResults saved to: {output_file}")
    
    # Print enhanced summary statistics
    print("\n=== ENHANCED MATCHING SUMMARY ===")
    for charity_id, matches in assigned_charities.items():
        charity_name = matches[0]['Charity_Project']['Organization']
        pmp_details = []
        for match in matches:
            pmp_info = match['PMP_Profile']
            linkedin_quality = pmp_info.get('LinkedIn_Quality_Score', 0)
            match_score = round(match['Score'], 2)
            pmp_details.append(f"{match['PMP_Name']} (Score: {match_score}, LinkedIn: {linkedin_quality}/10)")
        
        print(f"\n{charity_name}:")
        for detail in pmp_details:
            print(f"  - {detail}")


if __name__ == "__main__":
    main()