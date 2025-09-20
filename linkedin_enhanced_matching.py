"""
Enhanced PMP-Charity Matching with LinkedIn Profile Analysis
=============================================================

This script enhances the existing matching algorithm by incorporating:
1. LinkedIn profile URL validation and quality scoring
2. Profile completeness analysis
3. Improved matching weights based on professional presence

IMPORTANT: This approach respects LinkedIn's terms of service by:
- NOT scraping profile data automatically
- Only analyzing URL structure and format
- Providing framework for manual verification if needed
"""

import pandas as pd
import sys
import os

# Add the current directory to path to import existing matching functions
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

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
    if pd.notna(row.get('First Name', '')) and row.get('First Name', '') != '':
        score += 1
    if pd.notna(row.get('Last Name', '')) and row.get('Last Name', '') != '':
        score += 1
    if pd.notna(row.get('Email address', '')) and row.get('Email address', '') != '':
        score += 1
    
    # Professional fields (3 points)
    if (pd.notna(row.get('Current / Latest Job Title', '')) and 
        row.get('Current / Latest Job Title', '') != ''):
        score += 1
    if pd.notna(row.get('Company', '')) and row.get('Company', '') != '':
        score += 1
    if (pd.notna(row.get('PMI ID Number', '')) and 
        row.get('PMI ID Number', '') != ''):
        score += 1
    
    # Experience and interests (2 points)
    if (pd.notna(row.get('Year(s) as a Project Professional', '')) and 
        row.get('Year(s) as a Project Professional', '') != ''):
        score += 1
    if (pd.notna(row.get('Areas of Interest', '')) and 
        row.get('Areas of Interest', '') != ''):
        score += 1
    
    # LinkedIn presence (1 point)
    if (pd.notna(row.get('LinkedIn Profile URL', '')) and 
        row.get('LinkedIn Profile URL', '') != ''):
        score += 1
    
    # Skills completion (1 point)
    skill_columns = ['Project Management', 'Strategic Planning', 
                     'Business Change Management', 'Business Analysis', 
                     'Portfolio Management']
    filled_skills = sum(1 for skill in skill_columns 
                       if pd.notna(row.get(skill, '')) and row.get(skill, '') != '')
    if filled_skills >= len(skill_columns) // 2:
        score += 1
    
    return score


def enhanced_extract_pmp_skills(pmp_df):
    """
    Enhanced version of the original extract_pmp_skills function 
    that includes LinkedIn analysis.
    """
    
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
    
    # Create enhanced PMP profiles
    pmp_profiles = []
    
    for idx, row in pmp_df.iterrows():
        profile = {
            'ID': idx,
            'Name': f"{row['First Name']} {row['Last Name']}",
            'Experience': row['Year(s) as a Project Professional'],
            'Areas_of_Interest': row['Areas of Interest'],
            'LinkedIn_URL': row.get('LinkedIn Profile URL', ''),
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


def enhanced_calculate_match_score(pmp_profile, charity_project):
    """
    Enhanced matching algorithm that considers LinkedIn quality and 
    profile completeness.
    """
    total_score = 0
    max_possible_score = 0
    
    # Skills matching (60% of total score - reduced from 70%)
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


def validate_linkedin_urls(pmp_df):
    """
    Validate and analyze LinkedIn URLs without scraping.
    Returns DataFrame with validation results.
    """
    validation_results = []
    
    for idx, row in pmp_df.iterrows():
        linkedin_url = str(row.get('LinkedIn Profile URL', ''))
        
        result = {
            'Name': f"{row['First Name']} {row['Last Name']}",
            'Original_URL': linkedin_url,
            'Is_Valid': False,
            'Standardized_URL': '',
            'Quality_Score': 0,
            'Issues': []
        }
        
        if pd.isna(linkedin_url) or linkedin_url == '' or linkedin_url == 'nan':
            result['Issues'].append('No LinkedIn URL provided')
        else:
            # Basic validation
            url = linkedin_url.lower().strip()
            result['Quality_Score'] = analyze_linkedin_url_quality(linkedin_url)
            
            if 'linkedin' not in url:
                result['Issues'].append('Not a LinkedIn URL')
            elif '/in/' not in url:
                result['Issues'].append('Invalid LinkedIn profile format')
            else:
                result['Is_Valid'] = True
                # Standardize URL format
                if not url.startswith('http'):
                    result['Standardized_URL'] = 'https://' + url
                else:
                    result['Standardized_URL'] = linkedin_url
        
        validation_results.append(result)
    
    return pd.DataFrame(validation_results)


def generate_linkedin_analysis_report(pmp_profiles):
    """
    Generate a report on LinkedIn profile quality and completeness.
    """
    linkedin_data = []
    
    for profile in pmp_profiles:
        linkedin_data.append({
            'PMP_Name': profile['Name'],
            'LinkedIn_URL': profile['LinkedIn_URL'],
            'LinkedIn_Quality_Score': profile['LinkedIn_Quality_Score'],
            'Profile_Completeness_Score': profile['Profile_Completeness_Score'],
            'Overall_Score': round(profile['Overall_Score'], 2),
            'LinkedIn_Valid': 'Yes' if profile['LinkedIn_Quality_Score'] > 5 else 'No',
            'Profile_Complete': ('High' if profile['Profile_Completeness_Score'] >= 8 
                               else 'Medium' if profile['Profile_Completeness_Score'] >= 6 
                               else 'Low')
        })
    
    return pd.DataFrame(linkedin_data)


def main():
    """
    Main function to run LinkedIn-enhanced analysis.
    """
    print("Enhanced LinkedIn Profile Analysis for PMP-Charity Matching")
    print("=" * 65)
    
    # Load the PMP data
    pmp_file = ("input/2025 - PMI Sydney Chapter Project Management Day of Service "
                "(PMDoS) 2025 Professional Registration (Responses).xlsx")
    pmp_df = pd.read_excel(pmp_file)
    
    print(f"Loaded {len(pmp_df)} PMP professional profiles")
    
    # Validate LinkedIn URLs
    print("\n1. Validating LinkedIn URLs...")
    validation_df = validate_linkedin_urls(pmp_df)
    
    valid_linkedin_count = validation_df['Is_Valid'].sum()
    print(f"   - Valid LinkedIn URLs: {valid_linkedin_count}/{len(pmp_df)} "
          f"({valid_linkedin_count/len(pmp_df)*100:.1f}%)")
    
    # Generate enhanced profiles
    print("\n2. Generating enhanced PMP profiles...")
    enhanced_profiles = enhanced_extract_pmp_skills(pmp_df)
    
    # Generate LinkedIn analysis report
    print("\n3. Creating LinkedIn analysis report...")
    linkedin_report = generate_linkedin_analysis_report(enhanced_profiles)
    
    # Save results
    output_file = "LinkedIn_Analysis_Report.xlsx"
    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        
        # LinkedIn validation results
        validation_df.to_excel(writer, sheet_name='URL_Validation', index=False)
        
        # LinkedIn quality analysis
        linkedin_report.to_excel(writer, sheet_name='LinkedIn_Analysis', index=False)
        
        # Enhanced profile summary
        profile_summary = pd.DataFrame([{
            'Name': p['Name'],
            'Experience': p['Experience'],
            'LinkedIn_Quality': p['LinkedIn_Quality_Score'],
            'Profile_Completeness': p['Profile_Completeness_Score'],
            'Enhanced_Overall_Score': round(p['Overall_Score'], 2)
        } for p in enhanced_profiles])
        profile_summary.to_excel(writer, sheet_name='Enhanced_Profiles', index=False)
        
        # Summary statistics
        summary_stats = pd.DataFrame([
            ['Total PMPs', len(pmp_df)],
            ['Valid LinkedIn URLs', valid_linkedin_count],
            ['LinkedIn Coverage %', f"{valid_linkedin_count/len(pmp_df)*100:.1f}%"],
            ['Avg LinkedIn Quality Score', 
             f"{linkedin_report['LinkedIn_Quality_Score'].mean():.1f}"],
            ['Avg Profile Completeness', 
             f"{linkedin_report['Profile_Completeness_Score'].mean():.1f}"],
            ['High Quality LinkedIn Profiles', 
             sum(1 for p in enhanced_profiles if p['LinkedIn_Quality_Score'] >= 7)],
        ], columns=['Metric', 'Value'])
        summary_stats.to_excel(writer, sheet_name='Summary', index=False)
    
    print(f"\n4. Results saved to: {output_file}")
    
    # Print summary
    print("\n=== LINKEDIN ANALYSIS SUMMARY ===")
    print(f"LinkedIn Coverage: {valid_linkedin_count}/{len(pmp_df)} PMPs")
    print(f"Average LinkedIn Quality Score: "
          f"{linkedin_report['LinkedIn_Quality_Score'].mean():.1f}/10")
    print(f"Average Profile Completeness: "
          f"{linkedin_report['Profile_Completeness_Score'].mean():.1f}/10")
    
    high_quality_profiles = sum(1 for p in enhanced_profiles 
                               if p['LinkedIn_Quality_Score'] >= 7)
    print(f"High Quality LinkedIn Profiles: {high_quality_profiles}")
    
    # Show recommendations
    print("\n=== RECOMMENDATIONS ===")
    print("1. Profiles with missing/invalid LinkedIn URLs should be contacted for updates")
    print("2. High-quality LinkedIn profiles indicate stronger professional presence")
    print("3. Consider manual verification of LinkedIn profiles for key matches")
    print("4. Use enhanced scoring in your matching algorithm for better results")


if __name__ == "__main__":
    main()