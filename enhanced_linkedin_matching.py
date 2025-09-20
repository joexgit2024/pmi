import pandas as pd
import numpy as np
import re
from urllib.parse import urlparse

def enhance_pmp_profiles_with_linkedin(pmp_df):
    """
    Enhance PMP profiles with LinkedIn URL information for better matching.
    This approach uses only the URL structure and manual validation rather than scraping.
    """
    
    enhanced_profiles = []
    
    for idx, row in pmp_df.iterrows():
        profile = {
            'ID': idx,
            'Name': f"{row['First Name']} {row['Last Name']}",
            'Email': row['Email address'],
            'LinkedIn_URL': row.get('LinkedIn Profile URL', ''),
            'Experience': row['Year(s) as a Project Professional'],
            'Areas_of_Interest': row['Areas of Interest'],
            'Skills': {},
            'LinkedIn_Quality_Score': 0,
            'Profile_Completeness_Score': 0
        }
        
        # Analyze LinkedIn URL quality
        linkedin_url = str(row.get('LinkedIn Profile URL', ''))
        profile['LinkedIn_Quality_Score'] = analyze_linkedin_url_quality(linkedin_url)
        
        # Calculate profile completeness
        profile['Profile_Completeness_Score'] = calculate_profile_completeness(row)
        
        # Extract existing skills (unchanged from original)
        skill_columns = [
            'Project Management', 'Strategic Planning', 'Business Change Management',
            'Business Analysis', 'Portfolio Management', 'Development of User Requirements',
            'Technology Change Management', 'Understanding of Agile Principles',
            'Plan and Manage Agile Projects', 
            'Planning & Management of the Implementation of New Software Solutions',
            'Volunteering for a Non-profit Organisation', 'Events Planning and Management',
            'Systems Integration (Business and Technical)'
        ]
        
        for skill in skill_columns:
            try:
                rating = float(row[skill]) if pd.notna(row[skill]) else 0
                profile['Skills'][skill] = rating
            except:
                profile['Skills'][skill] = 0
        
        # Calculate weighted overall score (considering LinkedIn presence)
        base_score = sum(profile['Skills'].values()) / len(skill_columns)
        linkedin_bonus = profile['LinkedIn_Quality_Score'] * 0.1  # 10% bonus for good LinkedIn
        completeness_bonus = profile['Profile_Completeness_Score'] * 0.05  # 5% bonus for completeness
        
        profile['Overall_Score'] = base_score + linkedin_bonus + completeness_bonus
        
        enhanced_profiles.append(profile)
    
    return enhanced_profiles

def analyze_linkedin_url_quality(linkedin_url):
    """
    Analyze LinkedIn URL quality without scraping - based on URL structure only.
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
    
    # Essential fields
    if pd.notna(row.get('First Name', '')) and row.get('First Name', '') != '':
        score += 1
    if pd.notna(row.get('Last Name', '')) and row.get('Last Name', '') != '':
        score += 1
    if pd.notna(row.get('Email address', '')) and row.get('Email address', '') != '':
        score += 1
    
    # Professional fields
    if pd.notna(row.get('Current / Latest Job Title', '')) and row.get('Current / Latest Job Title', '') != '':
        score += 1
    if pd.notna(row.get('Company', '')) and row.get('Company', '') != '':
        score += 1
    if pd.notna(row.get('PMI ID Number', '')) and row.get('PMI ID Number', '') != '':
        score += 1
    
    # Experience and interests
    if pd.notna(row.get('Year(s) as a Project Professional', '')) and row.get('Year(s) as a Project Professional', '') != '':
        score += 1
    if pd.notna(row.get('Areas of Interest', '')) and row.get('Areas of Interest', '') != '':
        score += 1
    
    # LinkedIn presence
    if pd.notna(row.get('LinkedIn Profile URL', '')) and row.get('LinkedIn Profile URL', '') != '':
        score += 1
    
    # Skills completion (at least half filled)
    skill_columns = ['Project Management', 'Strategic Planning', 'Business Change Management',
                    'Business Analysis', 'Portfolio Management']
    filled_skills = sum(1 for skill in skill_columns if pd.notna(row.get(skill, '')) and row.get(skill, '') != '')
    if filled_skills >= len(skill_columns) // 2:
        score += 1
    
    return score

def enhanced_match_score_calculation(pmp_profile, charity_project):
    """
    Enhanced matching algorithm that considers LinkedIn quality and profile completeness.
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
            'Profile_Complete': 'High' if profile['Profile_Completeness_Score'] >= 8 else 
                              'Medium' if profile['Profile_Completeness_Score'] >= 6 else 'Low'
        })
    
    return pd.DataFrame(linkedin_data)

def validate_linkedin_urls(pmp_df):
    """
    Validate and standardize LinkedIn URLs without scraping.
    """
    validation_results = []
    
    for idx, row in pmp_df.iterrows():
        linkedin_url = str(row.get('LinkedIn Profile URL', ''))
        
        result = {
            'Name': f"{row['First Name']} {row['Last Name']}",
            'Original_URL': linkedin_url,
            'Is_Valid': False,
            'Standardized_URL': '',
            'Issues': []
        }
        
        if pd.isna(linkedin_url) or linkedin_url == '' or linkedin_url == 'nan':
            result['Issues'].append('No LinkedIn URL provided')
        else:
            # Basic validation
            url = linkedin_url.lower().strip()
            
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

# Example usage and recommendations
def main_enhanced_matching():
    """
    Main function demonstrating enhanced matching with LinkedIn considerations.
    """
    print("Enhanced PMP-Charity Matching with LinkedIn Analysis")
    print("=" * 60)
    
    # This would integrate with your existing matching system
    # with the enhanced profile analysis and scoring
    pass

if __name__ == "__main__":
    main_enhanced_matching()