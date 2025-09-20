import pandas as pd

def create_summary_report():
    """Create a text summary of the matching results"""
    
    # Read the Excel file
    excel_file = 'PMI_PMP_Charity_Matching_Results.xlsx'
    
    try:
        # Read all sheets
        matching_summary = pd.read_excel(excel_file, sheet_name='Matching_Summary')
        detailed_analysis = pd.read_excel(excel_file, sheet_name='Detailed_Analysis')
        
        # Create summary report
        with open('Matching_Results_Summary.txt', 'w', encoding='utf-8') as f:
            f.write("="*80 + "\n")
            f.write("PMI SYDNEY CHAPTER - PMP PROFESSIONALS TO CHARITY PROJECTS MATCHING\n")
            f.write("="*80 + "\n\n")
            
            f.write("EXECUTIVE SUMMARY:\n")
            f.write("-" * 40 + "\n")
            f.write(f"• Total PMP Professionals: 20\n")
            f.write(f"• Total Charity Projects: 10\n")
            f.write(f"• Assignment Ratio: 2 PMPs per project\n")
            f.write(f"• Matching Algorithm: Skills-based with experience weighting\n\n")
            
            f.write("DETAILED MATCHING RESULTS:\n")
            f.write("-" * 40 + "\n\n")
            
            # Group by organization
            for org in matching_summary['Charity_Organization'].unique():
                org_data = matching_summary[matching_summary['Charity_Organization'] == org]
                
                f.write(f"PROJECT: {org}\n")
                f.write(f"Initiative: {org_data.iloc[0]['Charity_Initiative']}\n")
                f.write(f"Description: {org_data.iloc[0]['Project_Description']}\n")
                f.write(f"Priority: {org_data.iloc[0]['Project_Priority']} | Complexity: {org_data.iloc[0]['Project_Complexity']}\n")
                f.write(f"Required Skills: {org_data.iloc[0]['Required_Skills']}\n\n")
                
                f.write("ASSIGNED PMPs:\n")
                for idx, row in org_data.iterrows():
                    f.write(f"  {row['PMP_Role']}: {row['PMP_Name']}\n")
                    f.write(f"    Experience: {row['PMP_Experience']}\n")
                    f.write(f"    Match Score: {row['Match_Score']}%\n")
                    f.write(f"    Top Skills: {row['PMP_Top_Skills']}\n")
                    f.write(f"    Overall Rating: {row['Overall_PMP_Rating']}/5\n\n")
                
                f.write("-" * 60 + "\n\n")
            
            f.write("MATCHING METHODOLOGY:\n")
            f.write("-" * 40 + "\n")
            f.write("The matching algorithm considers:\n")
            f.write("1. Skill Alignment (70%): PMP skill ratings (1-5) vs project requirements\n")
            f.write("2. Experience Level (20%): Years as project professional\n")
            f.write("3. Interest Alignment (10%): Stated areas of interest vs project type\n\n")
            
            f.write("Key matching criteria:\n")
            f.write("• Project requirements analyzed from descriptions and outcomes\n")
            f.write("• Skills weighted based on keyword frequency and relevance\n")
            f.write("• Experience bonus for senior professionals (8+ years)\n")
            f.write("• Non-profit interest bonus for charity work\n")
            f.write("• Optimal assignment ensuring 2 PMPs per project\n\n")
            
            # Add reasoning from detailed analysis
            f.write("SELECTION REASONING:\n")
            f.write("-" * 40 + "\n")
            for idx, row in detailed_analysis.iterrows():
                f.write(f"{row['Organization']}:\n")
                f.write(f"  {row['Selection_Reasoning']}\n\n")
        
        print("Summary report created: Matching_Results_Summary.txt")
        
        # Also create a simple CSV for easy viewing
        matching_summary.to_csv('Matching_Summary.csv', index=False)
        print("CSV file created: Matching_Summary.csv")
        
    except Exception as e:
        print(f"Error creating summary: {e}")

if __name__ == "__main__":
    create_summary_report()
