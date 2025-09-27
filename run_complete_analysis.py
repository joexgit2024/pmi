# flake8: noqa
"""
Master PMP-Charity Matching Analysis Script
==========================================

This is the ONLY script you need to run when Excel input files change.
It will automatically execute the complete analysis pipeline in the correct sequence.

Usage: Simply run this script after updating your input Excel files:
    python run_complete_analysis.py

The script will:
1. Validate and analyze LinkedIn profiles
2. Extract and enhance PMP skills data
3. Analyze charity project requirements
4. Perform optimal matching with LinkedIn considerations
5. Generate comprehensive reports

Input Files Required (automatically detected):
- Any Excel file in input/ containing "PMDoS", "Registration", "Responses" (most recent will be used)
- Any Excel file in input/ containing "Charities", "Information", "Responses" (most recent will be used)

Output Files Generated:
- LinkedIn_Analysis_Report.xlsx
- PMI_PMP_Charity_Matching_Results_Enhanced.xlsx
- Matching_Summary.csv (for quick reference)
- analysis_log.txt (processing log)
"""

import pandas as pd
import os
import sys
from datetime import datetime


def _safe_console_print(text: str):
    """Print to console in a way that won't crash on Windows code pages.

    If the active console encoding (e.g. cp1252) can't represent certain
    Unicode characters (✓, 🎉, ✅, etc.), we fall back to an ASCII-safe
    version by stripping or replacing those characters. This prevents
    UnicodeEncodeError that previously aborted the analysis pipeline.
    """
    try:
        print(text)
    except UnicodeEncodeError:
        # Replace known symbols with ASCII approximations and drop the rest
        replacements = {
            '✓': '[OK]',
            '✅': '[OK]',
            '🎉': '*',
            '➤': '>',
            '–': '-',
            '—': '-',
        }
        sanitized = ''.join(replacements.get(ch, ch) for ch in text)
        try:
            # Final defensive encode/decode cycle
            sanitized.encode(sys.stdout.encoding or 'ascii', errors='ignore')
            print(sanitized)
        except Exception:
            # Last resort: ASCII only
            print(sanitized.encode('ascii', errors='ignore').decode('ascii'))


def log_message(message, log_file="Output/analysis_log.txt"):
    """Log messages to both console and UTF-8 file without crashing on Unicode."""
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    log_entry = f"[{timestamp}] {message}"
    _safe_console_print(log_entry)

    # Ensure Output directory exists
    os.makedirs("Output", exist_ok=True)

    with open(log_file, "a", encoding="utf-8") as f:
        f.write(log_entry + "\n")


def _has_change_flag(value):
    """Return True when the Changes column indicates an update."""
    if pd.isna(value):
        return False
    if isinstance(value, str):
        return bool(value.strip())
    return True


def _filter_pmp_changes(pmp_df):
    """Drop PMP registrants flagged in the Changes column."""
    matching_cols = [
        col for col in pmp_df.columns
        if 'changes' in str(col).strip().lower()
    ]

    if not matching_cols:
        return pmp_df, 0, None

    changes_col = matching_cols[0]
    change_mask = pmp_df[changes_col].apply(_has_change_flag)
    filtered_df = pmp_df.loc[~change_mask].copy()
    removed_count = int(change_mask.sum())

    return filtered_df, removed_count, changes_col


def validate_input_files():
    """Validate that required input files exist (dynamic file detection)"""
    from dynamic_file_loader import get_latest_input_files
    
    reg_file, charity_file = get_latest_input_files()
    
    if not reg_file:
        log_message("ERROR: No PMDoS registration file found in input/")
        log_message(
            "Looking for files containing: "
            "PMDoS, Registration, Responses"
        )
        return False
        
    if not charity_file:
        log_message("ERROR: No charity information file found in input/")
        log_message(
            "Looking for files containing: "
            "Charities, Information, Responses"
        )
        return False
    
    log_message(f"✓ Found registration file: {os.path.basename(reg_file)}")
    log_message(f"✓ Found charity file: {os.path.basename(charity_file)}")
    
    return True


def run_linkedin_analysis():
    """Step 1: Run LinkedIn profile analysis"""
    log_message("Step 1: Running LinkedIn Profile Analysis...")
    
    try:
        # Import and run the LinkedIn analysis functions
        from linkedin_enhanced_matching import (
            validate_linkedin_urls,
            generate_linkedin_analysis_report,
            enhanced_extract_pmp_skills
        )
        
        # Load PMP data (dynamic file detection)
        from dynamic_file_loader import get_latest_input_files
        pmp_file, _ = get_latest_input_files()
        if not pmp_file:
            raise Exception("Could not find PMP registration file")
        
        log_message(f"Using registration file: {os.path.basename(pmp_file)}")
        pmp_df = pd.read_excel(pmp_file)

        pmp_df, removed_count, changes_col = _filter_pmp_changes(pmp_df)
        if removed_count:
            log_message(
                "  Excluding "
                f"{removed_count} registrant(s) flagged in '{changes_col}'"
            )
            log_message("  Changes column indicates unavailability; skipping")
        
        # Validate LinkedIn URLs
        validation_df = validate_linkedin_urls(pmp_df)
        valid_linkedin_count = validation_df['Is_Valid'].sum()
        total_pmp = len(pmp_df)
        coverage_pct = (
            valid_linkedin_count / total_pmp * 100
            if total_pmp else 0.0
        )
        
        # Generate enhanced profiles
        enhanced_profiles = enhanced_extract_pmp_skills(pmp_df)
        
        # Generate LinkedIn analysis report
        linkedin_report = generate_linkedin_analysis_report(enhanced_profiles)
        avg_quality = (
            linkedin_report['LinkedIn_Quality_Score'].mean()
            if not linkedin_report.empty else 0
        )
        avg_completeness = (
            linkedin_report['Profile_Completeness_Score'].mean()
            if not linkedin_report.empty else 0
        )
        
        # Save LinkedIn analysis results
        output_file = "Output/LinkedIn_Analysis_Report.xlsx"
        with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
            validation_df.to_excel(
                writer,
                sheet_name='URL_Validation',
                index=False
            )
            linkedin_report.to_excel(
                writer,
                sheet_name='LinkedIn_Analysis',
                index=False
            )
            
            # Summary statistics
            summary_rows = [
                ['Total PMPs', total_pmp],
                ['Valid LinkedIn URLs', valid_linkedin_count],
                ['LinkedIn Coverage %', f"{coverage_pct:.1f}%"],
                ['Avg LinkedIn Quality Score', f"{avg_quality:.1f}"],
                ['Avg Profile Completeness', f"{avg_completeness:.1f}"],
            ]
            summary_stats = pd.DataFrame(
                summary_rows,
                columns=['Metric', 'Value']
            )
            summary_stats.to_excel(writer, sheet_name='Summary', index=False)
        
        log_message(
            f"✓ LinkedIn analysis saved to: {output_file}"
        )
        log_message(
            "  LinkedIn Coverage: "
            f"{valid_linkedin_count}/{total_pmp} ({coverage_pct:.1f}%)"
        )
        log_message(
            f"  Average LinkedIn Quality: {avg_quality:.1f}/10"
        )
        
        return True
        
    except Exception as e:
        log_message(f"ERROR in LinkedIn analysis: {str(e)}")
        return False


def run_enhanced_matching(use_flexible_assignment=False):
    """Step 2: Run enhanced PMP-Charity matching"""
    if use_flexible_assignment:
        log_message("Step 2: Running Flexible PMP Assignment...")
        log_message("  Mode: All PMPs assigned to projects")
    else:
        log_message("Step 2: Running Standard PMP-Charity Matching...")
        log_message("  Mode: Two PMPs per charity")
    
    try:
        if use_flexible_assignment:
            # Import flexible assignment functions
            from flexible_pmp_assignment import (
                create_flexible_matching,
                generate_flexible_matching_report
            )
            from enhanced_pmp_charity_matching import (
                load_and_process_data,
                extract_pmp_skills,
                analyze_charity_requirements
            )
        else:
            # Import the standard enhanced matching functions
            from enhanced_pmp_charity_matching import (
                load_and_process_data,
                extract_pmp_skills,
                analyze_charity_requirements,
                create_optimal_matching,
                generate_matching_report,
                create_detailed_analysis
            )
        
        # Load and process data
        pmp_df, charity_df = load_and_process_data()
        
        pmp_df, removed_count, changes_col = _filter_pmp_changes(pmp_df)
        if removed_count:
            log_message(
                "  Excluding "
                f"{removed_count} registrant(s) flagged in '{changes_col}'"
            )
            log_message("  Changes column indicates unavailability; skipping")

        # Extract enhanced PMP skills
        pmp_profiles = extract_pmp_skills(pmp_df)
        
        # Analyze charity requirements
        charity_projects = analyze_charity_requirements(charity_df)
        
        # Create matching based on assignment type
        if use_flexible_assignment:
            final_matches, assigned_charities = create_flexible_matching(
                pmp_profiles,
                charity_projects
            )
            matching_summary = generate_flexible_matching_report(
                final_matches,
                assigned_charities
            )
            output_file = (
                'Output/PMI_PMP_Charity_Flexible_Matching_Results.xlsx'
            )
            sheet_name = 'Flexible_Matching'
            log_message("  Using flexible assignment (all PMPs assigned)")
        else:
            final_matches, assigned_charities = create_optimal_matching(
                pmp_profiles,
                charity_projects
            )
            matching_summary = generate_matching_report(
                final_matches,
                assigned_charities
            )
            detailed_analysis = create_detailed_analysis(
                pmp_profiles,
                charity_projects,
                final_matches
            )
            output_file = (
                'Output/PMI_PMP_Charity_Matching_Results_Enhanced.xlsx'
            )
            sheet_name = 'Enhanced_Matching_Summary'
            log_message("  Using standard assignment (2 PMPs per charity)")
        
        # Save enhanced matching results
        with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
            
            # Enhanced summary sheet
            matching_summary.to_excel(
                writer,
                sheet_name=sheet_name,
                index=False
            )
            
            # Detailed analysis (only for standard matching)
            if not use_flexible_assignment:
                detailed_analysis.to_excel(
                    writer,
                    sheet_name='Detailed_Analysis_Enhanced',
                    index=False
                )
            
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
            linkedin_analysis.to_excel(
                writer,
                sheet_name='LinkedIn_Analysis',
                index=False
            )
            
            # Enhanced PMP profiles
            pmp_summary = pd.DataFrame([{
                'ID': p['ID'],
                'Name': p['Name'],
                'Experience': p['Experience'],
                'LinkedIn_Quality': p['LinkedIn_Quality_Score'],
                'Profile_Completeness': p['Profile_Completeness_Score'],
                'Enhanced_Overall_Score': round(p['Overall_Score'], 2),
                'Areas_of_Interest': p['Areas_of_Interest']
            } for p in pmp_profiles])
            pmp_summary.to_excel(
                writer,
                sheet_name='Enhanced_PMP_Profiles',
                index=False
            )
            
            # Charity projects
            charity_summary = pd.DataFrame([{
                'ID': c['ID'],
                'Organization': c['Organization'],
                'Initiative': c['Initiative'],
                'Priority': c['Priority_Level'],
                'Complexity': c['Complexity']
            } for c in charity_projects])
            charity_summary.to_excel(
                writer,
                sheet_name='Charity_Projects',
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
        
        # Also save a quick CSV summary for easy viewing
        matching_summary_csv = matching_summary[[
            'Charity_Organization',
            'Charity_Initiative',
            'PMP_Name',
            'Match_Score',
            'LinkedIn_Quality',
            'PMP_Experience'
        ]].copy()
        matching_summary_csv.to_csv('Output/Matching_Summary.csv', index=False)
        
        log_message(
            f"✓ Enhanced matching saved to: {output_file}"
        )
        log_message(
            "✓ Quick reference CSV saved to: Output/Matching_Summary.csv"
        )
        log_message(f"  Total PMPs: {len(pmp_profiles)}")
        log_message(f"  Total Charity Projects: {len(charity_projects)}")
        log_message(f"  Total Matches Created: {len(final_matches)}")
        
        # Log summary of matches
        log_message("  Match Summary:")
        for charity_id, matches in assigned_charities.items():
            charity_name = matches[0]['Charity_Project']['Organization']
            pmp_names = [match['PMP_Name'] for match in matches]
            scores = [round(match['Score'], 2) for match in matches]
            log_message(f"    {charity_name}: {pmp_names} (Scores: {scores})")
        
        return True
        
    except Exception as e:
        log_message(f"ERROR in enhanced matching: {str(e)}")
        return False


def cleanup_old_outputs():
    """Clean up old output files before running new analysis"""
    # Ensure Output directory exists
    os.makedirs("Output", exist_ok=True)
    
    output_files = [
        "Output/LinkedIn_Analysis_Report.xlsx",
        "Output/PMI_PMP_Charity_Matching_Results_Enhanced.xlsx",
        "Output/Matching_Summary.csv",
        "Output/analysis_log.txt"
    ]
    
    cleaned = []
    for file_path in output_files:
        if os.path.exists(file_path):
            try:
                os.remove(file_path)
                cleaned.append(file_path)
            except Exception as e:
                log_message(f"Warning: Could not remove {file_path}: {str(e)}")
    
    if cleaned:
        log_message(f"Cleaned up old output files: {', '.join(cleaned)}")


def generate_summary_report():
    """Generate a final summary report"""
    log_message("Step 3: Generating Final Summary Report...")
    
    try:
        # Read the generated files to create a summary
        if os.path.exists("Output/LinkedIn_Analysis_Report.xlsx"):
            linkedin_df = pd.read_excel(
                "Output/LinkedIn_Analysis_Report.xlsx",
                sheet_name="Summary"
            )
            
        if os.path.exists(
            "Output/PMI_PMP_Charity_Matching_Results_Enhanced.xlsx"
        ):
            matching_df = pd.read_excel(
                "Output/PMI_PMP_Charity_Matching_Results_Enhanced.xlsx",
                sheet_name="Enhanced_Matching_Summary"
            )
        
        # Create a summary text file
        with open("Output/Analysis_Summary.txt", "w", encoding="utf-8") as f:
            f.write("PMP-CHARITY MATCHING ANALYSIS SUMMARY\n")
            f.write("=" * 50 + "\n")
            analysis_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            f.write(f"Analysis completed on: {analysis_time}\n\n")
            
            if 'linkedin_df' in locals():
                f.write("LINKEDIN ANALYSIS RESULTS:\n")
                f.write("-" * 30 + "\n")
                for _, row in linkedin_df.iterrows():
                    f.write(f"{row['Metric']}: {row['Value']}\n")
                f.write("\n")
            
            if 'matching_df' in locals():
                f.write("MATCHING RESULTS SUMMARY:\n")
                f.write("-" * 30 + "\n")
                total_matches = len(matching_df)
                charity_count = matching_df['Charity_Organization'].nunique()
                pmp_count = matching_df['PMP_Name'].nunique()
                avg_match_score = matching_df['Match_Score'].mean()
                avg_linkedin_quality = matching_df['LinkedIn_Quality'].mean()
                f.write(f"Total matches created: {total_matches}\n")
                f.write(f"Charities matched: {charity_count}\n")
                f.write(f"PMPs assigned: {pmp_count}\n")
                f.write(f"Average match score: {avg_match_score:.2f}\n")
                f.write(
                    "Average LinkedIn quality: "
                    f"{avg_linkedin_quality:.1f}/10\n\n"
                )
                
                f.write("TOP 5 MATCHES BY SCORE:\n")
                f.write("-" * 25 + "\n")
                top_matches = matching_df.nlargest(5, 'Match_Score')
                for _, match in top_matches.iterrows():
                    charity_name = match['Charity_Organization']
                    pmp_name = match['PMP_Name']
                    score = match['Match_Score']
                    f.write(
                        f"{charity_name} ← {pmp_name} (Score: {score:.2f})\n"
                    )
        
        log_message("✓ Summary report generated: Output/Analysis_Summary.txt")
        return True
        
    except Exception as e:
        log_message(f"Warning: Could not generate summary report: {str(e)}")
        return False


def main():
    """
    Main function - This is the ONLY function you need to run!
    
    When your Excel input files change, just run:
        python run_complete_analysis.py
        
    Or for flexible assignment (all PMPs to projects):
        python run_complete_analysis.py --flexible
    """
    
    import sys
    
    # Check for flexible assignment flag
    use_flexible = '--flexible' in sys.argv or '-f' in sys.argv
    
    _safe_console_print("=" * 70)
    if use_flexible:
        _safe_console_print(
            "PMP-CHARITY FLEXIBLE MATCHING COMPLETE ANALYSIS PIPELINE"
        )
        _safe_console_print(
            "(All PMPs assigned to projects - some projects may get 3+ PMPs)"
        )
    else:
        _safe_console_print(
            "PMP-CHARITY STANDARD MATCHING COMPLETE ANALYSIS PIPELINE"
        )
        _safe_console_print(
            "(Standard 2 PMPs per charity assignment)"
        )
    _safe_console_print("=" * 70)
    _safe_console_print(
        "This script runs the complete analysis when input files change."
    )
    _safe_console_print("Input files expected in 'input/' directory.")
    _safe_console_print("=" * 70)
    
    # Clean up old files
    cleanup_old_outputs()
    
    # Initialize log
    if use_flexible:
        log_message("Starting flexible PMP-Charity analysis pipeline...")
    else:
        log_message("Starting standard PMP-Charity analysis pipeline...")
    
    # Step 0: Validate input files
    log_message("Step 0: Validating input files...")
    if not validate_input_files():
        log_message("ANALYSIS ABORTED: Missing input files")
        return False
    
    # Step 1: LinkedIn Analysis
    if not run_linkedin_analysis():
        log_message("ANALYSIS ABORTED: LinkedIn analysis failed")
        return False
    
    # Step 2: Enhanced Matching (with assignment type choice)
    if not run_enhanced_matching(use_flexible_assignment=use_flexible):
        log_message("ANALYSIS ABORTED: Enhanced matching failed")
        return False
    
    # Step 3: Generate Summary
    generate_summary_report()
    
    # Final success message
    log_message("=" * 50)
    log_message("🎉 COMPLETE ANALYSIS FINISHED SUCCESSFULLY!")
    log_message("=" * 50)
    log_message("Output files generated:")
    log_message("  1. Output/LinkedIn_Analysis_Report.xlsx - LinkedIn profile analysis")
    log_message("  2. Output/PMI_PMP_Charity_Matching_Results_Enhanced.xlsx - Complete matching results")
    log_message("  3. Output/Matching_Summary.csv - Quick reference CSV")
    log_message("  4. Output/Analysis_Summary.txt - Executive summary")
    log_message("  5. analysis_log.txt - Detailed processing log")
    log_message("=" * 50)
    
    _safe_console_print("\n" + "=" * 70)
    _safe_console_print("ANALYSIS COMPLETE!")
    _safe_console_print("Check the generated Excel files for detailed results.")
    _safe_console_print("Check Output/Analysis_Summary.txt for a quick overview.")
    _safe_console_print("=" * 70)
    
    return True


if __name__ == "__main__":
    main()