import pandas as pd
import numpy as np
from pathlib import Path

# Read the Excel files
def read_excel_files():
    print("Reading Excel files...")
    
    # Read PMP professionals data
    pmp_file = r"c:\PMI\input\2025 - PMI Sydney Chapter Project Management Day of Service (PMDoS) 2025 Professional Registration (Responses).xlsx"
    pmp_df = pd.read_excel(pmp_file)
    
    # Read charity projects data
    charity_file = r"c:\PMI\input\Charities Project Information 2025 (Responses).xlsx"
    charity_df = pd.read_excel(charity_file)
    
    return pmp_df, charity_df

# Analyze the structure of both datasets
def analyze_data_structure(pmp_df, charity_df):
    print("=== PMP Professionals Data Structure ===")
    print(f"Shape: {pmp_df.shape}")
    print(f"Columns: {list(pmp_df.columns)}")
    print("\nFirst few rows:")
    print(pmp_df.head())
    
    print("\n" + "="*60)
    print("=== Charity Projects Data Structure ===")
    print(f"Shape: {charity_df.shape}")
    print(f"Columns: {list(charity_df.columns)}")
    print("\nFirst few rows:")
    print(charity_df.head())
    
    return pmp_df, charity_df

if __name__ == "__main__":
    try:
        pmp_df, charity_df = read_excel_files()
        analyze_data_structure(pmp_df, charity_df)
        
        # Save the dataframes info to text files for easier review
        with open("pmp_data_info.txt", "w", encoding='utf-8') as f:
            f.write("PMP Professionals Data Info\n")
            f.write("="*40 + "\n")
            f.write(f"Shape: {pmp_df.shape}\n\n")
            f.write("Columns:\n")
            for i, col in enumerate(pmp_df.columns):
                f.write(f"{i+1}. {col}\n")
            f.write("\nSample data:\n")
            f.write(pmp_df.to_string())
        
        with open("charity_data_info.txt", "w", encoding='utf-8') as f:
            f.write("Charity Projects Data Info\n")
            f.write("="*40 + "\n")
            f.write(f"Shape: {charity_df.shape}\n\n")
            f.write("Columns:\n")
            for i, col in enumerate(charity_df.columns):
                f.write(f"{i+1}. {col}\n")
            f.write("\nSample data:\n")
            f.write(charity_df.to_string())
            
    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()
