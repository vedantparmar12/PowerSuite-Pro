#!/usr/bin/env python3
"""
Excel Generation Script - Main entry point for creating spreadsheets
Usage: python generate_spreadsheet.py "prompt" [output_file.xlsx]
"""

import sys
import os
from pathlib import Path

# Add the current directory to path to import excel_master
sys.path.insert(0, str(Path(__file__).parent))

from excel_master import ExcelMaster

def main():
    if len(sys.argv) < 2:
        print("Usage: python generate_spreadsheet.py \"prompt\" [output_file.xlsx]")
        print("Example: python generate_spreadsheet.py \"Create a budget tracker\" budget_tracker.xlsx")
        sys.exit(1)
    
    prompt = sys.argv[1]
    output_file = sys.argv[2] if len(sys.argv) > 2 else None
    
    try:
        # Create Excel generator
        excel_master = ExcelMaster()
        
        # Generate spreadsheet
        output_path = excel_master.create_spreadsheet(prompt, output_file)
        
        print(f" Spreadsheet created successfully: {output_path}")
        return output_path
        
    except Exception as e:
        print(f"Error creating spreadsheet: {str(e)}")
        sys.exit(1)

if __name__ == "__main__":
    main()
