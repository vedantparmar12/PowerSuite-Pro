#!/usr/bin/env python3
"""
PowerPoint Generation Script - Main entry point for creating presentations
Usage: python generate_presentation.py "prompt" [output_file.pptx]
"""

import sys
import os
from pathlib import Path

# Add the current directory to path to import ppt_creator
sys.path.insert(0, str(Path(__file__).parent))

from ppt_creator import ProfessionalPPTCreator

def main():
    if len(sys.argv) < 2:
        print("Usage: python generate_presentation.py \"prompt\" [output_file.pptx]")
        print("Example: python generate_presentation.py \"Create a quarterly business review\" quarterly_review.pptx")
        sys.exit(1)
    
    prompt = sys.argv[1]
    output_file = sys.argv[2] if len(sys.argv) > 2 else None
    
    try:
        # Create presentation generator
        creator = ProfessionalPPTCreator()
        
        # Generate presentation
        output_path = creator.generate_presentation(prompt, output_file)
        
        print(f"Presentation created successfully: {output_path}")
        return output_path
        
    except Exception as e:
        print(f"Error creating presentation: {str(e)}")
        sys.exit(1)

if __name__ == "__main__":
    main()
