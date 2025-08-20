#!/usr/bin/env python3
"""
Simple demo script for AI PDF to Excel extraction.
This script demonstrates the basic functionality of the extraction tool.
"""

import os
import sys
from pathlib import Path

# Add the current directory to the path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from pdf_to_excel_ai import PDFToExcelAI

def demo():
    """Run a simple demonstration of the PDF to Excel extraction."""
    
    print("🚀 AI PDF to Excel Extraction Demo")
    print("=" * 40)
    
    # Check if we have any PDFs to process
    input_dir = Path("input_pdfs")
    pdf_files = list(input_dir.glob("*.pdf"))
    
    if not pdf_files:
        print("❌ No PDF files found for demo.")
        print(f"Please place some PDF files in the '{input_dir}' directory.")
        return False
    
    print(f"📁 Found {len(pdf_files)} PDF file(s) to process:")
    for pdf in pdf_files:
        print(f"   • {pdf.name}")
    
    # Initialize the converter
    converter = PDFToExcelAI()
    
    # Process the PDFs
    output_files = converter.process_all_pdfs()
    
    if output_files:
        print(f"\n✅ Successfully created {len(output_files)} Excel file(s):")
        for excel_file in output_files:
            print(f"   📊 {Path(excel_file).name}")
            
        # Show a sample of the first Excel file
        print(f"\n📄 Sample content from {Path(output_files[0]).name}:")
        try:
            import pandas as pd
            df = pd.read_excel(output_files[0], sheet_name=0)  # First sheet
            print(df.head().to_string())
        except Exception as e:
            print(f"   Could not display content: {e}")
        
        return True
    else:
        print("\n❌ No files were successfully processed.")
        return False

if __name__ == "__main__":
    success = demo()
    sys.exit(0 if success else 1)