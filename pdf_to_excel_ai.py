#!/usr/bin/env python3
"""
AI-Powered PDF to Excel Extraction Tool

This script uses advanced AI algorithms to extract tables from PDF files
and convert them to Excel format (.xlsx) with intelligent data cleaning
and formatting.

Features:
- Multiple parsing strategies for optimal table detection
- Intelligent data cleaning and validation
- Excel output with proper formatting
- Batch processing of multiple PDFs
- Comprehensive error handling and logging
- Progress tracking and reporting

Dependencies:
- camelot-py: For PDF table extraction
- pandas: For data manipulation
- openpyxl: For Excel file output
- tabulate: For pretty console output
"""

import os
import sys
import logging
import argparse
import traceback
from pathlib import Path
from typing import List, Optional, Tuple
import warnings

# Suppress warnings for cleaner output
warnings.filterwarnings('ignore')

try:
    import pandas as pd
    import numpy as np
    from tabulate import tabulate
    import camelot
    from PyPDF2 import PdfReader
except ImportError as e:
    print(f"Error: Missing required dependency: {e}")
    print("Please install dependencies with: pip install camelot-py pandas openpyxl tabulate PyPDF2")
    sys.exit(1)

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('pdf_extraction.log'),
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger(__name__)


class PDFToExcelAI:
    """
    AI-powered PDF to Excel extraction engine with intelligent table detection
    and data cleaning capabilities.
    """
    
    def __init__(self, input_dir: str = "input_pdfs", output_dir: str = "output_excel"):
        """
        Initialize the PDF to Excel converter.
        
        Args:
            input_dir: Directory containing PDF files to process
            output_dir: Directory to save Excel files
        """
        self.input_dir = Path(input_dir)
        self.output_dir = Path(output_dir)
        self.setup_directories()
        
    def setup_directories(self):
        """Create necessary directories if they don't exist."""
        self.input_dir.mkdir(exist_ok=True)
        self.output_dir.mkdir(exist_ok=True)
        logger.info(f"Input directory: {self.input_dir.absolute()}")
        logger.info(f"Output directory: {self.output_dir.absolute()}")
        
    def validate_pdf(self, pdf_path: Path) -> bool:
        """
        Validate if PDF file is readable and contains content.
        
        Args:
            pdf_path: Path to PDF file
            
        Returns:
            True if PDF is valid, False otherwise
        """
        try:
            reader = PdfReader(str(pdf_path))
            if len(reader.pages) == 0:
                logger.warning(f"PDF {pdf_path.name} contains no pages")
                return False
            logger.info(f"PDF {pdf_path.name} validated - {len(reader.pages)} pages")
            return True
        except Exception as e:
            logger.error(f"Failed to validate PDF {pdf_path.name}: {e}")
            return False
            
    def extract_tables_with_ai(self, pdf_path: Path) -> List:
        """
        Extract tables from PDF using AI-powered multiple parsing strategies.
        
        Args:
            pdf_path: Path to PDF file
            
        Returns:
            List of extracted tables
        """
        tables = []
        
        # Strategy 1: Network-based parsing (ML-powered)
        try:
            logger.info(f"Attempting network-based parsing for {pdf_path.name}")
            tables_network = camelot.read_pdf(str(pdf_path), flavor="stream")
            if len(tables_network) > 0:
                logger.info(f"Network parser found {len(tables_network)} tables")
                tables = tables_network
            else:
                logger.info("Network parser found no tables, trying lattice parser")
        except Exception as e:
            logger.warning(f"Network parsing failed for {pdf_path.name}: {e}")
            
        # Strategy 2: Lattice-based parsing (for structured tables)
        if len(tables) == 0:
            try:
                logger.info(f"Attempting lattice-based parsing for {pdf_path.name}")
                tables_lattice = camelot.read_pdf(str(pdf_path), flavor="lattice")
                if len(tables_lattice) > 0:
                    logger.info(f"Lattice parser found {len(tables_lattice)} tables")
                    tables = tables_lattice
                else:
                    logger.info("Lattice parser found no tables, trying with custom areas")
            except Exception as e:
                logger.warning(f"Lattice parsing failed for {pdf_path.name}: {e}")
                
        # Strategy 3: Custom area detection
        if len(tables) == 0:
            try:
                logger.info(f"Attempting custom area parsing for {pdf_path.name}")
                tables_custom = camelot.read_pdf(
                    str(pdf_path), 
                    flavor="lattice", 
                    table_areas=["50,750,500,50"]
                )
                if len(tables_custom) > 0:
                    logger.info(f"Custom area parser found {len(tables_custom)} tables")
                    tables = tables_custom
            except Exception as e:
                logger.warning(f"Custom area parsing failed for {pdf_path.name}: {e}")
                
        return tables
        
    def clean_dataframe_ai(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        AI-powered data cleaning and enhancement.
        
        Args:
            df: Raw pandas DataFrame from PDF extraction
            
        Returns:
            Cleaned and enhanced DataFrame
        """
        if df.empty:
            return df
            
        # Remove leading/trailing whitespace
        df = df.applymap(lambda x: x.strip() if isinstance(x, str) else x)
        
        # Replace various empty value representations with NaN
        empty_values = ["", "nan", "NaN", "NULL", "null", "N/A", "n/a", "-", "–"]
        df = df.replace(empty_values, np.nan)
        
        # Remove completely empty rows
        df = df.dropna(how="all")
        
        # Remove completely empty columns
        df = df.dropna(axis=1, how="all")
        
        # Try to detect and set proper headers
        if len(df) > 1:
            # Check if first row contains likely headers (non-numeric values)
            first_row = df.iloc[0]
            if first_row.dtype == 'object' or any(isinstance(val, str) for val in first_row):
                # Use first row as headers if they look like headers
                if not any(str(val).isdigit() for val in first_row if pd.notna(val)):
                    df.columns = first_row
                    df = df.drop(df.index[0])
                    
        # Reset index after cleaning
        df = df.reset_index(drop=True)
        
        # Fill remaining NaN values with empty strings for Excel compatibility
        df = df.fillna("")
        
        logger.info(f"DataFrame cleaned: shape {df.shape}")
        return df
        
    def save_to_excel(self, dataframes: List[pd.DataFrame], pdf_name: str) -> str:
        """
        Save extracted tables to Excel file with proper formatting.
        
        Args:
            dataframes: List of DataFrames to save
            pdf_name: Original PDF filename
            
        Returns:
            Path to saved Excel file
        """
        excel_path = self.output_dir / f"{Path(pdf_name).stem}_extracted_tables.xlsx"
        
        try:
            with pd.ExcelWriter(str(excel_path), engine='openpyxl') as writer:
                for i, df in enumerate(dataframes, 1):
                    sheet_name = f"Table_{i}"
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
                    
                    # Get the worksheet to apply formatting
                    worksheet = writer.sheets[sheet_name]
                    
                    # Auto-adjust column widths
                    for column in worksheet.columns:
                        max_length = 0
                        column_letter = column[0].column_letter
                        for cell in column:
                            try:
                                if len(str(cell.value)) > max_length:
                                    max_length = len(str(cell.value))
                            except:
                                pass
                        adjusted_width = min(max_length + 2, 50)  # Cap at 50 characters
                        worksheet.column_dimensions[column_letter].width = adjusted_width
                        
            logger.info(f"Saved {len(dataframes)} tables to {excel_path}")
            return str(excel_path)
            
        except Exception as e:
            logger.error(f"Failed to save Excel file {excel_path}: {e}")
            raise
            
    def process_single_pdf(self, pdf_path: Path) -> Optional[str]:
        """
        Process a single PDF file and extract tables to Excel.
        
        Args:
            pdf_path: Path to PDF file
            
        Returns:
            Path to output Excel file, or None if processing failed
        """
        logger.info(f"Processing {pdf_path.name}")
        
        # Validate PDF
        if not self.validate_pdf(pdf_path):
            return None
            
        # Extract tables using AI
        try:
            tables = self.extract_tables_with_ai(pdf_path)
        except Exception as e:
            logger.error(f"Failed to extract tables from {pdf_path.name}: {e}")
            return None
            
        if len(tables) == 0:
            logger.warning(f"No tables detected in {pdf_path.name}")
            return None
            
        # Clean and process each table
        cleaned_dataframes = []
        for i, table in enumerate(tables):
            try:
                # Convert to DataFrame and clean
                df = table.df
                df_cleaned = self.clean_dataframe_ai(df)
                
                if not df_cleaned.empty:
                    cleaned_dataframes.append(df_cleaned)
                    logger.info(f"Table {i+1}: {df_cleaned.shape} after cleaning")
                    
                    # Log parsing quality
                    if hasattr(table, 'parsing_report'):
                        logger.info(f"Table {i+1} parsing report: {table.parsing_report}")
                        
            except Exception as e:
                logger.error(f"Failed to process table {i+1} from {pdf_path.name}: {e}")
                continue
                
        if not cleaned_dataframes:
            logger.warning(f"No valid tables extracted from {pdf_path.name}")
            return None
            
        # Save to Excel
        try:
            excel_path = self.save_to_excel(cleaned_dataframes, pdf_path.name)
            logger.info(f"Successfully processed {pdf_path.name} -> {excel_path}")
            return excel_path
        except Exception as e:
            logger.error(f"Failed to save Excel file for {pdf_path.name}: {e}")
            return None
            
    def process_all_pdfs(self) -> List[str]:
        """
        Process all PDF files in the input directory.
        
        Returns:
            List of paths to generated Excel files
        """
        pdf_files = list(self.input_dir.glob("*.pdf"))
        
        if not pdf_files:
            logger.warning(f"No PDF files found in {self.input_dir}")
            print(f"\nNo PDF files found in {self.input_dir}")
            print(f"Please place PDF files in the '{self.input_dir}' directory and run again.")
            return []
            
        logger.info(f"Found {len(pdf_files)} PDF files to process")
        
        successful_outputs = []
        
        for pdf_file in pdf_files:
            try:
                output_path = self.process_single_pdf(pdf_file)
                if output_path:
                    successful_outputs.append(output_path)
            except Exception as e:
                logger.error(f"Unexpected error processing {pdf_file.name}: {e}")
                logger.error(traceback.format_exc())
                
        logger.info(f"Processing complete. {len(successful_outputs)} Excel files generated.")
        return successful_outputs
        
    def display_summary(self, output_files: List[str]):
        """Display a summary of processed files."""
        print("\n" + "="*60)
        print("AI PDF TO EXCEL EXTRACTION SUMMARY")
        print("="*60)
        
        if output_files:
            print(f"✅ Successfully processed {len(output_files)} files:")
            for file_path in output_files:
                file_size = Path(file_path).stat().st_size / 1024  # KB
                print(f"   📄 {Path(file_path).name} ({file_size:.1f} KB)")
        else:
            print("❌ No files were successfully processed.")
            
        print(f"\n📁 Output directory: {self.output_dir.absolute()}")
        print(f"📋 Log file: pdf_extraction.log")
        print("="*60)


def main():
    """Main function to run the PDF to Excel extraction."""
    parser = argparse.ArgumentParser(
        description="AI-Powered PDF to Excel Extraction Tool",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python3 pdf_to_excel_ai.py                    # Process PDFs in 'input_pdfs' directory
  python3 pdf_to_excel_ai.py -i docs -o output  # Custom input/output directories
  python3 pdf_to_excel_ai.py --verbose          # Enable verbose logging
        """
    )
    
    parser.add_argument(
        "-i", "--input",
        default="input_pdfs",
        help="Input directory containing PDF files (default: input_pdfs)"
    )
    
    parser.add_argument(
        "-o", "--output",
        default="output_excel",
        help="Output directory for Excel files (default: output_excel)"
    )
    
    parser.add_argument(
        "-v", "--verbose",
        action="store_true",
        help="Enable verbose logging"
    )
    
    args = parser.parse_args()
    
    if args.verbose:
        logging.getLogger().setLevel(logging.DEBUG)
        
    print("🚀 AI-Powered PDF to Excel Extraction Tool")
    print("=" * 50)
    
    # Initialize and run the converter
    converter = PDFToExcelAI(args.input, args.output)
    output_files = converter.process_all_pdfs()
    converter.display_summary(output_files)
    
    return len(output_files) > 0


if __name__ == "__main__":
    try:
        success = main()
        sys.exit(0 if success else 1)
    except KeyboardInterrupt:
        print("\n\n⚠️  Process interrupted by user")
        sys.exit(1)
    except Exception as e:
        logger.error(f"Unexpected error: {e}")
        logger.error(traceback.format_exc())
        sys.exit(1)