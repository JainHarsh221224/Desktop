#!/bin/bash

echo "🔧 Setting up AI PDF to Excel Extraction Tool..."

# Install Python dependencies
echo "📦 Installing Python dependencies..."
pip3 install --user camelot-py[cv] pandas openpyxl tabulate PyPDF2 numpy

# Create directories
echo "📁 Creating directories..."
mkdir -p input_pdfs
mkdir -p output_excel

echo "✅ Setup complete!"
echo ""
echo "📋 To use the tool:"
echo "   1. Place PDF files in the 'input_pdfs' directory"
echo "   2. Run: python3 pdf_to_excel_ai.py"
echo "   3. Check the 'output_excel' directory for results"
echo ""
echo "ℹ️  For help: python3 pdf_to_excel_ai.py --help"