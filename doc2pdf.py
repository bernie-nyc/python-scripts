#!/usr/bin/env python3
"""
Script: Word to PDF Batch Converter

Purpose:
Traverses the current directory and all its subdirectories, finds Microsoft Word files (.doc and .docx), and converts each into a PDF document with an identical name in the same directory.

Prerequisites and Dependencies:
- Python 3.x installed.
- Microsoft Office (Word) installation required (Windows users):
    - docx2pdf uses Word's COM interface for conversion.
    - Office must be fully installed, activated, and licensed.
- Install the Python library 'docx2pdf':
    pip install docx2pdf
- To convert older '.doc' files or if Microsoft Word isn't available:
    - LibreOffice and 'unoconv' must be installed.
        - Ubuntu/Linux:
            sudo apt install libreoffice unoconv
        - Windows:
            - Install LibreOffice: https://www.libreoffice.org/download/download/
            - Install unoconv: https://github.com/unoconv/unoconv
            - Add LibreOffice and unoconv to the PATH environment variable.

Usage:
Run this script from the desired root directory. Converted PDFs will appear in the source document directories.

"""

# Import Python libraries
import os  # Enables interaction with the file system
import subprocess  # Runs external commands (LibreOffice/unoconv)
from docx2pdf import convert  # Converts .docx using Microsoft Word

def convert_to_pdf(input_file, output_file):
    """
    Converts Word documents (.docx or .doc) to PDF format.

    Parameters:
    - input_file: Path to the Word document.
    - output_file: Path for saving the converted PDF.
    """
    ext = os.path.splitext(input_file)[1].lower()

    # Conversion method depends on file extension
    if ext == ".docx":
        # Uses docx2pdf (requires Microsoft Word on Windows)
        convert(input_file, output_file)
    elif ext == ".doc":
        # Uses LibreOffice/unoconv for older .doc files
        subprocess.run(['unoconv', '-f', 'pdf', '-o', output_file, input_file], check=True)

def traverse_and_convert(root_dir):
    """
    Navigates directories starting from 'root_dir', converting Word files to PDFs.

    Parameters:
    - root_dir: Root directory from which traversal begins.
    """
    # Walk through directories and subdirectories
    for dirpath, _, filenames in os.walk(root_dir):
        for file in filenames:
            # Check for Word documents (.doc and .docx)
            if file.lower().endswith((".doc", ".docx")):
                # Paths for input Word and output PDF files
                doc_path = os.path.join(dirpath, file)
                pdf_path = os.path.splitext(doc_path)[0] + ".pdf"
                try:
                    # Attempt conversion and notify success
                    convert_to_pdf(doc_path, pdf_path)
                    print(f"Converted: {doc_path} -> {pdf_path}")
                except Exception as e:
                    # Notify if conversion fails and display error
                    print(f"Failed to convert: {doc_path}. Reason: {e}")

if __name__ == "__main__":
    # Set current directory as starting point
    root_directory = os.getcwd()
    # Begin the conversion process
    traverse_and_convert(root_directory)
