#!/usr/bin/env python3
"""
Script: Word to PDF Batch Converter

Purpose:
Traverses the current directory and all its subdirectories, finds all Microsoft Word files (.doc and .docx), and converts each to a PDF document with the same name in the same directory.

Prerequisites and Dependencies:
- Python 3 must be installed.
- Install Python library 'docx2pdf' by running the command:
    pip install docx2pdf
- LibreOffice and 'unoconv' are required to convert older '.doc' files:
    - On Ubuntu/Linux, install by running:
        sudo apt install libreoffice unoconv
    - For Windows, LibreOffice and unoconv must be installed and accessible via the system PATH.

Usage:
Simply execute the script from your desired root folder. PDF files will appear alongside the original Word documents.

"""

# Import necessary Python libraries/modules
import os  # Allows access and navigation of the file system
import subprocess  # Used to run external programs (LibreOffice via unoconv)
from docx2pdf import convert  # Library to convert .docx files to PDF easily

def convert_to_pdf(input_file, output_file):
    """
    Converts a Word document to PDF format.

    Parameters:
    - input_file: Path to the original Word file (.doc or .docx).
    - output_file: Desired path of the resulting PDF file.
    """
    # Extract the file extension to decide conversion method
    ext = os.path.splitext(input_file)[1].lower()

    # If the document is a modern Word document (.docx)
    if ext == ".docx":
        # Convert using docx2pdf library
        convert(input_file, output_file)

    # If the document is an older Word format (.doc)
    elif ext == ".doc":
        # Convert using the external program unoconv (LibreOffice must be installed)
        subprocess.run(['unoconv', '-f', 'pdf', '-o', output_file, input_file], check=True)

def traverse_and_convert(root_dir):
    """
    Traverses directories starting from 'root_dir', converting all Word documents found.

    Parameters:
    - root_dir: The root directory from which the traversal begins.
    """
    # Walk through all directories and subdirectories
    for dirpath, _, filenames in os.walk(root_dir):
        for file in filenames:
            # Check if the file ends with .doc or .docx (case-insensitive)
            if file.lower().endswith((".doc", ".docx")):
                # Create full path to the original Word document
                doc_path = os.path.join(dirpath, file)
                # Define the PDF file path (same name, same directory, .pdf extension)
                pdf_path = os.path.splitext(doc_path)[0] + ".pdf"
                try:
                    # Attempt to convert the Word document to PDF
                    convert_to_pdf(doc_path, pdf_path)
                    # Inform the user upon successful conversion
                    print(f"Converted: {doc_path} -> {pdf_path}")
                except Exception as e:
                    # Inform the user if conversion fails, and provide the error details
                    print(f"Failed to convert: {doc_path}. Reason: {e}")

# Main entry-point of the script
if __name__ == "__main__":
    # Set the root directory to the current working directory (directory where script is executed)
    root_directory = os.getcwd()
    # Begin traversing directories and converting documents
    traverse_and_convert(root_directory)
