#!/usr/bin/env python3
"""
Script: Enhanced Word-to-PDF Batch Converter

Purpose:
Traverses the current directory and all subdirectories, identifies Microsoft Word files (.doc and .docx),
and converts each into a PDF document with an identical name in the same directory. Provides real-time
status updates, including total conversions needed, progress percentage, current file status, and
iterations per minute.

Prerequisites and Dependencies:
- Python 3.x installed.
- Microsoft Office (Word) required for '.docx' files (Windows users):
    - Fully installed, activated, and licensed.
- Install the Python library 'docx2pdf':
    pip install docx2pdf
- LibreOffice and 'unoconv' for older '.doc' files or if Word isn't available:
    - Ubuntu/Linux:
        sudo apt install libreoffice unoconv
    - Windows:
        - LibreOffice: https://www.libreoffice.org/download/download/
        - unoconv: https://github.com/unoconv/unoconv
        - Ensure LibreOffice and unoconv are in the PATH.

Usage:
Run script from the desired root directory. PDF files appear alongside original documents.
"""

import os
import subprocess
import time
from docx2pdf import convert

# Conversion function
def convert_to_pdf(input_file, output_file):
    ext = os.path.splitext(input_file)[1].lower()
    if ext == ".docx":
        convert(input_file, output_file)
    elif ext == ".doc":
        subprocess.run(['unoconv', '-f', 'pdf', '-o', output_file, input_file], check=True)

# Find all Word documents to convert
def find_word_documents(root_dir):
    word_files = []
    for dirpath, _, filenames in os.walk(root_dir):
        for file in filenames:
            if file.lower().endswith((".doc", ".docx")):
                word_files.append(os.path.join(dirpath, file))
    return word_files

# Main conversion and progress tracking function
def traverse_and_convert(root_dir):
    word_files = find_word_documents(root_dir)
    total_files = len(word_files)
    start_time = time.time()

    for idx, doc_path in enumerate(word_files, 1):
        pdf_path = os.path.splitext(doc_path)[0] + ".pdf"
        try:
            file_start_time = time.time()
            convert_to_pdf(doc_path, pdf_path)
            file_elapsed = time.time() - file_start_time

            total_elapsed = time.time() - start_time
            iterations_per_minute = (idx / total_elapsed) * 60 if total_elapsed > 0 else 0
            percent_complete = (idx / total_files) * 100

            print(f"[{idx}/{total_files}] {percent_complete:.2f}% | "
                  f"Current file: '{os.path.basename(doc_path)}' converted in {file_elapsed:.2f}s | "
                  f"Rate: {iterations_per_minute:.2f} files/min")

        except Exception as e:
            print(f"Conversion failed for {doc_path}. Reason: {e}")

# Entry point
if __name__ == "__main__":
    root_directory = os.getcwd()
    traverse_and_convert(root_directory)
