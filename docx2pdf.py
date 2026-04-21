
# Full-fidelity DOCX to PDF converter using Word automation (minimal parameters).
# This script uses Microsoft Word via COM automation to convert a .docx file to PDF,
# preserving all formatting, tables, images, and generating bookmarks from heading styles.
# Requirements: Microsoft Word installed, and pywin32 library (pip install pywin32).
# Note: This is Windows-specific due to COM usage.

import win32com.client as win32  # Library for COM automation to control Word
import os  # Standard library for file path operations and directory creation

def convert_docx_to_pdf(input_path, output_path):
    """
    Convert a .docx file to PDF using Microsoft Word.
    
    This function automates Word to open the input document, export it as PDF with
    bookmarks enabled (based on heading styles), and clean up resources afterward.
    
    Args:
        input_path (str): Path to the input .docx file.
        output_path (str): Path where the output PDF will be saved.
    
    Raises:
        FileNotFoundError: If the input .docx file does not exist.
        Exception: For any errors during the conversion process (e.g., Word issues).
    """
    # Check if the input file exists to avoid runtime errors later
    if not os.path.exists(input_path):
        raise FileNotFoundError(f"Input file not found: {input_path}")
    
    # Launch Microsoft Word application via COM (invisibly in the background)
    word = win32.Dispatch("Word.Application")  # Creates a COM object for Word
    word.Visible = False  # Hide Word's UI to run silently
    word.DisplayAlerts = False  # Suppress any dialog boxes or alerts from Word
    
    try:
        # Open the .docx document in Word
        doc = word.Documents.Open(os.path.abspath(input_path))  # Use absolute path for reliability
        
        # Ensure the output directory exists (create if necessary)
        output_dir = os.path.dirname(output_path) or "."  # Default to current dir if no path
        os.makedirs(output_dir, exist_ok=True)  # Creates dirs recursively if needed
        
        # Export the document as PDF with key settings
        doc.ExportAsFixedFormat(
            OutputFileName=os.path.abspath(output_path),  # Full path for the output PDF
            ExportFormat=17,  # Constant for PDF format (wdExportFormatPDF)
            OpenAfterExport=True,  # Do not open the PDF after saving
            OptimizeFor=0,  # Optimize for print quality (wdExportOptimizeForPrint)
            Range=0,  # Export the entire document (wdExportAllDocument)
            From=0,  # Start from the beginning (no specific page range)
            To=0,  # End at the last page
            Item=0,  # Export document content only (wdExportDocumentContent)
            IncludeDocProps=True,  # Include document properties (e.g., author, title) in PDF
            CreateBookmarks=1  # Enable bookmarks/outlines from Word heading styles
        )
        
        # Print success messages for user feedback
        print(f"Conversion complete: {output_path}")
        print("Bookmarks generated from heading styles—check PDF sidebar.")
        
    except Exception as e:
        # Catch and report any errors during the process (e.g., COM failures or file issues)
        print(f"Error during conversion: {e}")
    
    finally:
        # Always clean up resources to avoid leaving Word processes running
        if 'doc' in locals():  # Check if doc was successfully created
            doc.Close(SaveChanges=False)  # Close the document without saving changes
        word.Quit()  # Shut down the Word application
        word = None  # Release the COM object reference to free memory

