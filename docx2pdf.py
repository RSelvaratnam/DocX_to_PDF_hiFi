
# Full-fidelity DOCX to PDF converter using Word automation (minimal parameters).
# This script uses Microsoft Word via COM automation to convert a .docx file to PDF,
# preserving all formatting, tables, images, and generating bookmarks from heading styles.
# Requirements: Microsoft Word installed, and pywin32 library (pip install pywin32).
# Note: This is Windows-specific due to COM usage.
# fixed for OneDrive paths and spaces in filenames

import win32com.client as win32  # Library for COM automation to control Word
import os  # Standard library for file path operations and directory creation
import sys
import pythoncom

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
    if not os.path.exists(input_path):
        raise FileNotFoundError(f"Input file not found: {input_path}")

    input_abs = os.path.abspath(input_path)
    output_abs = os.path.abspath(output_path)
    os.makedirs(os.path.dirname(output_abs) or ".", exist_ok=True)

    pythoncom.CoInitialize()
    word = win32.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0 # wdAlertsNone

    doc = None
    try:
        # Open read-only, no conversion prompts, no recent files update
        doc = word.Documents.Open(
            FileName=input_abs,
            ConfirmConversions=False,
            ReadOnly=True,
            AddToRecentFiles=False,
            Visible=False,
            NoEncodingDialog=True
        )

        # First try ExportAsFixedFormat - best for bookmarks
        try:
            doc.ExportAsFixedFormat(
                OutputFileName=output_abs,
                ExportFormat=17, # wdExportFormatPDF
                OpenAfterExport=False, # <-- THIS was causing your error
                OptimizeFor=0, # wdExportOptimizeForPrint
                Range=0, # wdExportAllDocument
                IncludeDocProps=True,
                KeepIRM=True,
                CreateBookmarks=1, # wdExportCreateHeadingBookmarks
                DocStructureTags=True,
                BitmapMissingFonts=True,
                UseISO19005_1=False
            )
        except Exception:
            # Fallback that almost never fails
            doc.SaveAs2(output_abs, FileFormat=17) # 17 = wdFormatPDF

        print(f"Conversion complete: {output_abs}")

    except Exception as e:
        print(f"Error during conversion: {e}")
        raise
    finally:
        if doc is not None:
            try:
                doc.Close(SaveChanges=False)
            except:
                pass
        try:
            word.Quit()
        except:
            pass
        pythoncom.CoUninitialize()

if __name__ == "__main__":
    if len(sys.argv)!= 3:
        print('Usage: python docx2pdf.py "input.docx" "output.pdf"')
        sys.exit(1)
    convert_docx_to_pdf(sys.argv[1], sys.argv[2])
