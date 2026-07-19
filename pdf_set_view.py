# pdf_set_view.py - make PDF open to first page + show bookmarks panel
# pip install pypdf
import os
import sys
from pypdf import PdfReader, PdfWriter
from pypdf.generic import NameObject, ArrayObject

def set_pdf_to_open_with_bookmarks(pdf_path):
    pdf_path = os.path.abspath(pdf_path)
    if not os.path.exists(pdf_path):
        raise FileNotFoundError(pdf_path)

    reader = PdfReader(pdf_path)
    writer = PdfWriter()
    writer.clone_document_from_reader(reader) # keeps Word bookmarks

    # Show bookmarks sidebar - works on all pypdf versions
    writer._root_object[NameObject("/PageMode")] = NameObject("/UseOutlines")

    # Open to first page
    first_page_ref = writer.pages[0].indirect_reference
    writer._root_object[NameObject("/OpenAction")] = ArrayObject([first_page_ref, NameObject("/Fit")])

    tmp = pdf_path + ".tmp.pdf"
    with open(tmp, "wb") as f:
        writer.write(f)
    os.replace(tmp, pdf_path)
    print(f"Patched: {pdf_path} -> will open on page 1 with bookmarks")

if __name__ == "__main__":
    if len(sys.argv)!= 2:
        print('Usage: python pdf_set_view.py "C:\\path\\to\\file.pdf"')
    else:
        set_pdf_to_open_with_bookmarks(sys.argv[1])