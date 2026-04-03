import sys
import os
import re

try:
    from pdf2image import convert_from_path
    import pytesseract
    from docx import Document
except ImportError as e:
    print(f"Error: Missing library '{e.name}'. Please install it.")
    sys.exit(1)

OCR_DPI = 300
OCR_LANG = 'vie+eng'


def _clean_ocr_text(text):
    """Remove excessive blank lines produced by OCR."""
    text = re.sub(r'\n{3,}', '\n\n', text)
    return text.strip()


def ocr_pdf_to_docx(input_pdf, output_docx):
    if not os.path.exists(input_pdf):
        print(f"Error: File '{input_pdf}' not found.")
        sys.exit(1)

    # Check for tesseract
    try:
        pytesseract.get_tesseract_version()
    except pytesseract.TesseractNotFoundError:
        print("Error: Tesseract OCR is not installed or not in PATH.")
        print("Please install: sudo apt install tesseract-ocr tesseract-ocr-vie")
        sys.exit(1)

    print(f"Starting OCR for '{input_pdf}' (DPI={OCR_DPI})...")
    try:
        pages = convert_from_path(input_pdf, dpi=OCR_DPI)
    except Exception as e:
        print(f"Error: Could not convert PDF to images: {e}")
        sys.exit(1)

    doc = Document()
    total = len(pages)

    for i, page in enumerate(pages, start=1):
        print(f"  Processing page {i}/{total}...")
        text = pytesseract.image_to_string(page, lang=OCR_LANG)
        text = _clean_ocr_text(text)

        paragraph = doc.add_paragraph(text)

        # Add a page break after every page except the last
        if i < total:
            doc.add_page_break()

    try:
        doc.save(output_docx)
    except OSError as e:
        print(f"Error: Could not save '{output_docx}': {e}")
        sys.exit(1)

    print(f"OCR completed. Saved to '{output_docx}' ({total} page(s)).")


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python ocr_pdf_to_docx.py <input_pdf> <output_docx>")
        sys.exit(1)
    ocr_pdf_to_docx(sys.argv[1], sys.argv[2])
