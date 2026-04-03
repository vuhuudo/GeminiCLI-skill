import sys
import os

try:
    from pdf2image import convert_from_path
    import pytesseract
    from docx import Document
except ImportError as e:
    print(f"Error: Missing library {e.name}. Please install it.")
    sys.exit(1)

def ocr_pdf_to_docx(input_pdf, output_docx):
    if not os.path.exists(input_pdf):
        print(f"Error: File {input_pdf} not found.")
        return

    # Check for tesseract
    try:
        pytesseract.get_tesseract_version()
    except pytesseract.TesseractNotFoundError:
        print("Error: Tesseract OCR is not installed or not in PATH.")
        print("Please install Tesseract OCR (e.g., 'sudo apt install tesseract-ocr' and 'tesseract-ocr-vie' for Vietnamese).")
        return

    print(f"Starting OCR for {input_pdf}...")
    try:
        pages = convert_from_path(input_pdf)
        doc = Document()
        
        for i, page in enumerate(pages):
            print(f"Processing page {i+1}...")
            # Perform OCR (using Vietnamese and English)
            text = pytesseract.image_to_string(page, lang='vie+eng')
            doc.add_paragraph(text)
        
        doc.save(output_docx)
        print(f"OCR completed. Saved to {output_docx}")
    except Exception as e:
        print(f"An error occurred during OCR: {str(e)}")

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python ocr_pdf_to_docx.py <input_pdf> <output_docx>")
    else:
        ocr_pdf_to_docx(sys.argv[1], sys.argv[2])
