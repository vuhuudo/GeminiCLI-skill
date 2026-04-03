import os
import sys
import re
from docx import Document
from docx.shared import Pt, Mm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT

def format_docx(input_path, output_path):
    if not os.path.exists(input_path):
        print(f"Error: File {input_path} does not exist.")
        return

    doc = Document(input_path)
    
    # --- SMART ANALYSIS ---
    total_chars = sum(len(p.text) for p in doc.paragraphs)
    total_paragraphs = len(doc.paragraphs)
    
    # Determine if we need to be aggressive to save space
    is_long = total_chars > 3200 or total_paragraphs > 40
    
    font_size = Pt(13) if is_long else Pt(14)
    line_spacing = 1.0 if is_long else 1.15
    space_after = Pt(2) if is_long else Pt(6)

    # 1. Set Margins (Nghị định 30)
    for section in doc.sections:
        section.top_margin = Mm(20)
        section.bottom_margin = Mm(20)
        section.left_margin = Mm(30)
        section.right_margin = Mm(15) # Minimal right margin to prevent wrapping

    # 2. Set Default Styles
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.size = font_size
    
    paragraph_format = style.paragraph_format
    paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    paragraph_format.line_spacing = line_spacing
    paragraph_format.space_after = space_after
    paragraph_format.first_line_indent = Mm(10)

    # 3. Process Paragraphs
    paragraphs = list(doc.paragraphs)
    for i, paragraph in enumerate(paragraphs):
        text = paragraph.text.strip()
        
        # Remove empty paragraphs at the end of document
        if not text and i > len(paragraphs) * 0.9:
            p_element = paragraph._element
            if p_element.getparent() is not None:
                p_element.getparent().remove(p_element)
            continue

        if not text:
            continue

        # Reset to Normal style
        paragraph.style = doc.styles['Normal']
        for run in paragraph.runs:
            run.font.name = 'Times New Roman'
            run.font.size = font_size

        # --- SPECIAL LOGIC ---

        # Quốc hiệu (Fixed 12pt, No Indent to prevent wrap)
        if "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM" in text.upper():
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            paragraph.paragraph_format.first_line_indent = 0
            paragraph.paragraph_format.space_after = 0
            for run in paragraph.runs:
                run.bold = True
                run.font.size = Pt(12) # Slightly smaller to fit one line

        # Tiêu ngữ
        elif "ĐỘC LẬP - TỰ DO - HẠNH PHÚC" in text.upper() or "ĐỘC LẬP-TỰ DO-HẠNH PHÚC" in text.upper():
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            paragraph.paragraph_format.first_line_indent = 0
            paragraph.paragraph_format.space_after = Pt(10)
            for run in paragraph.runs:
                run.bold = True
                run.font.size = Pt(13)

        # Đề mục (I., II., 1., 2. ...) - No Indent, Bold
        elif re.match(r'^([IVXLCDM]+\.|[0-9]+\.|[a-z]\))\s+', text):
            paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
            paragraph.paragraph_format.first_line_indent = 0
            paragraph.paragraph_format.space_before = Pt(10)
            for run in paragraph.runs:
                run.bold = True
                run.font.size = font_size + Pt(0.5)

        # Địa danh, ngày tháng
        elif re.search(r",\s*ngày\s+\d+\s+tháng\s+\d+\s+năm\s+\d+", text, re.IGNORECASE):
            paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            paragraph.paragraph_format.first_line_indent = 0
            for run in paragraph.runs:
                run.italic = True

        # Tiêu đề văn bản
        elif text.isupper() and len(text) < 100:
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            paragraph.paragraph_format.first_line_indent = 0
            paragraph.paragraph_format.space_before = Pt(12)
            paragraph.paragraph_format.space_after = Pt(12)
            for run in paragraph.runs:
                run.bold = True
                run.font.size = font_size + Pt(1)

        # Chữ ký (Signature block detection)
        signature_keywords = ["người lập", "giám đốc", "phê duyệt", "trưởng phòng", "kế toán"]
        if any(k in text.lower() for k in signature_keywords) and len(text) < 40:
            # If it's near the end, make it compact to fit same page
            paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            paragraph.paragraph_format.first_line_indent = 0
            paragraph.paragraph_format.space_before = Pt(10) if is_long else Pt(20)
            for run in paragraph.runs:
                run.bold = True

    # 4. Process Tables
    for table in doc.tables:
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    paragraph.paragraph_format.first_line_indent = 0
                    paragraph.paragraph_format.space_after = 0
                    # Check if it's a signature table (usually at the end)
                    if any(k in paragraph.text.lower() for k in ["giám đốc", "người lập"]):
                        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    for run in paragraph.runs:
                        run.font.name = 'Times New Roman'
                        run.font.size = Pt(12)

    doc.save(output_path)
    print(f"Professional formatted: {output_path} (Long mode: {is_long})")

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python format_docx_vn.py <input_path> <output_path>")
    else:
        format_docx(sys.argv[1], sys.argv[2])
