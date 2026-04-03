---
name: vn-office-pro
description: Chuyên viên văn phòng AI xử lý văn bản chuẩn Việt Nam, OCR PDF sang Word và phân tích báo cáo Excel. Sử dụng khi cần căn chỉnh tài liệu theo Nghị định 30/2020/NĐ-CP, trích xuất văn bản từ PDF scan, hoặc phân tích dữ liệu Excel thành báo cáo.
---

# VN Office Pro - Chuyên viên Văn phòng AI

Kỹ năng này giúp Gemini CLI thực hiện các nghiệp vụ văn phòng chuyên sâu tại Việt Nam một cách tự động và chuẩn xác.

## Các tính năng chính

1. **Căn chỉnh văn bản chuẩn Việt Nam**: Tự động định dạng file Word (.docx) theo Nghị định 30/2020/NĐ-CP (Font Times New Roman, cỡ chữ 13-14, lề chuẩn, căn lề đều hai bên).
2. **OCR PDF sang Word**: Chuyển đổi file PDF scan thành văn bản Word có thể chỉnh sửa được (Yêu cầu Tesseract OCR).
3. **Phân tích Excel**: Đọc, lọc, phân tích và tổng hợp dữ liệu từ các tệp Excel thành báo cáo tóm tắt.

## Hướng dẫn sử dụng

### 1. Căn chỉnh file Word chuẩn Việt Nam
Khi người dùng yêu cầu "căn chỉnh chuẩn VN" hoặc "định dạng theo Nghị định 30", hãy sử dụng:
```bash
./.venv/bin/python vn-office-pro/scripts/format_docx_vn.py <input.docx> <output.docx>
```

### 2. Phân tích Excel
Khi người dùng yêu cầu "phân tích file excel" hoặc "tổng hợp báo cáo", hãy sử dụng:
```bash
./.venv/bin/python vn-office-pro/scripts/analyze_excel.py <data.xlsx> <report.txt>
```

### 3. OCR PDF sang Word (Thử nghiệm)
*Lưu ý: Tính năng này yêu cầu hệ thống có sẵn tesseract-ocr và poppler-utils.*
Nếu khả dụng, hãy sử dụng:
```bash
./.venv/bin/python vn-office-pro/scripts/ocr_pdf_to_docx.py <scanned.pdf> <output.docx>
```

## Tiêu chuẩn Việt Nam (Nghị định 30/2020/NĐ-CP)
- **Font**: Times New Roman, bảng mã Unicode.
- **Cỡ chữ**: 13pt - 14pt.
- **Lề văn bản**: 
  - Trên: 20mm - 25mm.
  - Dưới: 20mm - 25mm.
  - Trái: 30mm - 35mm.
  - Phải: 15mm - 20mm.
- **Căn lề**: Đều hai bên (Justified).
- **Giãn dòng**: 1.0 - 1.5 lines.
