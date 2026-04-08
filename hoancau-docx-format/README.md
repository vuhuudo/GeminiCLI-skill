# HoanCauGroup DOCX Format Skill

Skill này hỗ trợ căn chỉnh văn bản Word (.docx) tuân thủ nghiêm ngặt Quy định Thể thức và Kỹ thuật trình bày của HoanCauGroup (phiên bản 2025).

## Quy định chính (HoanCauGroup)

- **Khổ giấy:** A4.
- **Lề (Margins):** 20mm-25mm (Trên/Dưới), 30mm-35mm (Trái), 15mm-20mm (Phải).
- **Font:** Times New Roman.
- **Cỡ chữ nội dung:** 12pt - 13pt.
- **Giãn dòng:** 1.3 - 1.5 lines. Khoảng cách đoạn: 3pt - 6pt.
- **Đánh số trang:** Giữa, dưới cùng (footer), 12pt.

## Cách sử dụng

### 1. Căn chỉnh nhanh toàn bộ văn bản
Sử dụng script để thiết lập khổ giấy, lề và định dạng chung cho toàn bộ paragraph:

```bash
/home/johndoe/.gemini/skills/hoancau-docx-format/scripts/format_hoancau.sh <file.docx>
```

### 2. Định dạng các thành phần đặc biệt bằng officecli
Dưới đây là các lệnh `officecli` mẫu để tuân thủ quy định:

#### Thiết lập Quốc hiệu & Tiêu ngữ (Ô số 1)
Quốc hiệu (In hoa, 12pt, đậm) và Tiêu ngữ (13pt, đậm, đứng) thường nằm trong một bảng 2 cột ở đầu trang hoặc căn lề phải.
```bash
# Ví dụ căn lề paragraph
officecli set doc.docx "/body/p[1]" --prop alignment=center --prop bold=true --prop size=12pt --prop text="CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM"
officecli set doc.docx "/body/p[2]" --prop alignment=center --prop bold=true --prop size=13pt --prop text="Độc lập - Tự do - Hạnh phúc"
```

#### Thiết lập Tên loại & Trích yếu (Ô số 6)
```bash
officecli set doc.docx "/body/p[N]" --prop alignment=center --prop bold=true --prop size=16pt --prop text="QUYẾT ĐỊNH"
officecli set doc.docx "/body/p[N+1]" --prop alignment=center --prop bold=true --prop italic=true --prop size=13pt --prop text="Về việc ban hành quy định thể thức văn bản"
```

#### Thiết lập Nơi nhận (Ô số 9)
```bash
officecli set doc.docx "/body/p[M]" --prop size=10pt --prop bold=true --prop text="Nơi nhận:"
officecli set doc.docx "/body/p[M+1]" --prop size=10pt --prop text="- Như Điều 3;"
```

### 3. Quy trình làm việc đề xuất
1. **Phân tích:** Xác định loại văn bản (Quy định, Quyết định, Tờ trình...).
2. **Setup chung:** Chạy `format_hoancau.sh` để thiết lập khung chuẩn.
3. **Căn chỉnh chi tiết:** Sử dụng `officecli` hoặc `batch` để định dạng chính xác từng ô theo [references/hoancau_standards.md](references/hoancau_standards.md).
4. **Kiểm tra:** Đảm bảo footer có số trang (12pt), font toàn bộ là Times New Roman.

## Tài liệu tham khảo
- [references/hoancau_standards.md](references/hoancau_standards.md): Chi tiết các thông số kỹ thuật (2025).
- Phụ lục 01: Bảng chữ viết tắt tên loại văn bản.
