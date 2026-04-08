---
name: hoancau-docx-format
description: "Skill này hỗ trợ căn chỉnh văn bản Word (.docx) tuân thủ nghiêm ngặt Quy định Thể thức và Kỹ thuật trình bày của HoanCauGroup (phiên bản cập nhật 2026), dựa trên tiêu chuẩn từ bộ phận Pháp chế và mẫu chuẩn 'Tờ trình v3.0'."
---

# HoanCauGroup DOCX Format Skill (Standard 2026)

Skill này hỗ trợ căn chỉnh văn bản Word (.docx) tuân thủ nghiêm ngặt Quy định Thể thức và Kỹ thuật trình bày của HoanCauGroup (phiên bản cập nhật 2026), dựa trên tiêu chuẩn từ bộ phận Pháp chế và mẫu chuẩn "Tờ trình v3.0".

## Quy định chính (Standard 2026)

- **Khổ giấy:** A4.
- **Lề (Margins):** Trên 2.0cm, Dưới 2.0cm, Trái 3.0cm, Phải 1.5cm.
- **Font:** Times New Roman.
- **Cỡ chữ nội dung:** 13pt.
- **Giãn dòng:** 1.3x. Khoảng cách đoạn: Before 3pt, After 6pt.
- **Thụt đầu dòng:** 1.25cm.
- **Đánh số trang:** Giữa, dưới cùng (footer), 12pt.

## Cách sử dụng

### 1. Căn chỉnh nhanh toàn bộ văn bản
Sử dụng script để thiết lập khung chuẩn (Lề, Font, Giãn dòng, Header/Footer cơ bản):

```bash
/home/johndoe/.gemini/skills/hoancau-docx-format/scripts/format_hoancau.sh <file.docx>
```

### 2. Định dạng các thành phần đặc biệt (Theo mẫu Tờ trình v3.0)

#### Tiêu đề & Trích yếu (Ô số 6 & 7)
- **TỜ TRÌNH:** In hoa, 16pt, đậm, căn giữa.
- **V/v:** 13pt, đậm, nghiêng, căn giữa, ngay dưới tiêu đề.

```bash
officecli set doc.docx "/body/p[N]" --prop alignment=center --prop bold=true --prop size=16pt --prop text="TỜ TRÌNH"
officecli set doc.docx "/body/p[N+1]" --prop alignment=center --prop bold=true --prop italic=true --prop size=13pt --prop find="V/v:"
```

#### Kính gửi
- **Kính gửi:** 13pt, đậm, đứng.
```bash
officecli set doc.docx "/body/p[M]" --prop bold=true --prop size=13pt --prop find="Kính gửi:"
```

#### Đầu trang (Header Table)
Sử dụng bảng 3 cột ẩn viền để trình bày Tên đơn vị và Quốc hiệu. Tham khảo `references/hoancau_standards.md`.

## Quy trình làm việc đề xuất
1. **Phân tích:** Xác định cấu trúc văn bản.
2. **Setup khung:** Chạy `format_hoancau.sh`.
3. **Căn chỉnh chi tiết:** 
   - Đảm bảo các mục 1., 2. , 3. là 13pt Bold.
   - Đảm bảo bảng biểu có font Times New Roman, cỡ chữ 11-13pt.
4. **Kiểm tra:** So sánh với `Tờ_trình_mua_PC_v3.0.docx` về mặt trực quan.

## Tài liệu tham khảo
- [references/hoancau_standards.md](references/hoancau_standards.md): Chi tiết thông số 2026.
- [Pháp Chế - QUY ĐỊNH THỂ THỨC TRÌNH BÀY VĂN BẢN - Trình ký- 26.09.25.docx]: Tài liệu gốc.
