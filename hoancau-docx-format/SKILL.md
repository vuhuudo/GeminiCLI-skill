---
name: hoancau-docx-format
description: "Skill này hỗ trợ căn chỉnh văn bản Word (.docx) tuân thủ nghiêm ngặt Quy định Thể thức và Kỹ thuật trình bày của HoanCauGroup (phiên bản cập nhật 2026), dựa trên tiêu chuẩn từ bộ phận Pháp chế và mẫu chuẩn 'Tờ trình v3.0'."
---

# HoanCauGroup DOCX Format Skill (Standard 2026)

Skill này hỗ trợ căn chỉnh văn bản Word (.docx) tuân thủ nghiêm ngặt Quy định Thể thức và Kỹ thuật trình bày của HoanCauGroup (phiên bản cập nhật 2026), dựa trên tiêu chuẩn từ bộ phận Pháp chế và mẫu chuẩn "Tờ trình v3.0".

## Quy định chính (Standard 2026)

- **Khổ giấy:** A4.
- **Lề (Margins):** Trên 2.5cm (để Header thông thoáng), Dưới 2.0cm, Trái 3.0cm, Phải 1.5cm.
- **Font:** Times New Roman.
- **Cỡ chữ nội dung:** 13pt.
- **Giãn dòng:** 1.3x. Khoảng cách đoạn: Before 3pt, After 6pt.
- **Thụt đầu dòng:** 1.25cm.
- **Đánh số trang:** Giữa, dưới cùng (footer), 12pt.

## Cách sử dụng

### 1. Căn chỉnh nhanh toàn bộ văn bản (Optimized v2026)
Sử dụng script đã được tối ưu hóa (One-pass optimization) để thiết lập toàn bộ khung chuẩn và nội dung (bao gồm cả Bảng biểu) chỉ trong một lần chạy:

```bash
/home/johndoe/.gemini/skills/hoancau-docx-format/scripts/format_hoancau.sh <file.docx>
```

**Cải tiến mới:**
- **Tốc độ:** Giảm số lần gọi tool, xử lý nhanh gấp 10 lần so với bản cũ.
- **Bảng biểu:** Tự động căn chỉnh font Times New Roman 12pt và giãn dòng **1.3x**. Phân loại nội dung: Chữ căn trái, Số/Tiền tệ căn giữa (ngoại trừ bảng Header luôn căn giữa).
- **Tiêu đề:** Tự động in đậm và **gạch chân** cho các mục số La Mã (I, II, III...).
- **Danh sách (Numbering/Bullet):** Thụt lề phân cấp thông minh dựa trên ngữ cảnh dấu ":", đảm bảo nội dung chữ luôn thẳng hàng tại **1.25cm** (hoặc sâu hơn theo cấp).

### 2. Định dạng các thành phần đặc biệt (Theo mẫu Tờ trình v3.0)

#### Tiêu đề & Trích yếu (Ô số 6 & 7)
- **TỜ TRÌNH:** In hoa, 16pt, đậm, căn giữa. Khoảng cách: **SpaceBefore 18pt, SpaceAfter 12pt**.
- **V/v:** 13pt, đậm, nghiêng, căn giữa, ngay dưới tiêu đề. Khoảng cách: **SpaceAfter 12pt**.

```bash
officecli set doc.docx "/body/p[N]" --prop alignment=center --prop bold=true --prop size=16pt --prop text="TỜ TRÌNH" --prop spaceBefore=18pt --prop spaceAfter=12pt
officecli set doc.docx "/body/p[N+1]" --prop alignment=center --prop bold=true --prop italic=true --prop size=13pt --prop find="V/v:" --prop spaceAfter=12pt
```

#### Kính gửi
- **Kính gửi:** 13pt, đậm, đứng.
```bash
officecli set doc.docx "/body/p[M]" --prop bold=true --prop size=13pt --prop find="Kính gửi:"
```

#### Đầu trang (Header Table)
Sử dụng bảng 3 cột ẩn viền để trình bày Tên đơn vị và Quốc hiệu. Tham khảo `references/hoancau_standards.md`.
- **Lưu ý Spacing:** Để tránh văn bản bị "xẹp", toàn bộ các đoạn văn trong Header Table phải đặt **lineSpacing=1.15** và **spaceBefore/After=2pt**.

## Quy trình làm việc đề xuất (Smart Control 2026 v3.0)

1. **Phân tích (Xác định ParaId):** Sử dụng `officecli view <file> text` để xác định cấu trúc các đoạn văn bản (đặc biệt là Header cũ và Tiêu đề).
2. **Setup khung & Header:** Chạy `format_hoancau.sh`. 
   - **LƯU Ý ĐƠN VỊ:** Luôn dùng hậu tố `x` cho giãn dòng (VD: `1.3x`) và `pt` cho khoảng cách đoạn (VD: `12pt`) để tránh Word hiểu nhầm sang twips gây lỗi đè chữ.
3. **Làm sạch dữ liệu (Deduplication):**
   - Cập nhật Số hiệu và Ngày tháng thực tế vào `tbl[1]` bằng `officecli set`.
   - **Xóa triệt để:** Dùng `officecli remove` để xóa các dòng text Header cũ (Tên đơn vị, Số hiệu, Ngày tháng) và các đoạn trống dư thừa.
4. **Căn chỉnh chi tiết (Safety Spacing Strategy):**
   - **TỜ TRÌNH:** `size=16pt`, `bold=true`, `alignment=center`, `spaceBefore=18pt`, `spaceAfter=12pt`, **`lineSpacing=1.3x`**.
   - **V/v:** `size=13pt`, `bold=true`, `italic=true`, `alignment=center`, **`spaceAfter=12pt`**, **`lineSpacing=1.3x`**.
   - **Kính gửi:** `bold=true`, `italic=false`, **`spaceAfter=12pt`**.
   - **Mục lục (1., 2., 3.):** `bold=true`, `firstLineIndent=0`.
   - **Nội dung:** `alignment=justify`, `lineSpacing=1.3x`, `spaceBefore=3pt`, `spaceAfter=6pt`, `firstLineIndent=1.25cm`.
5. **Kiểm tra tự động (Double-Check):**
   - Chạy `officecli view <file> stats` để xác nhận Font (Times New Roman) và Size đã chuẩn.
   - Chạy `officecli view <file> issues --type format` để quét lỗi thụt đầu dòng, khoảng trắng thừa hoặc lỗi spacing.
6. **Xác nhận thủ công:** So sánh trực quan lần cuối trước khi bàn giao.

## Tài liệu tham khảo
- [references/hoancau_standards.md](references/hoancau_standards.md): Chi tiết thông số 2026.
- [Pháp Chế - QUY ĐỊNH THỂ THỨC TRÌNH BÀY VĂN BẢN - Trình ký- 26.09.25.docx]: Tài liệu gốc.
