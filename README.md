# chuan-hoa-tai-lieu

Bộ công cụ kiểm tra và chuẩn hoá thể thức tài liệu Word (`.docx`) theo **Nghị định 30/2020/NĐ-CP** về công tác văn
thư và các quy ước thường dùng cho báo cáo khoa học, luận văn ở Việt Nam. Dùng được như một skill cho Claude Code
(xem [SKILL.md](SKILL.md)) hoặc chạy trực tiếp bằng dòng lệnh.

## Tính năng

- **Kiểm tra** (`scripts/audit_docx.py`): khổ giấy, lề, phông, cỡ chữ; tiêu đề chưa gán style; số chỉ mục trùng hoặc
  nhảy cóc; chú thích bảng/biểu đồ/hình trùng số, thiếu số, sai dấu; mục lục gõ tay; thành phần thể thức hành chính
  (quốc hiệu, tiêu ngữ, cơ quan, số ký hiệu, ngày tháng, nơi nhận, chữ ký); ảnh, bảng, số trang, dấu thập phân.
- **Sửa tự động** (`scripts/fix_docx.py`): trang A4 và lề chuẩn, Times New Roman, định dạng đoạn văn; gán style
  tiêu đề 4 cấp và định dạng đậm/nghiêng theo cấp; tạo mục lục, danh mục bảng/biểu đồ/hình tự động; chuẩn hoá chú
  thích; thể thức hành chính (cỡ, kiểu chữ, đường kẻ, khối chữ ký); ảnh; bảng số liệu; số trang.
- **Tạo văn bản mẫu** (`scripts/tao_mau.py`): báo cáo, kế hoạch, tờ trình, công văn… đúng thể thức.
- **Quy chuẩn tham chiếu**: [references/the-thuc-nd30.md](references/the-thuc-nd30.md).

Script không tự đánh lại số chương, mục, bảng và không đổi dấu thập phân hàng loạt. Các lỗi này chỉ được liệt kê
để người viết tự quyết định.

## Cài đặt

```bash
python3 -m venv .venv && . .venv/bin/activate
pip install python-docx
```

## Sử dụng

```bash
python scripts/audit_docx.py tai-lieu.docx                          # kiểm tra
python scripts/fix_docx.py tai-lieu.docx -o tai-lieu-chuan.docx --all   # sửa toàn bộ
python scripts/fix_docx.py tai-lieu.docx -o out.docx --headings --captions --toc   # chỉ mục lục, danh mục
python scripts/fix_docx.py --help                                   # các bước và tuỳ chọn

python scripts/tao_mau.py -o bao-cao.docx --loai "BÁO CÁO" --co-quan "TÊN CƠ QUAN" --so 5 \
    --ky-hieu BC-XYZ --ngay 2026-11-19 --trich-yeu "Kết quả công tác năm 2026" \
    --chuc-vu "Giám đốc" --nguoi-ky "Nguyễn Văn A" --noi-nhan "Bộ Y tế;Ban Giám đốc"
```

Mở file kết quả bằng Word rồi nhấn **Ctrl+A → F9** để cập nhật mục lục, danh mục và số trang.

Thư mục [`mau/`](mau/) có văn bản mẫu và một cặp ví dụ trước/sau khi chuẩn hoá.

## Giấy phép

[MIT](LICENSE)
