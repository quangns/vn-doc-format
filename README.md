# Thể thức văn bản theo Nghị định 30/2020/NĐ-CP – Công cụ chuẩn hoá file Word (.docx) tự động

[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
![Python 3.9+](https://img.shields.io/badge/Python-3.9%2B-blue.svg)
![Định dạng: .docx](https://img.shields.io/badge/%C4%91%E1%BB%8Bnh%20d%E1%BA%A1ng-.docx-informational.svg)

Công cụ mã nguồn mở giúp **kiểm tra và tự động định dạng văn bản Word đúng thể thức Nghị định 30/2020/NĐ-CP về công
tác văn thư**. Công cụ cũng chuẩn hoá **luận văn, khoá luận, báo cáo khoa học, báo cáo tổng kết, công văn, tờ trình,
kế hoạch** theo cách trình bày phổ biến ở Việt Nam. Với một file `.docx`, công cụ có thể:

- **tạo mục lục tự động** và **danh mục bảng, danh mục biểu đồ, danh mục hình** tự động;
- **đánh số chương, mục, tiểu mục** thống nhất (số La Mã, số Ả Rập);
- đặt **font Times New Roman**, **cỡ chữ 13–14**, **căn lề chuẩn** (trên, dưới 20–25 mm; trái 30–35 mm; phải 15–20 mm),
  khổ A4, giãn dòng, **đánh số trang**;
- trình bày đúng **quốc hiệu, tiêu ngữ**, tên cơ quan, số và ký hiệu, địa danh và ngày tháng, tên loại, trích yếu,
  **nơi nhận**, **vị trí chữ ký**;
- định dạng **bảng biểu, hình ảnh** và chú thích;
- rà soát **văn phong hành chính, văn phong khoa học**.

Dùng được qua dòng lệnh (Python) hoặc làm skill cho Claude Code ([SKILL.md](SKILL.md)).

> **Từ khoá:** thể thức văn bản hành chính, định dạng văn bản theo Nghị định 30, cách trình bày văn bản hành chính,
> font chữ văn bản hành chính, căn lề văn bản chuẩn, cỡ chữ quốc hiệu tiêu ngữ, tạo mục lục tự động trong Word,
> danh mục bảng biểu luận văn, format luận văn thạc sĩ, chuẩn hoá báo cáo, kiểm tra lỗi trình bày văn bản.
> (Không dấu: the thuc van ban nghi dinh 30, dinh dang van ban hanh chinh, tao muc luc tu dong word, format luan van.)

---

## Mục lục

- [Dùng cho ai?](#dùng-cho-ai)
- [Tính năng](#tính-năng)
- [Tóm tắt quy định thể thức Nghị định 30](#tóm-tắt-quy-định-thể-thức-nghị-định-30)
- [Cài đặt](#cài-đặt)
- [Hướng dẫn sử dụng](#hướng-dẫn-sử-dụng)
- [Ví dụ kết quả kiểm tra](#ví-dụ-kết-quả-kiểm-tra)
- [Câu hỏi thường gặp](#câu-hỏi-thường-gặp)
- [English summary](#english-summary)
- [Giấy phép](#giấy-phép)

## Dùng cho ai?

- **Cán bộ văn thư, chuyên viên** soạn công văn, báo cáo, tờ trình, kế hoạch, quyết định cần đúng thể thức
  Nghị định 30/2020/NĐ-CP.
- **Học viên cao học, sinh viên, nghiên cứu sinh** cần định dạng luận văn, khoá luận, luận án: mục lục, danh mục
  bảng, danh mục biểu đồ, đánh số chương mục, chú thích bảng và hình.
- **Nhóm nghiên cứu, đơn vị sự nghiệp** cần chuẩn hoá báo cáo khoa học, báo cáo đề tài.

## Tính năng

### 1. Kiểm tra thể thức văn bản – `scripts/audit_docx.py`

Chỉ đọc file và liệt kê lỗi kèm vị trí, không sửa gì:

- khổ giấy, lề trang, font chữ, cỡ chữ, số trang;
- tiêu đề chưa gán style (nên không lên mục lục), **số chương, mục bị trùng hoặc nhảy cóc**, tên chương ghi không
  thống nhất (`Chương 1.` / `Chương II:`);
- chú thích **bảng, biểu đồ, hình** trùng số, thiếu số, sai thứ tự, sai dấu, bị ngắt dòng; mục lục gõ tay; lỗi
  "Error! Bookmark not defined";
- **thành phần thể thức**: quốc hiệu, tiêu ngữ, tên cơ quan, số ký hiệu, ngày tháng, tên loại, trích yếu, nơi nhận,
  chữ ký. Kiểm tra cỡ chữ, in đậm, in nghiêng, in hoa, căn giữa, đường kẻ;
- **ảnh**: ảnh nổi (wrap), ảnh tràn lề, ảnh chưa căn giữa, chú thích sai vị trí; **bảng**: tên bảng sai vị trí, không có
  viền, không lặp dòng tiêu đề, tràn lề;
- **văn phong**: khẩu ngữ, từ chủ quan, câu quá dài, lặp từ, dấu câu, số viết liền đơn vị, chữ viết tắt chưa giải
  thích, i/y và dấu thanh không thống nhất (`hoà`/`hòa`), dấu thập phân.

### 2. Tự động sửa – `scripts/fix_docx.py`

| Tuỳ chọn | Việc làm |
|---|---|
| `--page` | khổ A4, căn lề 20–20–30–15 mm |
| `--font` `--body` | font Times New Roman, cỡ 13, căn đều, thụt đầu dòng 1,27 cm, cách đoạn 6 pt |
| `--headings` `--heading-format` | gán style tiêu đề theo số thứ tự, định dạng đậm và nghiêng thống nhất theo cấp |
| `--toc` | **mục lục tự động** (tối đa 3 cấp), **danh mục bảng, biểu đồ, hình** tự động, xoá mục lục gõ tay |
| `--captions` | chuẩn hoá chú thích `Bảng 3.1.`, `Biểu đồ 3.1.`, `Hình 1.1.` |
| `--chapter arabic\|roman` | thống nhất `Chương 1.` hoặc `Chương I.` |
| `--admin` | quốc hiệu, tiêu ngữ, tên cơ quan, `Số: 05/…`, `ngày 05 tháng 02`, nơi nhận, khối chữ ký, đường kẻ |
| `--images` `--tables` | căn giữa và thu nhỏ ảnh vừa lề; bảng rộng bằng vùng chữ, có viền, dòng tiêu đề lặp lại |
| `--page-number` | số trang giữa lề trên, ẩn ở trang đầu |
| `--van-phong` `--don-vi` `--dau-thanh` | khoảng trắng, dấu câu, `23,5mm` → `23,5 mm`, dấu thanh `hòa/thủy` |
| `--all` | chạy tất cả |

Công cụ **không** tự đánh lại số chương, bảng và **không** sửa câu chữ. Các lỗi nội dung chỉ được liệt kê để người
viết quyết định.

### 3. Tạo văn bản hành chính mẫu – `scripts/tao_mau.py`

Sinh khung **báo cáo, kế hoạch, tờ trình, công văn**… đúng thể thức: bảng đầu trang (cơ quan, quốc hiệu, tiêu ngữ),
tên loại và trích yếu, khối nơi nhận và chữ ký, số trang. Xem file mẫu trong thư mục [`mau/`](mau/).

## Tóm tắt quy định thể thức Nghị định 30

Bản đầy đủ: [references/the-thuc-nd30.md](references/the-thuc-nd30.md). Tài liệu này gồm bố cục trang, đánh số,
bảng biểu, mục lục, chữ số, ngày tháng, viết hoa, chữ viết tắt và văn phong.

| Thành phần | Cỡ chữ | Kiểu chữ |
|---|---|---|
| Quốc hiệu `CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM` | 12–13 | IN HOA, đứng, đậm |
| Tiêu ngữ `Độc lập - Tự do - Hạnh phúc` | 13–14 | đứng, đậm, có đường kẻ bên dưới |
| Tên cơ quan ban hành | 12–13 | IN HOA, đứng, đậm |
| `Số: 05/BC-UBND` | 13 | đứng |
| Địa danh, ngày tháng | 13–14 | nghiêng |
| Tên loại văn bản (BÁO CÁO, KẾ HOẠCH…) | 13–14 | IN HOA, đứng, đậm |
| Nội dung văn bản | 13–14 | đứng; giãn dòng 1–1,5; cách đoạn ≥ 6 pt |
| Quyền hạn, chức vụ người ký | 13–14 | IN HOA, đứng, đậm |
| Họ tên người ký | 13–14 | đứng, đậm |
| `Nơi nhận:` / danh sách nơi nhận | 12 / 11 | nghiêng đậm / đứng |

Khổ A4, font **Times New Roman** (Unicode TCVN 6909:2001). **Lề trên, dưới 20–25 mm; lề trái 30–35 mm; lề phải
15–20 mm.** Số trang đặt giữa lề trên, không hiện ở trang đầu.

## Cài đặt

Cần Python 3.9 trở lên.

```bash
git clone https://github.com/quangns/the-thuc-van-ban-nghi-dinh-30.git
cd the-thuc-van-ban-nghi-dinh-30
python3 -m venv .venv && . .venv/bin/activate      # Windows: .venv\Scripts\activate
pip install python-docx
```

## Hướng dẫn sử dụng

```bash
# 1. Kiểm tra lỗi thể thức, văn phong
python scripts/audit_docx.py "bao-cao.docx"

# 2. Chuẩn hoá toàn bộ (không ghi đè file gốc)
python scripts/fix_docx.py "bao-cao.docx" -o "bao-cao-chuan.docx" --all

# Chỉ tạo mục lục, danh mục bảng, danh mục biểu đồ cho luận văn
python scripts/fix_docx.py "luan-van.docx" -o "luan-van-muc-luc.docx" --headings --captions --toc

# 3. Tạo báo cáo mẫu đúng thể thức
python scripts/tao_mau.py -o bao-cao.docx --loai "BÁO CÁO" --chu-quan "UBND TỈNH A" --co-quan "SỞ Y TẾ" \
    --so 5 --ky-hieu BC-SYT --dia-danh "Hà Nội" --ngay 2026-11-19 --trich-yeu "Kết quả công tác năm 2026" \
    --quyen-han KT --cap-tren "Giám đốc" --chuc-vu "Phó Giám đốc" --nguoi-ky "Nguyễn Văn A" \
    --noi-nhan "UBND tỉnh;Ban Giám đốc"

python scripts/fix_docx.py --help     # xem đầy đủ tuỳ chọn
```

Sau khi chạy, mở file bằng **Microsoft Word** rồi nhấn **Ctrl + A → F9** để Word cập nhật mục lục, danh mục và số
trang.

## Ví dụ kết quả kiểm tra

```text
== KIỂM TRA THỂ THỨC: vi-du-sai-the-thuc.docx
## Trang giấy (khổ, lề)
  - section 1: khổ 215.9×279.4 mm (cần A4 210×297); lề trên 25.4 (20–25); lề phải 31.8 (15–20)
  - chưa có số trang (giữa lề trên, ẩn trang đầu)
## Cỡ/kiểu chữ thành phần thể thức
  - Quốc hiệu "CỘNG HOÀ XÃ HỘI CHỦ NGHĨA VIỆT NAM": cỡ 14 (cần 12–13), cần in đậm, cần canh giữa
  - Tiêu ngữ "Độc Lập – Tự Do – Hạnh Phúc": cỡ 11 (cần 13–14), cần in đậm, cần canh giữa
  - "Nơi nhận:" "Nơi nhận": cỡ 11 (cần 12), cần in đậm, cần in nghiêng
## Cách ghi thành phần thể thức
  - "Số 5/BC-SYT": số nhỏ hơn 10 phải thêm 0 (Số: 05/...)
  - "Hà Nội, ngày 3 tháng 2 năm 2026": ngày < 10 và tháng 1, 2 phải thêm 0
```

Sau khi chạy `fix_docx.py --all`, kết quả là [`mau/vi-du-sau-khi-sua.docx`](mau/vi-du-sau-khi-sua.docx), không còn
lỗi.

## Câu hỏi thường gặp

**Nghị định 30/2020 quy định font chữ và cỡ chữ văn bản hành chính như thế nào?**
Font Times New Roman (bộ mã Unicode TCVN 6909:2001), màu đen. Cỡ chữ phần nội dung từ 13 đến 14. Quốc hiệu cỡ 12–13
in hoa đậm; tiêu ngữ cỡ 13–14 đậm.

**Căn lề văn bản theo Nghị định 30 là bao nhiêu?**
Lề trên 20–25 mm, lề dưới 20–25 mm, lề trái 30–35 mm, lề phải 15–20 mm, khổ giấy A4.

**Làm sao tạo mục lục, danh mục bảng, danh mục biểu đồ tự động cho luận văn?**
Chạy `fix_docx.py --headings --captions --toc`. Công cụ gán style cho tiêu đề và chú thích, rồi chèn trường mục lục
(TOC) của Word. Mở file, nhấn Ctrl + A → F9 để cập nhật số trang.

**Ngày tháng trong văn bản hành chính ghi thế nào?**
`Hà Nội, ngày 05 tháng 02 năm 2026`: ngày nhỏ hơn 10 và tháng 1, 2 thêm số 0 phía trước, dòng này in nghiêng.

**Công cụ có sửa nội dung văn bản không?**
Không. Công cụ chỉ sửa trình bày (định dạng, khoảng trắng, dấu câu, số 0 trong ngày tháng…). Số chương, bảng bị trùng
hay câu văn chưa chuẩn chỉ được liệt kê để người viết tự sửa. File gốc không bị ghi đè.

**Có dùng được với Google Docs, LibreOffice không?**
Công cụ đọc và ghi định dạng `.docx`. Có thể tải file từ Google Docs về dạng `.docx` để xử lý. Riêng mục lục cần mở
bằng Microsoft Word để cập nhật (LibreOffice: Tools → Update → Update All).

## English summary

**Vietnamese document formatting checker and auto-formatter for Word (.docx)** following *Decree 30/2020/ND-CP* on
clerical work (official administrative document format in Vietnam) and common conventions for theses and research
reports. It audits and fixes page setup, Times New Roman fonts, margins, heading numbering, automatic table of
contents and lists of tables/figures, captions, the national title and motto block, signature block, tables, images,
page numbers, and technical writing style. Written in Python with `python-docx`; also usable as a Claude Code skill.

## Giấy phép

Phát hành theo giấy phép [MIT](LICENSE). Bạn được tự do sử dụng, sửa đổi và phân phối lại.

Tài liệu quy chuẩn trong repo là bản tóm tắt để tra cứu nhanh. Khi cần căn cứ pháp lý, hãy đối chiếu với văn bản
gốc của Nghị định 30/2020/NĐ-CP.
