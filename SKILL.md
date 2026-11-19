---
name: chuan-hoa-tai-lieu
description: Chuẩn hoá tài liệu Word (.docx) theo thể thức báo cáo nhà nước (Nghị định 30/2020/NĐ-CP) và quy ước báo cáo khoa học, luận văn. Gồm tạo mục lục, danh mục bảng/biểu đồ/hình tự động; đánh số chỉ mục, chương (số La Mã/Ả Rập); chú thích bảng biểu; phông, cỡ chữ, lề, giãn dòng; văn phong kỹ thuật (khẩu ngữ, câu dài, viết tắt, dấu câu, chính tả thống nhất); thể thức văn bản hành chính (quốc hiệu, tiêu ngữ, tên cơ quan, số ký hiệu, địa danh ngày tháng, tên loại, trích yếu, nơi nhận, vị trí và định dạng chữ ký); cỡ chữ, đậm/nghiêng từng thành phần và từng cấp tiêu đề; vị trí, kích thước ảnh; định dạng bảng; số trang; tạo văn bản mẫu đúng thể thức; kiểm tra số liệu, dấu thập phân, ngày tháng. Dùng khi người dùng nói "đánh mục lục", "danh mục bảng", "chuẩn hoá format", "format theo chuẩn nhà nước", "Nghị định 30", "kiểm tra thể thức", "quốc hiệu tiêu ngữ", "chữ ký", "nơi nhận", "soạn công văn/báo cáo/tờ trình mẫu", "văn phong", "rà soát câu chữ", "luận văn", "báo cáo".
---

# Chuẩn hoá tài liệu theo thể thức nhà nước

Quy chuẩn chi tiết: [references/the-thuc-nd30.md](references/the-thuc-nd30.md). Đọc file này trước khi
sửa, và khi báo cáo cho người dùng thì trích đúng quy định.

## Môi trường

```bash
PY=/opt/shared/venvs/docx/bin/python        # venv dùng chung, có python-docx
S=/workspace/chuan-hoa-tai-lieu/scripts
# nếu venv chưa có:
# UV_PYTHON_INSTALL_DIR=/opt/shared/uv/python UV_CACHE_DIR=/opt/shared/uv/cache uv venv /opt/shared/venvs/docx
# VIRTUAL_ENV=/opt/shared/venvs/docx uv pip install python-docx
```

Máy không có Word/LibreOffice, nên số trang của mục lục chỉ có sau khi người dùng mở file và bấm **Ctrl+A → F9**.
Luôn nhắc người dùng bước này.

## Quy trình

1. **Sao lưu.** Giữ bản gốc: `"<tên> (bản gốc trước khi chuẩn hoá).docx"` (`cp -n`). Làm việc trên bản sao ở `/tmp`.
   Tên file tiếng Việt có thể ở dạng Unicode NFD: lấy tên bằng `ls`/glob, đừng gõ lại tên, nếu không sẽ sinh
   ra 2 file trùng tên hiển thị.

2. **Kiểm tra** (chỉ đọc):
   ```bash
   $PY $S/audit_docx.py file.docx          # thêm --json nếu cần xử lý tiếp
   ```
   Báo cáo các mục: khổ và lề trang, phông và cỡ chữ, tiêu đề chưa gán style (nên không lên mục lục), tiêu đề cùng cấp
   dùng style khác nhau, số chỉ mục trùng hoặc nhảy cóc, tên chương ghi không thống nhất, chú thích bảng/biểu đồ/hình
   trùng số, thiếu số, sai dấu hoặc bị tách đoạn, mục lục gõ tay hoặc lỗi "Error! Bookmark", dấu thập phân, dấu ±.
   Với văn bản hành chính, kiểm tra thêm: thiếu thành phần thể thức; cỡ, đậm, nghiêng, in hoa, canh lề của từng thành phần
   (so với bảng ở mục 1a của reference); quốc hiệu, tiêu ngữ ghi sai; thiếu đường kẻ; số văn bản, ngày tháng thiếu số 0;
   dấu cuối dòng nơi nhận; khối chữ ký không nằm cùng hàng với nơi nhận. Kiểm tra cả ảnh (nổi, tràn lề, chưa canh giữa,
   chú thích sai vị trí), bảng (tên bảng sai vị trí, không viền, không lặp dòng tiêu đề, tràn lề, cỡ chữ) và số trang.
   **Văn phong kỹ thuật** (`scripts/van_phong.py`, tắt bằng `--no-van-phong`): khẩu ngữ, từ cảm tính, chủ quan, dấu
   chấm than; câu trên 70 âm tiết, đoạn quá dài; lặp từ; dấu câu và khoảng trắng; khoảng giá trị và ngoặc kép không
   thống nhất; số viết liền đơn vị; chữ viết tắt chưa giải thích lần đầu, chưa có hoặc thừa trong Danh mục chữ viết tắt,
   chỉ dùng 1 lần; i/y và dấu thanh không thống nhất; trộn "tôi" với "chúng tôi". Bỏ qua mục lục, tài liệu tham khảo,
   phụ lục.
   Số thứ tự `[i]` là chỉ số đoạn trong `doc.paragraphs`.

3. **Xem cấu trúc phần đầu** trước khi sửa: in khoảng 150 đoạn đầu kèm style và `instrText` để biết mục lục
   cũ là field hay gõ tay, và các tiêu đề "MỤC LỤC", "DANH MỤC …" nằm ở đâu.

4. **Sửa tự động** (không ghi đè file nguồn):
   ```bash
   $PY $S/fix_docx.py in.docx -o out.docx --all                 # toàn bộ
   $PY $S/fix_docx.py in.docx -o out.docx --headings --captions --toc   # chỉ mục lục + danh mục
   ```
   | Cờ | Việc làm |
   |---|---|
   | `--page` | A4, lề `--margins 20,20,30,15` (trên,dưới,trái,phải mm) |
   | `--font` / `--font-runs` | Times New Roman cho style + docDefaults / cả run gán phông trực tiếp |
   | `--body` | Normal: cỡ `--size 13`, canh đều, thụt 1,27 cm, cách sau 6 pt, bỏ Tab đầu dòng; `--line 1.3` giãn dòng |
   | `--headings` | gán style tiêu đề theo số thứ tự (Chương/I. → cấp 1, 1.1./1. → cấp 2, 1.1.1. → cấp 3, 1.1.1.1. → cấp 4); dùng lại style tiêu đề sẵn có, nhân bản nếu một style đang gánh 2 cấp; đặt outline level |
   | `--chapter arabic\|roman` | thống nhất `Chương 1.` hoặc `Chương I.` |
   | `--heading-dot` | `3.1 Tên` → `3.1. Tên` |
   | `--captions` | style riêng cho từng loại Bảng/Biểu đồ/Hình; `Bảng 3.4:` → `Bảng 3.4.`; ghép chú thích bị Enter/Shift+Enter giữa câu; sửa "Biều đồ" |
   | `--toc` | xoá dòng mục lục gõ tay, thay field cũ hoặc chèn field TOC mới sau "MỤC LỤC", "DANH MỤC BẢNG/BIỂU ĐỒ/HÌNH (ẢNH)"; bật `updateFields`. Mục lục **chỉ tới cấp 3** (`1.1.1.`), không đưa `a)`, `b)`, `c)` vào; đổi bằng `--toc-levels` |
   | `--heading-format` | đặt định dạng cho style từng cấp: cấp 1 đậm cỡ 14 canh giữa (mục `I.` của báo cáo hành chính thì canh trái), cấp 2 đậm, cấp 3 đậm nghiêng, cấp 4 nghiêng; xoá đậm/nghiêng/cỡ/màu gán tay trong tiêu đề |
   | `--admin` | thể thức NĐ 30: sửa chữ quốc hiệu, tiêu ngữ; cỡ, đậm, nghiêng, in hoa, canh giữa cho từng thành phần; kẻ đường dưới tiêu ngữ, tên cơ quan, trích yếu (nếu chưa có); `Số: 05/`, `ngày 05 tháng 02`; nơi nhận (`Nơi nhận:`, `-` đầu dòng, `;` và `.` cuối dòng, `- Lưu: `); khối chữ ký in hoa, đậm, canh giữa, chừa ít nhất 3 dòng trống để ký |
   | `--images` | ảnh cùng dòng chữ: canh giữa, bỏ thụt, thu nhỏ theo tỷ lệ nếu tràn lề, gắn liền với chú thích dưới |
   | `--tables` | chỉ bảng có tên "Bảng …": rộng bằng vùng chữ, canh giữa, viền đơn 0,5 pt, dòng đầu đậm và lặp lại khi sang trang, không ngắt dòng giữa trang, cỡ `--table-size 12`, ô số canh giữa; `--tables-all` để chuẩn hoá cả bảng không có tên |
   | `--van-phong` | sửa lỗi trình bày văn phong có quy tắc rõ: khoảng trắng trước/sau dấu câu, trong ngoặc, khoảng trắng kép, `v.v...` → `v.v.`, `±` có cách hai bên, bỏ hư từ lặp (`trong trong`) |
   | `--don-vi` | cách số và đơn vị: `23,5mm` → `23,5 mm` (không đụng `%`, `°` và đơn vị 1 chữ dễ nhầm như `1A`, `3D`) |
   | `--dau-thanh moi\|cu` | thống nhất vị trí dấu thanh `hòa, khỏe, thủy` (mới) hoặc `hoà, khoẻ, thuỷ` (cũ); không nằm trong `--all` |
   | `--page-number` | số trang giữa lề trên, ẩn trang đầu (chỉ khi chưa có số trang và header trống) |
   | `--all` | tất cả trừ `--chapter`, `--line`, `--tables-all`, `--dau-thanh` |

   Muốn đổi cỡ chữ mặc định của thành phần thể thức thì sửa `SIZES` trong `scripts/the_thuc.py` (vẫn phải nằm
   trong khoảng NĐ 30 ghi chú bên cạnh).

   Đọc kỹ log, nhất là các dòng `!!`. Rồi chạy lại `audit_docx.py` trên file kết quả để xác nhận.

   **Tạo văn bản mới đúng thể thức** (báo cáo, kế hoạch, tờ trình, công văn…):
   ```bash
   $PY $S/tao_mau.py -o bao-cao.docx --loai "BÁO CÁO" --chu-quan "BỘ Y TẾ" --co-quan "BỆNH VIỆN X" \
       --so 5 --ky-hieu BC-BVX --dia-danh "Hà Nội" --ngay 2026-11-19 --trich-yeu "Kết quả công tác năm 2026" \
       --quyen-han KT --cap-tren "Giám đốc" --chuc-vu "Phó Giám đốc" --nguoi-ky "Nguyễn Văn A" \
       --noi-nhan "Bộ Y tế;Ban Giám đốc"
   $PY $S/tao_mau.py -o cv.docx --loai "CÔNG VĂN" --trich-yeu "Về việc ..." --kinh-gui "Sở Y tế;Bệnh viện A"
   ```
   File mẫu có sẵn trong `mau/`: `bao-cao-mau.docx`, `cong-van-mau.docx`, và một cặp ví dụ trước/sau khi chạy
   `fix_docx.py --all` (`vi-du-sai-the-thuc.docx` → `vi-du-sau-khi-sua.docx`).
   Kết quả: khổ A4, lề chuẩn, bảng đầu trang và bảng chữ ký 2 cột không viền, các đường kẻ, số trang. Phần nội dung là
   đoạn mẫu để thay bằng nội dung thật.

5. **Sửa tay phần script không làm** (viết Python riêng bằng python-docx khi cần):
   - Tiêu đề không theo mẫu số (ví dụ `2 Tồn tại` thiếu dấu chấm, không in đậm) thì gán style bằng tay.
   - Số trang La Mã cho phần đầu: cần section break, hướng dẫn người dùng làm trong Word nếu không chắc chắn.
   - Ảnh nổi (wrap text) hoặc đường kẻ vẽ bằng shape lệch vị trí: hướng dẫn người dùng chọn "In line with text" trong Word.
   - Văn bản hành chính không trình bày bằng bảng 2 cột (quốc hiệu và tên cơ quan nằm ở 2 đoạn rời): script vẫn định dạng
     chữ nhưng không dàn được 2 cột. Dựng lại phần đầu bằng `tao_mau.py` rồi chép nội dung sang.

6. **Văn phong: chỉ tự sửa lỗi trình bày.** Câu dài, khẩu ngữ, chủ quan, viết tắt, i/y, ngôi xưng thì liệt kê kèm vị trí
   và gợi ý theo mục 10 của reference. Nếu người dùng nhờ viết lại câu thì đề xuất câu mới để họ duyệt, không tự thay
   hàng loạt. Giữ nguyên trích dẫn, tên riêng, thuật ngữ chuyên ngành.

7. **Không tự sửa nội dung học thuật/số liệu.** Không đánh lại số chương, mục, bảng khi bị trùng/thiếu; không đổi
   dấu thập phân hàng loạt (dấu phẩy còn là dấu liệt kê, dấu chấm còn là phân cách nghìn). Liệt kê cụ thể để
   người viết quyết định. Chỉ sửa chính tả hiển nhiên ở chú thích ("Biều" → "Biểu") và ghi rõ trong báo cáo.

8. **Lưu và báo cáo.** Ghi đè file đích (giữ bản gốc), sau đó báo cáo theo 3 nhóm: đã sửa tự động, cần người dùng bấm
   F9, lỗi nội dung cần tác giả quyết định (kèm vị trí và trích đoạn).

## Lưu ý kỹ thuật (đã gặp)

- **Nhân bản style từ Normal phải xoá `w:default="1"`.** Nếu không, cả văn bản sẽ nhận style mới làm style mặc định.
  `new_style()` trong `fix_docx.py` đã xử lý việc này.
- **Dấu phân cách trong `\t "style,1,style,2"`** phụ thuộc cài đặt vùng của Word. Dùng theo field có sẵn trong file
  (`list_sep()`), mặc định là dấu phẩy.
- **Field TOC cũ trải trên nhiều đoạn**: từ `fldChar begin` đến `fldChar end` khớp cấp. Thay cả vùng đó, không chèn
  thêm field mới bên cạnh (`_field_end()`).
- **Mục lục gõ tay** thường có Tab + số trang, hoặc là dòng trùng chữ với tiêu đề thật phía sau (ví dụ "DANH MỤC BẢNG"
  nằm ngay dưới "MỤC LỤC").
- **Thứ tự phần tử OOXML**: `w:outlineLvl` phải đứng trước `w:rPr` trong `w:pPr`; `w:rFonts` phải đứng sau `w:rStyle`;
  không lồng `w:r` trong `w:r`. Sai thứ tự thì Word báo "unreadable content".
- **Tiêu đề phần đầu đã có style riêng** (ví dụ `001`, `Heading 1`): giữ nguyên style, chỉ thêm style đó vào field mục lục.
- **Chú thích ngay dưới hình** (biểu đồ, hình): không ghép đoạn có ảnh (`w:drawing`) hoặc đoạn bắt đầu bằng "Nguồn",
  "Nhận xét".
- **Nhận diện thành phần thể thức** dựa trên chữ (`the_thuc.detect_layout`): quốc hiệu, tiêu ngữ, `Số:` và địa danh
  ngày tháng tìm trong khoảng 60 đoạn đầu (kể cả trong bảng); tên cơ quan là các dòng IN HOA đứng trước `Số:` trong
  cùng ô; khối chữ ký là ô bên phải cùng hàng với `Nơi nhận:`, trong đó dòng cuối là họ tên, các dòng trên là chức vụ.
  Nếu người dùng bố cục khác, chạy `audit_docx.py` trước để xem script nhận ra những thành phần nào.
- **Đường kẻ** dưới tiêu ngữ, tên cơ quan, trích yếu: là một đoạn trống cao 2 pt có viền dưới và thụt trái/phải đều
  nhau để vạch kẻ có độ dài mong muốn (`add_rule_after`). Không thêm nếu đã có shape, viền hoặc gạch chân.
- **Bảng bố cục và bảng số liệu**: bảng có chữ quốc hiệu, "Nơi nhận", "Số:", hoặc không viền và ≤ 2 cột ≤ 4 dòng là bảng
  bố cục, không động vào. Bảng số liệu chỉ tính bảng có tên "Bảng …" ngay trên hoặc dưới.
