#!/usr/bin/env python3
"""Tạo văn bản hành chính mẫu đúng thể thức Nghị định 30/2020 (khung để điền nội dung).

    python tao_mau.py -o bao-cao.docx --loai "BÁO CÁO" --co-quan "BỆNH VIỆN MẮT TRUNG ƯƠNG" \\
        --chu-quan "BỘ Y TẾ" --so 5 --ky-hieu BC-BVMTW --dia-danh "Hà Nội" --ngay 2026-11-19 \\
        --trich-yeu "Kết quả công tác năm 2026" --quyen-han KT --cap-tren "GIÁM ĐỐC" \\
        --chuc-vu "PHÓ GIÁM ĐỐC" --nguoi-ky "Nguyễn Văn A" --noi-nhan "Bộ Y tế;Ban Giám đốc"
    python tao_mau.py -o cong-van.docx --loai "CÔNG VĂN" --trich-yeu "Về việc ..." --kinh-gui "Sở Y tế Hà Nội"

Bố cục (Phụ lục I NĐ 30): bảng 2 cột không viền ở đầu trang (trái: cơ quan, số ký hiệu, V/v; phải: quốc hiệu,
tiêu ngữ, địa danh ngày tháng) → tên loại + trích yếu → nội dung → bảng 2 cột cuối (trái: nơi nhận; phải: chữ ký).
"""
import argparse, datetime, os, sys
import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Mm, Pt, Cm

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import the_thuc as T  # noqa: E402

C, L, J = WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.JUSTIFY
TEXT_W = 165   # mm: 210 - 30 - 15


def no_border_table(doc, widths_mm):
    t = doc.add_table(rows=1, cols=len(widths_mm))
    t.autofit = False
    tblPr = t._tbl.tblPr
    b = T.ensure(tblPr, 'tblBorders', 'tblPr')
    for side in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        e = OxmlElement('w:' + side)
        e.set(qn('w:val'), 'nil')
        b.append(e)
    T.ensure(tblPr, 'tblLayout', 'tblPr').set(qn('w:type'), 'fixed')
    ind = T.ensure(tblPr, 'tblInd', 'tblPr')       # bù lề trong ô để chữ thẳng lề trang
    ind.set(qn('w:w'), '-108')
    ind.set(qn('w:type'), 'dxa')
    for i, w in enumerate(widths_mm):
        t.columns[i].width = Mm(w)
        t.rows[0].cells[i].width = Mm(w)
    return t


def cell_par(cell, text, key, first=False, align=C):
    p = cell.paragraphs[0] if first else cell.add_paragraph()
    if text:
        p.add_run(text)
    T.fmt(p, key, align=align)
    return p


def blank(cell, size=13):
    p = cell.add_paragraph()
    T.fmt(p, size=size)
    return p


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('-o', '--out', required=True)
    ap.add_argument('--loai', default='BÁO CÁO', help='tên loại: BÁO CÁO, KẾ HOẠCH, TỜ TRÌNH, CÔNG VĂN, ...')
    ap.add_argument('--chu-quan', default='', help='tên cơ quan chủ quản (nếu có)')
    ap.add_argument('--co-quan', default='TÊN CƠ QUAN, TỔ CHỨC')
    ap.add_argument('--so', default='', help='số văn bản (để trống = "   ")')
    ap.add_argument('--ky-hieu', default='BC-...')
    ap.add_argument('--dia-danh', default='Hà Nội')
    ap.add_argument('--ngay', default='', help='YYYY-MM-DD; để trống = "ngày   tháng   năm ..."')
    ap.add_argument('--trich-yeu', default='Trích yếu nội dung văn bản')
    ap.add_argument('--kinh-gui', default='', help='công văn: nơi nhận chính, nhiều nơi cách nhau bằng ;')
    ap.add_argument('--quyen-han', default='', choices=['', 'TM', 'KT', 'Q', 'TL', 'TUQ'])
    ap.add_argument('--cap-tren', default='', help='chức vụ cấp trên khi ký KT./TL./TUQ. hoặc tên tập thể khi TM.')
    ap.add_argument('--chuc-vu', default='CHỨC VỤ NGƯỜI KÝ')
    ap.add_argument('--nguoi-ky', default='Họ và tên')
    ap.add_argument('--noi-nhan', default='Như trên', help='danh sách cách nhau bằng ;')
    ap.add_argument('--luu', default='VT', help='Lưu: VT, đơn vị soạn thảo')
    ap.add_argument('--size', type=float, default=14, help='cỡ chữ nội dung (13–14)')
    a = ap.parse_args()
    cong_van = a.loai.upper() in ('CÔNG VĂN', 'CV')

    doc = docx.Document()
    s = doc.sections[0]
    s.page_width, s.page_height = Mm(210), Mm(297)
    s.top_margin, s.bottom_margin, s.left_margin, s.right_margin = Mm(20), Mm(20), Mm(30), Mm(15)
    st = doc.styles['Normal']
    st.font.name = 'Times New Roman'
    rf = st.element.get_or_add_rPr().get_or_add_rFonts()
    for k in ('w:ascii', 'w:hAnsi', 'w:cs', 'w:eastAsia'):
        rf.set(qn(k), 'Times New Roman')
    st.font.size = Pt(a.size)
    pf = st.paragraph_format
    pf.space_before, pf.space_after, pf.line_spacing = Pt(0), Pt(6), 1.15
    T.step_page_number(doc)

    # ---------- đầu văn bản ----------
    left_w, right_w = 62, TEXT_W - 62
    hd = no_border_table(doc, [left_w, right_w])
    lc, rc = hd.rows[0].cells
    first = True
    if a.chu_quan:
        cell_par(lc, a.chu_quan.upper(), 'co_quan_chu_quan', first=True)
        first = False
    p = cell_par(lc, a.co_quan.upper(), 'co_quan', first=first)
    cq_w = T.text_width_mm(a.co_quan.upper(), T.SIZES['co_quan'][0], True)
    T.add_rule_after(p, max(15, min(cq_w, left_w - 4) * 0.4), left_w - 3.8)
    so = f'{int(a.so):02d}' if a.so.isdigit() else (a.so or '      ')
    ky_hieu = a.ky_hieu if not cong_van or a.ky_hieu != 'BC-...' else '...'
    cell_par(lc, f'Số: {so}/{ky_hieu}', 'so_ky_hieu')
    if cong_van:
        cell_par(lc, f'V/v {a.trich_yeu[0].lower() + a.trich_yeu[1:] if a.trich_yeu.startswith("Về") else a.trich_yeu}',
                 'vv')
    p = cell_par(rc, T.QUOC_HIEU, 'quoc_hieu', first=True)
    p = cell_par(rc, T.TIEU_NGU, 'tieu_ngu')
    T.add_rule_after(p, T.text_width_mm(T.TIEU_NGU, T.SIZES['tieu_ngu'][0], True), right_w - 3.8)
    if a.ngay:
        d = datetime.date.fromisoformat(a.ngay)
        thang = f'{d.month:02d}' if d.month <= 2 else str(d.month)
        ngay = f'{a.dia_danh}, ngày {d.day:02d} tháng {thang} năm {d.year}'
    else:
        ngay = f'{a.dia_danh}, ngày    tháng    năm {datetime.date.today().year}'
    blank(rc, 6)
    cell_par(rc, ngay, 'dia_danh')

    # ---------- tên loại, trích yếu / kính gửi ----------
    if cong_van:
        doc.add_paragraph()
        nhan = [x.strip() for x in a.kinh_gui.split(';') if x.strip()] or ['Tên cơ quan nhận']
        if len(nhan) == 1:
            p = doc.add_paragraph(f'Kính gửi: {nhan[0]}.')
            T.fmt(p, 'kinh_gui', align=C)
        else:
            p = doc.add_paragraph('Kính gửi:')
            T.fmt(p, 'kinh_gui', align=L)
            p.paragraph_format.left_indent = Cm(4)
            for i, x in enumerate(nhan):
                q = doc.add_paragraph(f'- {x}' + ('.' if i == len(nhan) - 1 else ';'))
                T.fmt(q, 'kinh_gui', align=L)
                q.paragraph_format.left_indent = Cm(5.5)
    else:
        p = doc.add_paragraph(a.loai.upper())
        T.fmt(p, 'ten_loai', align=C, tight=False)
        p.paragraph_format.space_before, p.paragraph_format.space_after = Pt(18), Pt(0)
        p.paragraph_format.keep_with_next = True
        p = doc.add_paragraph(a.trich_yeu)
        T.fmt(p, 'trich_yeu', align=C)
        T.add_rule_after(p, max(20, min(T.text_width_mm(a.trich_yeu, 14, True), TEXT_W) * 0.4), TEXT_W)

    # ---------- nội dung ----------
    for txt in ('Nội dung văn bản. Đoạn văn canh đều hai bên, cỡ chữ 13–14, lùi đầu dòng 1,27 cm, cách đoạn '
                'tối thiểu 6 pt, giãn dòng từ 1 đến 1,5 dòng.',
                'I. TÊN MỤC LỚN (số La Mã, in hoa, đậm)', '1. Tên mục nhỏ (số Ả Rập, đậm)',
                'a) Ý nhỏ (chữ cái thường, dấu ngoặc đơn)', '- Ý chi tiết (gạch đầu dòng)'):
        p = doc.add_paragraph(txt)
        p.paragraph_format.alignment = J
        p.paragraph_format.first_line_indent = Cm(1.27)
        if txt[:2] in ('I.', '1.'):
            p.runs[0].bold = True
    doc.add_paragraph()

    # ---------- nơi nhận + chữ ký ----------
    sg = no_border_table(doc, [80, TEXT_W - 80])
    lc, rc = sg.rows[0].cells
    cell_par(lc, 'Nơi nhận:', 'noi_nhan', first=True, align=L)
    items = [x.strip() for x in a.noi_nhan.split(';') if x.strip()] + [f'Lưu: {a.luu}']
    for i, x in enumerate(items):
        cell_par(lc, f'- {x}' + ('.' if i == len(items) - 1 else ';'), 'noi_nhan_ds', align=L)
    qh = f'{a.quyen_han}. ' if a.quyen_han else ''
    if a.quyen_han and a.cap_tren:
        cell_par(rc, f'{qh}{a.cap_tren.upper()}', 'chuc_vu', first=True)
        cell_par(rc, a.chuc_vu.upper(), 'chuc_vu')
    else:
        cell_par(rc, f'{qh}{a.chuc_vu.upper()}', 'chuc_vu', first=True)
    for _ in range(4):
        blank(rc, 14)
    cell_par(rc, a.nguoi_ky, 'ho_ten')

    doc.save(a.out)
    print(f'Đã tạo: {a.out}')


if __name__ == '__main__':
    main()
