#!/usr/bin/env python3
"""Kiểm tra thể thức file .docx theo Nghị định 30/2020 và quy ước báo cáo khoa học.

Chỉ đọc, không sửa file. In báo cáo các lỗi: trang giấy, phông/cỡ chữ, tiêu đề chưa gán style,
đánh số chỉ mục trùng/thiếu, chú thích bảng/biểu đồ/hình trùng/thiếu/không thống nhất, mục lục gõ tay,
dấu thập phân không thống nhất.

    python audit_docx.py file.docx [--json]
"""
import argparse, json, re, sys
from collections import Counter, defaultdict
import docx
from docx.oxml.ns import qn
from docx.shared import Emu
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import the_thuc  # noqa: E402
import van_phong  # noqa: E402

HEADING_RE = re.compile(
    r'^(?P<num>(?:Chương|CHƯƠNG|Phần|PHẦN)\s+[0-9IVXLC]+|\d+(?:\.\d+)*|[IVXLC]+|[a-zđ])'
    r'(?P<sep>[.:)]?)\s+\S')
CAPTION_RE = re.compile(r'^(?P<kind>Bảng|Biểu đồ|Biều đồ|Hình|Sơ đồ|Ảnh)\s+(?:(?P<ch>\d+)\.)?(?P<no>\d+)(?!\.\d)\s*(?P<sep>[.:\-–]?)')
ROMAN = {'I': 1, 'II': 2, 'III': 3, 'IV': 4, 'V': 5, 'VI': 6, 'VII': 7, 'VIII': 8, 'IX': 9, 'X': 10,
         'XI': 11, 'XII': 12}
UNNUMBERED = ('ĐẶT VẤN ĐỀ', 'MỞ ĐẦU', 'KẾT LUẬN', 'KIẾN NGHỊ', 'TÀI LIỆU THAM KHẢO', 'PHỤ LỤC',
              'LỜI CAM ĐOAN', 'LỜI CẢM ƠN', 'MỤC LỤC', 'DANH MỤC')
TOC_TITLES = ('MỤC LỤC', 'DANH MỤC')


def mm(v):
    return None if v is None else round(Emu(v).mm, 1)


def is_bold(p):
    runs = [r for r in p.runs if r.text.strip()]
    if not runs:
        return False
    if all(r.bold for r in runs):
        return True
    st = p.style
    while st is not None:
        if st.font.bold:
            return all(r.bold is not False for r in runs)
        st = st.base_style
    return False


def outline_level(p):
    st = p.style
    for el in (p._p.pPr, *([s.element.pPr for s in iter_styles(st)])):
        if el is not None:
            o = el.find(qn('w:outlineLvl'))
            if o is not None:
                return int(o.get(qn('w:val')))
    return None


def iter_styles(st):
    while st is not None:
        yield st
        st = st.base_style


DEFAULT_FONT = [None]


def default_font(d):
    """Phông của docDefaults; nếu trỏ tới theme thì đọc tên phông trong theme1.xml."""
    if DEFAULT_FONT[0] is None:
        name = None
        rf = d.styles.element.find(qn('w:docDefaults') + '/' + qn('w:rPrDefault') + '/' + qn('w:rPr')
                                   + '/' + qn('w:rFonts'))
        if rf is not None:
            name = rf.get(qn('w:ascii'))
            theme = rf.get(qn('w:asciiTheme'))
            if not name and theme:
                for rel in d.part.rels.values():
                    if rel.reltype.endswith('/theme'):
                        x = rel.target_part.blob.decode('utf8', 'ignore')
                        tag = 'majorFont' if theme.startswith('major') else 'minorFont'
                        m = re.search(tag + r'>\s*<a:latin typeface="([^"]+)"', x)
                        name = m and m.group(1) + ' (theme)'
        DEFAULT_FONT[0] = name or '(không xác định)'
    return DEFAULT_FONT[0]


def eff_font(p, run):
    """Phông, cỡ chữ hiệu lực của run (run -> style đoạn -> Normal -> docDefaults)."""
    name, size = run.font.name, run.font.size
    for st in iter_styles(p.style):
        name = name or st.font.name
        size = size or st.font.size
    return name or default_font(p.part.document), (size.pt if size else None)


def classify_heading(p):
    """Trả về (level, label) nếu đoạn trông như tiêu đề, ngược lại None."""
    t = p.text.strip()
    if not t or len(t) > 200:
        return None
    up = t.upper()
    if any(up.startswith(u) for u in UNNUMBERED) and (is_bold(p) or outline_level(p) is not None):
        return (1, t)
    m = HEADING_RE.match(t)
    if not m:
        return None
    num = m.group('num')
    if not (is_bold(p) or outline_level(p) is not None):
        return None
    if re.match(r'(?i)(chương|phần)', num):
        return (1, num)
    if re.fullmatch(r'\d+(?:\.\d+)+', num):
        return (num.count('.') + 1, num)
    if re.fullmatch(r'[IVXLC]+', num) and m.group('sep') == '.':
        return (1, num)
    if re.fullmatch(r'\d+', num) and m.group('sep') == '.' and len(t) < 120 and is_bold(p):
        return (2, num)          # báo cáo hành chính: I. → 1. → a)
    if re.fullmatch(r'[a-zđ]', num) and m.group('sep') in ('.', ')'):
        return (9, num)          # điểm a/b/c, cấp tuỳ ngữ cảnh
    return None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('file')
    ap.add_argument('--json', action='store_true')
    ap.add_argument('--no-van-phong', action='store_true', help='bỏ qua kiểm tra văn phong')
    a = ap.parse_args()
    d = docx.Document(a.file)
    ps = d.paragraphs
    R = defaultdict(list)

    # ---- 1. Trang giấy ----
    for i, s in enumerate(d.sections):
        w, h = mm(s.page_width), mm(s.page_height)
        size_ok = {w, h} == {210.0, 297.0} or (abs(min(w, h) - 210) < 1 and abs(max(w, h) - 297) < 1)
        margins = dict(tren=mm(s.top_margin), duoi=mm(s.bottom_margin),
                       trai=mm(s.left_margin), phai=mm(s.right_margin))
        msg = []
        if not size_ok:
            msg.append(f'khổ {w}×{h} mm (cần A4 210×297)')
        if not 20 <= margins['tren'] <= 25: msg.append(f"lề trên {margins['tren']} (20–25)")
        if not 20 <= margins['duoi'] <= 25: msg.append(f"lề dưới {margins['duoi']} (20–25)")
        if not 30 <= margins['trai'] <= 35: msg.append(f"lề trái {margins['trai']} (30–35)")
        if not 15 <= margins['phai'] <= 20: msg.append(f"lề phải {margins['phai']} (15–20)")
        if msg:
            R['trang_giay'].append(f'section {i + 1}: ' + '; '.join(msg))

    # ---- 2. Phông, cỡ chữ ----
    fonts, sizes = Counter(), Counter()
    for p in ps:
        if p.style.name.lower().startswith(('toc', 'table of')):
            continue
        for r in p.runs:
            if not r.text.strip():
                continue
            n, s = eff_font(p, r)
            fonts[n] += len(r.text)
            if p.style.name in ('Normal', 'Body Text', 'List Paragraph', 'Normal (Web)'):
                sizes[s] += len(r.text)
    for n, c in fonts.most_common():
        if n != 'Times New Roman':
            R['phong_chu'].append(f'phông "{n}": {c} ký tự (cần Times New Roman)')
    for s, c in sizes.most_common():
        if s is not None and not 13 <= s <= 14:
            R['co_chu'].append(f'cỡ {s} pt trong đoạn nội dung: {c} ký tự (cần 13–14)')

    # ---- 3. Tiêu đề và đánh số ----
    heads = []
    for i, p in enumerate(ps):
        if any(p.text.strip().upper().startswith(t) for t in TOC_TITLES) and i < len(ps) * 0.15:
            pass
        if p.style.name.lower().startswith(('toc', 'table of')):
            continue
        c = classify_heading(p)
        if c and not CAPTION_RE.match(p.text.strip()):
            heads.append((i, c[0], c[1], p.style.name, p.text.strip()))
    style_by_level = defaultdict(Counter)
    heads_num = [h for h in heads if h[2] != h[4]]   # bỏ tiêu đề không đánh số (ĐẶT VẤN ĐỀ, PHỤ LỤC…)
    for i, lv, num, st, t in heads_num:
        style_by_level[lv][st] += 1
    for lv, cnt in sorted(style_by_level.items()):
        if lv == 9:
            continue
        body = ('Normal', 'Normal (Web)', 'No Spacing', 'Body Text', 'List Paragraph')
        ranked = [x for x, _ in cnt.most_common() if x not in body] or [cnt.most_common(1)[0][0]]
        main_style = ranked[0]
        for i, l2, num, st, t in heads_num:
            if l2 == lv and st != main_style:
                R['tieu_de_lech_style'].append(
                    f'[{i}] cấp {lv} dùng style "{st}" (đa số dùng "{main_style}"): {t[:70]}')
    for i, lv, num, st, t in heads:
        if st in ('Normal', 'Normal (Web)', 'No Spacing', 'Body Text') and outline_level(ps[i]) is None \
                and t.upper() != 'MỤC LỤC':
            R['tieu_de_chua_gan_style'].append(f'[{i}] {t[:80]}')
    # trùng / nhảy số
    seen = {}
    last_child = {}
    for i, lv, num, st, t in heads:
        if not re.fullmatch(r'\d+(?:\.\d+)+', num):
            continue
        if num in seen:
            R['so_trung'].append(f'{num} lặp lại: [{seen[num][0]}] {seen[num][1][:50]} | [{i}] {t[:50]}')
        seen[num] = (i, t)
        parts = list(map(int, num.split('.')))
        parent, k = tuple(parts[:-1]), parts[-1]
        prev = last_child.get(parent, 0)
        if k > prev + 1:
            R['so_nhay_coc'].append(f'{num}: trước đó là {".".join(map(str, parent))}.{prev} — [{i}] {t[:60]}')
        last_child[parent] = max(prev, k)
    # kiểu tên chương
    chap_styles = Counter()
    for i, lv, num, st, t in heads:
        m = re.match(r'(?i)chương\s+([0-9IVXLC]+)\s*([.:]?)', t)
        if m:
            kind = 'La Mã' if m.group(1).upper() in ROMAN else 'Ả Rập'
            chap_styles[(kind, m.group(2) or '(không dấu)')] += 1
            R['_chuong'].append(f'[{i}] {t[:60]}')
    if len(chap_styles) > 1:
        R['ten_chuong_khong_thong_nhat'] = [f'{k[0]} + dấu "{k[1]}": {v}' for k, v in chap_styles.items()] \
            + R['_chuong']
    R.pop('_chuong', None)
    for i, lv, num, st, t in heads:
        if re.fullmatch(r'\d+(?:\.\d+)+', num):
            m = HEADING_RE.match(t)
            if m.group('sep') != '.':
                R['so_muc_thieu_dau_cham'].append(f'[{i}] {t[:70]}')

    # ---- 4. Chú thích bảng / biểu đồ / hình ----
    caps = defaultdict(list)
    for i, p in enumerate(ps):
        if p.style.name.lower().startswith(('toc', 'table of')):
            continue
        m = CAPTION_RE.match(p.text.strip())
        if m and (is_bold(p) or len(p.text) < 250):
            kind = 'Biểu đồ' if m.group('kind') == 'Biều đồ' else m.group('kind')
            if m.group('kind') == 'Biều đồ':
                R['chinh_ta_chu_thich'].append(f'[{i}] "Biều đồ" → "Biểu đồ"')
            caps[kind].append((i, int(m.group('ch') or 0), int(m.group('no')), m.group('sep'), p.style.name,
                               p.text.strip()))
    for kind, items in caps.items():
        st = Counter(x[4] for x in items)
        main = st.most_common(1)[0][0]
        for i, ch, no, sep, s, t in items:
            if s != main:
                R['chu_thich_lech_style'].append(f'[{i}] {kind} {ch}.{no} style "{s}" (đa số "{main}")')
        seps = Counter(x[3] or '(không)' for x in items)
        if len(seps) > 1:
            R['chu_thich_dau_khong_thong_nhat'].append(
                f'{kind}: ' + ', '.join(f'"{k}"×{v}' for k, v in seps.items()))
        by_ch = defaultdict(list)
        for x in items:
            by_ch[x[1]].append(x)
        for ch, xs in sorted(by_ch.items()):
            nos = [x[2] for x in xs]
            for n_, c in Counter(nos).items():
                if c > 1:
                    R['chu_thich_trung'].append(f'{kind} {ch}.{n_} xuất hiện {c} lần'.replace(' 0.', ' '))
            missing = sorted(set(range(1, max(nos) + 1)) - set(nos))
            if missing:
                R['chu_thich_thieu'].append(f'{kind}' + (f' chương {ch}' if ch else '') + ': thiếu ' +
                                            ', '.join(f'{ch}.{m}' if ch else str(m) for m in missing))
            if nos != sorted(nos):
                R['chu_thich_sai_thu_tu'].append(f'{kind} chương {ch}: thứ tự {nos}')
        # chú thích bị ngắt dòng/ngắt đoạn
        for i, ch, no, sep, s, t in items:
            if ps[i]._p.findall('.//' + qn('w:br')):
                R['chu_thich_ngat_dong'].append(f'[{i}] {kind} {ch}.{no} có ngắt dòng tay (Shift+Enter)')
            nxt = ps[i + 1] if i + 1 < len(ps) else None
            if nxt is not None and nxt.text.strip() and nxt.style.name == s \
                    and len(nxt.text.strip()) < 120 and not CAPTION_RE.match(nxt.text.strip()):
                R['chu_thich_ngat_dong'].append(f'[{i}] {kind} {ch}.{no} bị tách 2 đoạn: "{nxt.text.strip()[:40]}"')

    # ---- 5. Mục lục / danh mục ----
    instr = [x.text or '' for x in d.element.body.iter(qn('w:instrText'))]
    toc_fields = [x.strip() for x in instr if x.strip().startswith('TOC')]
    long_doc = len(ps) > 120 or sum(1 for h in heads if h[1] == 1) >= 3
    if not toc_fields and long_doc:
        R['muc_luc'].append('không có trường TOC tự động: mục lục/danh mục đang gõ tay hoặc chưa có')
    if not long_doc:
        R.pop('tieu_de_chua_gan_style', None)   # văn bản ngắn không cần mục lục
    else:
        R['_info_toc'] = toc_fields
    errs = [i for i, p in enumerate(ps) if 'Error! Bookmark' in p.text or 'Lỗi! Không tìm thấy' in p.text]
    if errs:
        R['muc_luc'].append(f'{len(errs)} dòng lỗi "Error! Bookmark not defined" tại {errs[:10]}')
    for kind in (caps if long_doc else []):
        if not any(kind.lower() in x.lower() or 'table of figures' in x.lower() or '\\c' in x
                   for x in toc_fields):
            R['muc_luc'].append(f'có {len(caps[kind])} chú thích "{kind}" nhưng chưa thấy trường danh mục cho loại này '
                                '(kiểm tra field TOC \\t "<style>")')

    # ---- 6. Số liệu ----
    dec_dot = dec_comma = 0
    ex_dot, ex_comma = [], []
    num_dot = re.compile(r'(?<![\d.])\d+\.\d{1,2}(?![\d.])')
    num_comma = re.compile(r'(?<![\d,])\d+,\d{1,2}(?![\d,])')
    for i, p in enumerate(ps):
        t = p.text
        if classify_heading(p) or CAPTION_RE.match(t.strip()) or p.style.name.lower().startswith(('toc', 'endnote')):
            continue
        for m in num_dot.finditer(t):
            dec_dot += 1
            len(ex_dot) < 3 and ex_dot.append(f'[{i}] …{t[max(0, m.start() - 15):m.end() + 10]}…')
        for m in num_comma.finditer(t):
            dec_comma += 1
            len(ex_comma) < 3 and ex_comma.append(f'[{i}] …{t[max(0, m.start() - 15):m.end() + 10]}…')
    if dec_dot and dec_comma:
        R['so_thap_phan_khong_thong_nhat'] = [f'dấu chấm: {dec_dot} lần, dấu phẩy: {dec_comma} lần '
                                              '(lưu ý: dấu phẩy có thể là phân cách liệt kê, cần xem lại)'] + ex_dot + ex_comma
    elif dec_dot:
        R['so_thap_phan_dau_cham'] = [f'{dec_dot} số dùng dấu chấm thập phân (tiếng Việt thường dùng dấu phẩy)'] + ex_dot
    pm = sum(len(re.findall(r'\S±|±\S', p.text)) for p in ps)
    if pm:
        R['dau_cong_tru'].append(f'{pm} chỗ "±" thiếu khoảng trắng hai bên')
    for p_i, p in enumerate(ps):
        if re.search(r'[ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ]', p.text):
            R['so_la_ma_unicode'].append(f'[{p_i}] dùng ký tự La Mã Unicode, nên dùng chữ I, V, X')

    # ---- 7. Thể thức hành chính, ảnh, bảng, số trang ----
    the_thuc.audit_layout(d, R)
    for x in toc_fields:
        lv = [int(n) for n in re.findall(r'[,;](\d)(?=[,;"])', x)] + \
             [int(n) for n in re.findall(r'\\o\s+"\d-(\d)"', x)]
        if lv and max(lv) > 3:
            R['muc_luc'].append(f'mục lục đang lấy tới cấp {max(lv)}: chỉ nên tới cấp 3 (1.1.1.), '
                                'không đưa a), b), c) vào mục lục (fix_docx.py --toc)')

    # ---- 8. Văn phong kỹ thuật ----
    if not a.no_van_phong:
        van_phong.audit(d, R)

    # ---- Xuất ----
    R = {k: v for k, v in R.items() if v}
    if a.json:
        json.dump(R, sys.stdout, ensure_ascii=False, indent=1)
        return
    labels = {
        'trang_giay': 'Trang giấy (khổ, lề)', 'phong_chu': 'Phông chữ', 'co_chu': 'Cỡ chữ nội dung',
        'tieu_de_chua_gan_style': 'Tiêu đề chưa gán style (sẽ không lên mục lục)',
        'tieu_de_lech_style': 'Tiêu đề cùng cấp dùng style khác nhau',
        'so_trung': 'Số chỉ mục trùng', 'so_nhay_coc': 'Số chỉ mục nhảy cóc',
        'so_muc_thieu_dau_cham': 'Số chỉ mục thiếu dấu chấm cuối',
        'ten_chuong_khong_thong_nhat': 'Cách ghi tên chương không thống nhất',
        'chinh_ta_chu_thich': 'Chính tả chú thích', 'chu_thich_lech_style': 'Chú thích lệch style',
        'chu_thich_dau_khong_thong_nhat': 'Dấu sau số chú thích không thống nhất',
        'chu_thich_trung': 'Chú thích trùng số', 'chu_thich_thieu': 'Chú thích thiếu số',
        'chu_thich_sai_thu_tu': 'Chú thích sai thứ tự', 'chu_thich_ngat_dong': 'Chú thích bị ngắt dòng/đoạn',
        'muc_luc': 'Mục lục / danh mục', '_info_toc': '(thông tin) Trường TOC đang có',
        'so_thap_phan_khong_thong_nhat': 'Dấu thập phân không thống nhất',
        'so_thap_phan_dau_cham': 'Dấu thập phân', 'dau_cong_tru': 'Dấu ±',
        'so_la_ma_unicode': 'Số La Mã',
        'the_thuc_thieu': 'Thiếu thành phần thể thức (NĐ 30)', 'the_thuc_dinh_dang': 'Cỡ/kiểu chữ thành phần thể thức',
        'the_thuc_noi_dung': 'Cách ghi thành phần thể thức', 'anh': 'Ảnh', 'bang': 'Bảng',
        'vp_khau_ngu': 'Văn phong: khẩu ngữ, từ không trang trọng', 'vp_chu_quan': 'Văn phong: chủ quan, cảm tính',
        'vp_cam_than': 'Văn phong: dấu cảm thán, biểu tượng', 'vp_cau_dai': 'Văn phong: câu quá dài (> 70 âm tiết)',
        'vp_doan_dai': 'Văn phong: đoạn quá dài', 'vp_lap_tu': 'Văn phong: lặp từ',
        'vp_dau_cau': 'Văn phong: dấu câu, khoảng trắng', 'vp_don_vi': 'Văn phong: số và đơn vị',
        'vp_viet_tat': 'Văn phong: chữ viết tắt', 'vp_chinh_ta': 'Văn phong: chính tả thống nhất',
        'vp_ngoi': 'Văn phong: ngôi xưng',
    }
    print(f'== KIỂM TRA THỂ THỨC: {a.file}')
    print(f'   {len(ps)} đoạn, {len(heads)} tiêu đề, ' +
          ', '.join(f'{len(v)} {k.lower()}' for k, v in caps.items()))
    if not R:
        print('Không phát hiện lỗi.')
    for k, v in R.items():
        print(f'\n## {labels.get(k, k)} ({len(v)})')
        for x in v[:40]:
            print('  -', x)
        if len(v) > 40:
            print(f'  … và {len(v) - 40} dòng nữa')


if __name__ == '__main__':
    main()
