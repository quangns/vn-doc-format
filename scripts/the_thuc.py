"""Thể thức văn bản hành chính (Nghị định 30/2020, Phụ lục I), ảnh, bảng, định dạng tiêu đề.

Dùng chung cho fix_docx.py (sửa), audit_docx.py (kiểm tra), tao_mau.py (tạo văn bản mẫu).
Muốn đổi cỡ chữ mặc định thì sửa SIZES (giá trị phải nằm trong khoảng NĐ 30 ghi ở chú thích).
"""
import re
import docx
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, Emu, Cm
from docx.table import Table
from docx.text.paragraph import Paragraph

# ------------------------------------------------------------------ quy chuẩn
#        cỡ  đậm   nghiêng in-hoa  (khoảng NĐ 30)
SIZES = {
    'co_quan_chu_quan': (13, False, False, True),   # 12–13, đứng, in hoa
    'co_quan':          (13, True,  False, True),   # 12–13, đứng, đậm, in hoa
    'quoc_hieu':        (12, True,  False, True),   # 12–13, đứng, đậm, in hoa
    'tieu_ngu':         (13, True,  False, False),  # 13–14, đứng, đậm
    'so_ky_hieu':       (13, False, False, False),  # 13, đứng
    'dia_danh':         (14, False, True,  False),  # 13–14, nghiêng
    'ten_loai':         (14, True,  False, True),   # 13–14, đứng, đậm, in hoa
    'trich_yeu':        (14, True,  False, False),  # 13–14, đứng, đậm
    'vv':               (12, False, False, False),  # 12–13, đứng (công văn)
    'kinh_gui':         (14, False, False, False),  # 14, đứng
    'noi_nhan':         (12, True,  True,  False),  # 12, nghiêng, đậm
    'noi_nhan_ds':      (11, False, False, False),  # 11, đứng
    'chuc_vu':          (14, True,  False, True),   # 13–14, đứng, đậm, in hoa
    'ho_ten':           (14, True,  False, False),  # 13–14, đứng, đậm
    'so_trang':         (13, False, False, False),  # 13–14, đứng
}
RANGES = {'co_quan_chu_quan': (12, 13), 'co_quan': (12, 13), 'quoc_hieu': (12, 13), 'tieu_ngu': (13, 14),
          'so_ky_hieu': (13, 13), 'dia_danh': (13, 14), 'ten_loai': (13, 14), 'trich_yeu': (13, 14),
          'vv': (12, 13), 'kinh_gui': (13, 14), 'noi_nhan': (12, 12), 'noi_nhan_ds': (11, 11),
          'chuc_vu': (13, 14), 'ho_ten': (13, 14)}
LABEL = {'co_quan_chu_quan': 'Tên cơ quan chủ quản', 'co_quan': 'Tên cơ quan ban hành', 'quoc_hieu': 'Quốc hiệu',
         'tieu_ngu': 'Tiêu ngữ', 'so_ky_hieu': 'Số, ký hiệu', 'dia_danh': 'Địa danh, ngày tháng',
         'ten_loai': 'Tên loại văn bản', 'trich_yeu': 'Trích yếu', 'vv': 'Trích yếu công văn (V/v)',
         'kinh_gui': 'Kính gửi', 'noi_nhan': '"Nơi nhận:"', 'noi_nhan_ds': 'Danh sách nơi nhận',
         'chuc_vu': 'Quyền hạn, chức vụ người ký', 'ho_ten': 'Họ tên người ký'}

QUOC_HIEU = 'CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM'
TIEU_NGU = 'Độc lập - Tự do - Hạnh phúc'
QH_RE = re.compile(r'^\s*CỘNG\s+H[OÒ]À?A?\s+XÃ\s+HỘI\s+CHỦ\s+NGHĨA\s+VIỆT\s+NAM\s*$', re.I)
TN_RE = re.compile(r'^\s*Độc\s+lập\s*[-–—]\s*Tự\s+do\s*[-–—]\s*Hạnh\s+phúc\s*$', re.I)
SO_RE = re.compile(r'^\s*Số\s*:?\s*(\d+|…+|\.+)?\s*/\s*([\w\-–.…]+)', re.I)
NGAY_RE = re.compile(r'^\s*[^\d,]{2,40},\s*ngày\s*(\d{1,2}|…+|\.+)?\s+tháng\s*(\d{1,2}|…+|\.+)?\s+năm\s*(\d{4}|…+|\.+)?')
LOAI = ('BÁO CÁO', 'QUYẾT ĐỊNH', 'TỜ TRÌNH', 'KẾ HOẠCH', 'THÔNG BÁO', 'CHỈ THỊ', 'QUY CHẾ', 'QUY ĐỊNH',
        'ĐỀ ÁN', 'PHƯƠNG ÁN', 'CHƯƠNG TRÌNH', 'HƯỚNG DẪN', 'BIÊN BẢN', 'NGHỊ QUYẾT', 'HỢP ĐỒNG', 'GIẤY MỜI',
        'GIẤY GIỚI THIỆU', 'GIẤY ỦY QUYỀN', 'GIẤY NGHỈ PHÉP', 'CÔNG ĐIỆN', 'BẢN GHI NHỚ', 'BẢN THỎA THUẬN',
        'PHIẾU GỬI', 'PHIẾU CHUYỂN', 'PHIẾU BÁO', 'THƯ CÔNG', 'ĐỀ CƯƠNG', 'DỰ ÁN', 'TỜ KHAI', 'ĐƠN')
QUYEN_HAN_RE = re.compile(r'^(TM|KT|Q|TL|TUQ)\.\s*', re.I)
BODY_START_RE = re.compile(r'^(Căn cứ|Kính gửi|Thực hiện|Theo|Xét|I\.|1\.|Phần|PHẦN|Chương|CHƯƠNG)')

TW_PER_MM = 1440 / 25.4

log = []


# ------------------------------------------------------------------ tiện ích XML
_ORDER = {
    'pPr': ['pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'framePr', 'widowControl', 'numPr',
            'suppressLineNumbers', 'pBdr', 'shd', 'tabs', 'suppressAutoHyphens', 'kinsoku', 'wordWrap',
            'overflowPunct', 'topLinePunct', 'autoSpaceDE', 'autoSpaceDN', 'bidi', 'adjustRightInd',
            'snapToGrid', 'spacing', 'ind', 'contextualSpacing', 'mirrorIndents', 'suppressOverlap', 'jc',
            'textDirection', 'textAlignment', 'textboxTightWrap', 'outlineLvl', 'divId', 'cnfStyle', 'rPr',
            'sectPr', 'pPrChange'],
    'tblPr': ['tblStyle', 'tblpPr', 'tblOverlap', 'bidiVisual', 'tblStyleRowBandSize', 'tblStyleColBandSize',
              'tblW', 'jc', 'tblCellSpacing', 'tblInd', 'tblBorders', 'shd', 'tblLayout', 'tblCellMar',
              'tblLook', 'tblCaption', 'tblDescription'],
    'trPr': ['cnfStyle', 'divId', 'gridBefore', 'gridAfter', 'wBefore', 'wAfter', 'cantSplit', 'trHeight',
             'tblHeader', 'tblCellSpacing', 'jc', 'hidden'],
}


def ensure(parent, tag, kind):
    """Lấy hoặc tạo phần tử con `w:tag` đúng thứ tự schema của `kind` (pPr/tblPr/trPr)."""
    el = parent.find(qn('w:' + tag))
    if el is not None:
        return el
    el = OxmlElement('w:' + tag)
    order = _ORDER[kind]
    after = set(order[order.index(tag) + 1:])
    for c in parent:
        if c.tag.split('}')[-1] in after:
            c.addprevious(el)
            return el
    parent.append(el)
    return el


def replace_span(p, start, end, new):
    """Thay ký tự [start, end) của nội dung các run, giữ định dạng run đầu tiên."""
    pos, first = 0, True
    for r in p.runs:
        t = r.text
        a, b = pos, pos + len(t)
        pos = b
        if b <= start or a >= end:
            continue
        s0, s1 = max(start, a) - a, min(end, b) - a
        r.text = t[:s0] + (new if first else '') + t[s1:]
        first = False


def set_text(p, text):
    """Thay toàn bộ chữ của đoạn, giữ định dạng run đầu tiên."""
    runs = [r for r in p.runs]
    if not runs:
        p.add_run(text)
        return
    runs[0].text = text
    for r in runs[1:]:
        r._r.getparent().remove(r._r)


def run_text(p):
    return ''.join(r.text for r in p.runs)


def fmt(p, key=None, size=None, bold=None, italic=None, upper=False, align=None, indent0=True, tight=True):
    """Định dạng đoạn: cỡ, đậm, nghiêng, in hoa (đổi chữ), canh lề, bỏ thụt đầu dòng, giãn dòng đơn."""
    if key:
        size, bold, italic, upper = SIZES[key]
    for r in p.runs:
        if not r.text:
            continue
        if size is not None:
            r.font.size = Pt(size)
        if bold is not None:
            r.font.bold = bold
        if italic is not None:
            r.font.italic = italic
        if upper:
            r.text = r.text.upper()
    # cỡ dấu đoạn cũng chỉnh để chiều cao dòng trống/cuối đúng
    if size is not None:
        rPr = ensure(p._p.get_or_add_pPr(), 'rPr', 'pPr')
        for t in ('w:sz', 'w:szCs'):
            e = rPr.find(qn(t))
            if e is None:
                e = OxmlElement(t)
                rPr.append(e)
            e.set(qn('w:val'), str(int(size * 2)))
    pf = p.paragraph_format
    if align is not None:
        pf.alignment = align
    if indent0:
        pf.first_line_indent = Cm(0)
        pf.left_indent = Cm(0)
        pf.right_indent = Cm(0)
    if tight:
        pf.space_before = Pt(0)
        pf.space_after = Pt(0)
        pf.line_spacing = 1.0


def text_width_mm(text, size_pt, bold=False):
    """Ước lượng bề rộng dòng chữ Times New Roman (đủ dùng để vẽ đường kẻ)."""
    k = 0.52 if bold else 0.47
    upper = sum(1 for c in text if c.isupper())
    em = (len(text) - upper) * k + upper * (k + 0.17)
    return em * size_pt * 25.4 / 72


def container_width_mm(doc, p):
    """Bề rộng vùng chứa đoạn: ô bảng (nếu nằm trong bảng) hoặc vùng chữ của trang."""
    tc = p._p.getparent()
    if tc.tag == qn('w:tc'):
        w = tc.find(qn('w:tcPr') + '/' + qn('w:tcW'))
        if w is not None and w.get(qn('w:type')) in (None, 'dxa') and int(w.get(qn('w:w'), 0)) > 0:
            return int(w.get(qn('w:w'))) / TW_PER_MM - 3.8   # trừ lề trong ô ~1,9 mm mỗi bên
    s = doc.sections[0]
    return Emu(s.page_width - s.left_margin - s.right_margin).mm


def has_rule_near(p, ahead=2):
    """Đã có đường kẻ (shape, đường viền đoạn, gạch chân) ngay dưới đoạn chưa?"""
    el, k = p._p, 0
    while el is not None and k <= ahead:
        if el.tag == qn('w:p'):
            if el.findall('.//' + qn('w:pict')) or el.findall('.//' + qn('w:drawing')) \
                    or el.find(qn('w:pPr') + '/' + qn('w:pBdr')) is not None:
                return True
            if k == 0 and all(r.font.underline for r in Paragraph(el, p._parent).runs if r.text.strip()) \
                    and run_text(p).strip():
                return True
            if k > 0 and Paragraph(el, p._parent).text.strip():
                return False
        el, k = el.getnext(), k + 1
    return False


def add_rule_after(p, width_mm, container_mm):
    """Chèn đường kẻ ngang nét liền, canh giữa, dài `width_mm` ngay dưới đoạn `p`."""
    side = max(0.0, (container_mm - width_mm) / 2)
    np_ = OxmlElement('w:p')
    pPr = OxmlElement('w:pPr')
    np_.append(pPr)
    bdr = ensure(pPr, 'pBdr', 'pPr')
    b = OxmlElement('w:bottom')
    for k, v in (('w:val', 'single'), ('w:sz', '6'), ('w:space', '1'), ('w:color', 'auto')):
        b.set(qn(k), v)
    bdr.append(b)
    sp = ensure(pPr, 'spacing', 'pPr')
    for k, v in (('w:before', '0'), ('w:after', '0'), ('w:line', '240'), ('w:lineRule', 'auto')):
        sp.set(qn(k), v)
    ind = ensure(pPr, 'ind', 'pPr')
    for k in ('w:left', 'w:right'):
        ind.set(qn(k), str(int(side * TW_PER_MM)))
    ind.set(qn('w:firstLine'), '0')
    ensure(pPr, 'jc', 'pPr').set(qn('w:val'), 'center')
    rPr = ensure(pPr, 'rPr', 'pPr')
    for t in ('w:sz', 'w:szCs'):
        e = OxmlElement(t)
        e.set(qn('w:val'), '4')       # dòng kẻ thấp 2 pt
        rPr.append(e)
    p._p.addnext(np_)
    return Paragraph(np_, p._parent)


def iter_paragraphs(doc, include_tables=True):
    """Đoạn văn theo đúng thứ tự trong tài liệu, kể cả trong bảng (và bảng lồng)."""
    def walk(parent_el, parent_obj):
        for el in parent_el.iterchildren():
            if el.tag == qn('w:p'):
                yield Paragraph(el, parent_obj)
            elif el.tag == qn('w:tbl') and include_tables:
                t = Table(el, parent_obj)
                for row in t.rows:
                    seen = set()
                    for c in row.cells:
                        if id(c._tc) in seen:
                            continue
                        seen.add(id(c._tc))
                        yield from walk(c._tc, c)
            elif el.tag == qn('w:sdt'):
                content = el.find(qn('w:sdtContent'))
                if content is not None:
                    yield from walk(content, parent_obj)
    yield from walk(doc.element.body, doc._body)


def in_table(p):
    return p._p.getparent().tag == qn('w:tc')


def eff(p, r, attr):
    """Giá trị hiệu lực bold/italic/size của run (run → style ký tự → style đoạn → docDefaults)."""
    v = getattr(r.font, attr)
    if v is not None:
        return v.pt if attr == 'size' else v
    if r.style is not None and getattr(r.style.font, attr, None) is not None:
        v = getattr(r.style.font, attr)
        return v.pt if attr == 'size' else v
    st = p.style
    while st is not None:
        v = getattr(st.font, attr)
        if v is not None:
            return v.pt if attr == 'size' else v
        st = st.base_style
    if attr == 'size':
        sz = p.part.document.styles.element.find(
            qn('w:docDefaults') + '/' + qn('w:rPrDefault') + '/' + qn('w:rPr') + '/' + qn('w:sz'))
        return int(sz.get(qn('w:val'))) / 2 if sz is not None else 10
    return False


def eff_align(p):
    if p.paragraph_format.alignment is not None:
        return p.paragraph_format.alignment
    st = p.style
    while st is not None:
        if st.paragraph_format.alignment is not None:
            return st.paragraph_format.alignment
        st = st.base_style
    return WD_ALIGN_PARAGRAPH.LEFT


def para_props(p):
    runs = [r for r in p.runs if r.text.strip()]
    if not runs:
        return None
    return dict(size=sorted({eff(p, r, 'size') for r in runs}), bold=all(eff(p, r, 'bold') for r in runs),
                italic=all(eff(p, r, 'italic') for r in runs),
                any_bold=any(eff(p, r, 'bold') for r in runs), any_italic=any(eff(p, r, 'italic') for r in runs))


# ------------------------------------------------------------------ nhận diện thể thức
def detect_layout(doc):
    """Trả về dict: vai trò -> list Paragraph (quốc hiệu, tiêu ngữ, cơ quan, số, ngày, tên loại, ...)."""
    ps = list(iter_paragraphs(doc))
    roles = {}
    add = lambda k, p: roles.setdefault(k, []).append(p)
    head = ps[:60]
    qh = next((p for p in head if QH_RE.match(p.text)), None)
    tn = next((p for p in head if TN_RE.match(p.text)), None)
    if qh is not None:
        add('quoc_hieu', qh)
    if tn is not None:
        add('tieu_ngu', tn)
    so = next((p for p in head if SO_RE.match(p.text)), None)
    if so is not None:
        add('so_ky_hieu', so)
    ngay = next((p for p in head if NGAY_RE.match(p.text)), None)
    if ngay is not None:
        add('dia_danh', ngay)
    # tên cơ quan: các dòng in hoa, không phải quốc hiệu, nằm trước "Số:" trong cùng ô (hoặc đầu văn bản)
    if so is not None:
        cell_ps = [p for p in ps if p._p.getparent() is so._p.getparent()]
        before = cell_ps[:cell_ps.index(so)] if so in cell_ps else []
        lines = [p for p in before if p.text.strip() and p.text.strip() == p.text.strip().upper()
                 and not QH_RE.match(p.text)]
        if len(lines) >= 2:
            add('co_quan_chu_quan', lines[0])
            for p in lines[1:]:
                add('co_quan', p)
        elif lines:
            add('co_quan', lines[0])
        after = cell_ps[cell_ps.index(so) + 1:] if so in cell_ps else []
        vv = next((p for p in after if re.match(r'\s*V/v\b', p.text)), None)
        if vv is not None:
            add('vv', vv)
    if 'vv' not in roles:
        vv = next((p for p in head if re.match(r'\s*V/v\b', p.text)), None)
        if vv is not None:
            add('vv', vv)
    # tên loại + trích yếu (ngoài bảng, phần đầu)
    body_head = [p for p in head if not in_table(p)]
    for i, p in enumerate(body_head):
        t = re.sub(r'\s+', ' ', p.text.strip())
        if t and t.upper() in LOAI or (t and t == t.upper() and len(t) < 60 and any(t.startswith(x + ' ') for x in LOAI)):
            add('ten_loai', p)
            for q in body_head[i + 1:]:
                qt = q.text.strip()
                if not qt:
                    continue
                if not BODY_START_RE.match(qt) and len(qt) < 250:
                    add('trich_yeu', q)
                break
            break
    kg = next((p for p in ps[:80] if re.match(r'\s*Kính\s+gửi\b', p.text)), None)
    if kg is not None:
        add('kinh_gui', kg)
    # nơi nhận + chữ ký
    nn = next((p for p in reversed(ps) if re.match(r'\s*Nơi\s+nhận\s*:?', p.text)), None)
    if nn is not None:
        add('noi_nhan', nn)
        start = ps.index(nn)
        for q in ps[start + 1:]:
            if q._p.getparent() is not nn._p.getparent():
                break
            if re.match(r'\s*[-+•]', q.text) or re.match(r'\s*Lưu\s*:', q.text):
                add('noi_nhan_ds', q)
        tc = nn._p.getparent()
        if tc.tag == qn('w:tc'):
            tr = tc.getparent()
            cells = [c for c in tr.iterchildren(qn('w:tc')) if c is not tc]
            if cells:
                sig = [Paragraph(x, nn._parent) for x in cells[-1].iterchildren(qn('w:p'))]
                filled = [p for p in sig if p.text.strip() and not re.match(r'^\s*[(\[]', p.text)]
                if filled:
                    name = filled[-1]
                    for p in filled[:-1]:      # mọi dòng trên họ tên: quyền hạn, chức vụ
                        add('chuc_vu', p)
                    if not (name.text.strip() == name.text.strip().upper() and len(filled) == 1):
                        add('ho_ten', name)
                    roles['_sig_cell'] = sig
    return roles


# ------------------------------------------------------------------ sửa thể thức hành chính
def pad_date(p):
    """ngày < 10 và tháng 1, 2 thêm số 0 (NĐ 30)."""
    t = run_text(p)
    m = re.search(r'ngày\s+(\d{1,2})\s+tháng\s+(\d{1,2})\s+năm', t)
    if not m:
        return False
    d, mo = int(m.group(1)), int(m.group(2))
    nd = f'{d:02d}'
    nm = f'{mo:02d}' if mo <= 2 else str(mo)
    changed = False
    if m.group(2) != nm:
        replace_span(p, m.start(2), m.end(2), nm)
        changed = True
    if m.group(1) != nd:
        replace_span(p, m.start(1), m.end(1), nd)
        changed = True
    return changed


def pad_so(p):
    t = run_text(p)
    m = re.match(r'(\s*)Số\s*:?\s*(\d+)\s*/\s*', t, re.I)
    if not m:
        return False
    new = f'Số: {int(m.group(2)):02d}/'
    if m.group(0)[len(m.group(1)):] == new:
        return False
    replace_span(p, m.start(0) + len(m.group(1)), m.end(0), new)
    return True


def step_admin(doc):
    roles = detect_layout(doc)
    if not ({'quoc_hieu', 'tieu_ngu', 'noi_nhan', 'so_ky_hieu'} & set(roles)):
        log.append('thể thức hành chính: không thấy quốc hiệu/tiêu ngữ/số ký hiệu/nơi nhận, bỏ qua')
        return
    C = WD_ALIGN_PARAGRAPH.CENTER
    done = []
    for p in roles.get('quoc_hieu', []):
        if p.text.strip() != QUOC_HIEU:
            set_text(p, QUOC_HIEU)
        fmt(p, 'quoc_hieu', align=C)
        done.append('quốc hiệu')
    for p in roles.get('tieu_ngu', []):
        if p.text.strip() != TIEU_NGU:
            set_text(p, TIEU_NGU)
        for r in p.runs:
            r.font.underline = None
        fmt(p, 'tieu_ngu', align=C)
        if not has_rule_near(p):
            add_rule_after(p, text_width_mm(TIEU_NGU, SIZES['tieu_ngu'][0], True), container_width_mm(doc, p))
            done.append('kẻ dưới tiêu ngữ')
        done.append('tiêu ngữ')
    for k in ('co_quan_chu_quan', 'co_quan'):
        for p in roles.get(k, []):
            fmt(p, k, align=C)
            done.append(LABEL[k].lower())
    if roles.get('co_quan'):
        last = roles['co_quan'][-1]
        if not has_rule_near(last):
            w = max(text_width_mm(p.text.strip(), SIZES['co_quan'][0], True) for p in roles['co_quan'])
            add_rule_after(last, max(15, w * 0.4), container_width_mm(doc, last))
            done.append('kẻ dưới tên cơ quan (1/3–1/2 dòng chữ)')
    for p in roles.get('so_ky_hieu', []):
        if pad_so(p):
            done.append('số văn bản thêm 0')
        fmt(p, 'so_ky_hieu', align=C)
    for p in roles.get('dia_danh', []):
        if pad_date(p):
            done.append('ngày/tháng thêm 0')
        fmt(p, 'dia_danh', align=C)
    for p in roles.get('vv', []):
        fmt(p, 'vv', align=C)
    for p in roles.get('ten_loai', []):
        fmt(p, 'ten_loai', align=C, tight=False)
        p.paragraph_format.space_before = Pt(12)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.keep_with_next = True
        done.append('tên loại')
    for p in roles.get('trich_yeu', []):
        fmt(p, 'trich_yeu', align=C)
        p.paragraph_format.keep_with_next = True
        if not has_rule_near(p):
            w = text_width_mm(p.text.strip(), SIZES['trich_yeu'][0], True)
            cw = container_width_mm(doc, p)
            add_rule_after(p, max(20, min(w, cw) * 0.4), cw)
            done.append('kẻ dưới trích yếu')
    for p in roles.get('kinh_gui', []):
        size = SIZES['kinh_gui'][0]
        for r in p.runs:
            r.font.size = Pt(size)
    for p in roles.get('noi_nhan', []):
        t = run_text(p)
        m = re.match(r'\s*Nơi\s+nhận\s*:?', t)
        replace_span(p, m.start(), m.end(), 'Nơi nhận:')
        fmt(p, 'noi_nhan', align=WD_ALIGN_PARAGRAPH.LEFT)
        done.append('nơi nhận')
    ds = roles.get('noi_nhan_ds', [])
    for i, p in enumerate(ds):
        t = run_text(p).rstrip()
        last = i == len(ds) - 1
        end = '.' if last else ';'
        if t and t[-1] in ';.,':
            if t[-1] != end:
                replace_span(p, len(t) - 1, len(t), end)
        elif t:
            p.runs[-1].text = p.runs[-1].text.rstrip() + end
        m = re.match(r'(\s*[-•]?\s*)Lưu\s*:?\s*', run_text(p))
        if m and m.group(0) != '- Lưu: ':
            replace_span(p, m.start(), m.end(), '- Lưu: ')
        t = run_text(p)
        m = re.match(r'\s*[-•]\s*', t)
        if m and m.group(0).strip() == '•':
            replace_span(p, m.start(), m.end(), '- ')
        fmt(p, 'noi_nhan_ds', align=WD_ALIGN_PARAGRAPH.LEFT)
    for p in roles.get('chuc_vu', []):
        fmt(p, 'chuc_vu', align=C)
        done.append('chức vụ người ký')
    for p in roles.get('ho_ten', []):
        fmt(p, 'ho_ten', align=C)
        done.append('họ tên người ký')
    # khoảng ký: tối thiểu 3 dòng trống giữa chức vụ và họ tên
    if roles.get('chuc_vu') and roles.get('ho_ten'):
        cv, ht = roles['chuc_vu'][-1], roles['ho_ten'][0]
        blanks, el = 0, cv._p.getnext()
        while el is not None and el is not ht._p:
            if el.tag == qn('w:p') and not Paragraph(el, None).text.strip():
                blanks += 1
            el = el.getnext()
        for _ in range(max(0, 3 - blanks)):
            ht._p.addprevious(OxmlElement('w:p'))
        if blanks < 3:
            done.append(f'thêm {3 - blanks} dòng trống chỗ ký')
    log.append('thể thức hành chính: ' + ', '.join(dict.fromkeys(done)))
    missing = [LABEL[k] for k in ('quoc_hieu', 'tieu_ngu', 'co_quan', 'so_ky_hieu', 'dia_danh', 'noi_nhan')
               if k not in roles]
    if missing:
        log.append('!! thiếu thành phần thể thức: ' + ', '.join(missing))


# ------------------------------------------------------------------ ảnh
WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'


def text_width_emu(doc):
    s = doc.sections[0]
    return s.page_width - s.left_margin - s.right_margin


def step_images(doc):
    maxw = text_width_emu(doc)
    n_center = n_scale = n_float = 0
    for p in doc.paragraphs:
        inl = p._p.findall('.//{%s}inline' % WP)
        anc = p._p.findall('.//{%s}anchor' % WP)
        if anc:
            n_float += len(anc)
        if not inl:
            continue
        for d in inl:
            ext = d.find('{%s}extent' % WP)
            cx, cy = int(ext.get('cx')), int(ext.get('cy'))
            if cx > maxw:
                k = maxw / cx
                ext.set('cx', str(int(cx * k)))
                ext.set('cy', str(int(cy * k)))
                for e in d.iter('{%s}ext' % A):
                    if e.get('cx') and e.getparent().tag.endswith('}xfrm'):
                        e.set('cx', str(int(int(e.get('cx')) * k)))
                        e.set('cy', str(int(int(e.get('cy')) * k)))
                n_scale += 1
        if len(p.text.strip()) == 0:
            pf = p.paragraph_format
            if pf.alignment != WD_ALIGN_PARAGRAPH.CENTER:
                n_center += 1
            pf.alignment = WD_ALIGN_PARAGRAPH.CENTER
            pf.first_line_indent = Cm(0)
            pf.left_indent = Cm(0)
            pf.keep_with_next = True          # ảnh không tách khỏi chú thích bên dưới
            pf.space_before = Pt(6)
    log.append(f'ảnh: canh giữa {n_center}, thu nhỏ cho vừa lề {n_scale}'
               + (f'; !! {n_float} ảnh nổi (wrap text) nên đổi sang "In line with text"' if n_float else ''))


# ------------------------------------------------------------------ bảng
LAYOUT_HINT = re.compile(r'CỘNG H[OÒ]À?A? XÃ HỘI|Nơi nhận|^\s*Số\s*:|Độc lập', re.I | re.M)
NUM_CELL = re.compile(r'^[\s(]*[-+±<>≤≥]?\s*\d[\d.,]*\s*(%|‰)?[\s)]*([±–\-]\s*\d[\d.,]*\s*%?)?[\s)]*$')


def has_borders(tbl):
    tblPr = tbl._tbl.tblPr
    b = tblPr.find(qn('w:tblBorders'))
    if b is not None:
        return any(e.get(qn('w:val')) not in ('nil', 'none') for e in b)
    st = tbl.style
    return bool(st is not None and st.name and st.name not in ('Normal Table', 'Table Normal'))


def is_layout_table(tbl):
    text = '\n'.join(c.text for r in tbl.rows for c in r.cells)
    if LAYOUT_HINT.search(text):
        return True
    return not has_borders(tbl) and len(tbl.columns) <= 2 and len(tbl.rows) <= 4


def _near_text(el, direction):
    e = el.getprevious() if direction < 0 else el.getnext()
    while e is not None and e.tag == qn('w:p') and not Paragraph(e, None).text.strip():
        e = e.getprevious() if direction < 0 else e.getnext()
    return Paragraph(e, None).text.strip() if e is not None and e.tag == qn('w:p') else ''


def table_caption(tbl):
    """'above' / 'below' / None: vị trí tên "Bảng ..." so với bảng."""
    cap = re.compile(r'^Bảng\s+\d')
    if cap.match(_near_text(tbl._tbl, -1)):
        return 'above'
    if cap.match(_near_text(tbl._tbl, 1)):
        return 'below'
    return None


def set_borders(tbl, sz=4):
    tblPr = tbl._tbl.tblPr
    b = ensure(tblPr, 'tblBorders', 'tblPr')
    for c in list(b):
        b.remove(c)
    for side in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        e = OxmlElement('w:' + side)
        for k, v in (('w:val', 'single'), ('w:sz', str(sz)), ('w:space', '0'), ('w:color', 'auto')):
            e.set(qn(k), v)
        b.append(e)


def step_tables(doc, size=12, all_tables=False):
    n = 0
    skipped = 0
    for tbl in doc.tables:
        # chỉ chuẩn hoá bảng số liệu có tên "Bảng x.x." (bảng bìa, chữ ký, phiếu phụ lục giữ nguyên)
        if is_layout_table(tbl) or not (all_tables or table_caption(tbl)):
            skipped += 1
            continue
        n += 1
        tblPr = tbl._tbl.tblPr
        w = ensure(tblPr, 'tblW', 'tblPr')
        w.set(qn('w:w'), '5000')
        w.set(qn('w:type'), 'pct')                 # rộng bằng vùng chữ, không tràn lề
        ensure(tblPr, 'jc', 'tblPr').set(qn('w:val'), 'center')
        ind = tblPr.find(qn('w:tblInd'))
        if ind is not None:
            tblPr.remove(ind)
        ensure(tblPr, 'tblLayout', 'tblPr').set(qn('w:type'), 'autofit')
        set_borders(tbl)
        for ri, row in enumerate(tbl.rows):
            trPr = row._tr.get_or_add_trPr()
            ensure(trPr, 'cantSplit', 'trPr')
            if ri == 0:
                ensure(trPr, 'tblHeader', 'trPr')     # lặp dòng tiêu đề khi bảng sang trang
            seen = set()
            for ci, cell in enumerate(row.cells):
                if id(cell._tc) in seen:
                    continue
                seen.add(id(cell._tc))
                cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                numeric = bool(NUM_CELL.match(cell.text)) and cell.text.strip() != ''
                for p in cell.paragraphs:
                    pf = p.paragraph_format
                    pf.first_line_indent = Cm(0)
                    pf.left_indent = Cm(0)
                    pf.space_before = Pt(2)
                    pf.space_after = Pt(2)
                    pf.line_spacing = 1.0
                    if ri == 0:
                        pf.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    elif numeric:
                        pf.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    elif pf.alignment != WD_ALIGN_PARAGRAPH.CENTER:
                        pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    for r in p.runs:
                        r.font.size = Pt(size)
                        if ri == 0:
                            r.font.bold = True
    log.append(f'bảng: chuẩn hoá {n} bảng số liệu (rộng bằng vùng chữ, canh giữa, viền đơn 0,5 pt, dòng tiêu đề '
               f'đậm + lặp lại khi sang trang, không ngắt dòng giữa trang, cỡ {size}, số canh giữa); '
               f'bỏ qua {skipped} bảng bố cục/không có tên "Bảng" (bìa, quốc hiệu, chữ ký, phụ lục; dùng --tables-all nếu cần)')


# ------------------------------------------------------------------ định dạng style tiêu đề
HEADING_FMT = {   # cấp: (cỡ, đậm, nghiêng, canh giữa, cách trước, cách sau)
    0: (14, True, False, True, 12, 12),
    1: (14, True, False, True, 12, 12),
    2: (None, True, False, False, 6, 6),
    3: (None, True, True, False, 6, 3),
    4: (None, False, True, False, 6, 3),
}


def format_heading_styles(doc, chosen, body_size=13):
    """Đồng bộ cỡ/đậm/nghiêng/canh lề cho style từng cấp và xoá định dạng gán trực tiếp ở run tiêu đề."""
    styles = {}
    for lv, name in chosen.items():
        level = lv if isinstance(lv, int) else 0
        styles[name] = level
        size, b, i, center, before, after = HEADING_FMT.get(level, HEADING_FMT[4])
        texts = [p.text.strip() for p in doc.paragraphs if p.style.name == name and p.text.strip()]
        if level == 1 and texts and sum(bool(re.match(r'[IVXLC]+\.\s', x)) for x in texts) > len(texts) / 2:
            size, center = None, False       # báo cáo hành chính: mục I., II. in hoa đậm, canh trái
        st = doc.styles[name]
        st.font.size = Pt(size or body_size)
        st.font.bold, st.font.italic = b, i
        st.font.name = st.font.name or None
        st.font.color.rgb = None
        pf = st.paragraph_format
        pf.alignment = WD_ALIGN_PARAGRAPH.CENTER if center else WD_ALIGN_PARAGRAPH.LEFT
        pf.first_line_indent = Cm(0)
        pf.left_indent = Cm(0)
        pf.space_before, pf.space_after = Pt(before), Pt(after)
        pf.keep_with_next = True
    n = 0
    for p in doc.paragraphs:
        if p.style.name in styles:
            for r in p.runs:
                rPr = r._r.rPr
                if rPr is None:
                    continue
                for t in ('w:b', 'w:bCs', 'w:i', 'w:iCs', 'w:sz', 'w:szCs', 'w:color'):
                    e = rPr.find(qn(t))
                    if e is not None:
                        rPr.remove(e)
                        n += 1
            pPr = p._p.pPr
            if pPr is not None:
                for t in ('w:ind', 'w:jc'):
                    e = pPr.find(qn(t))
                    if e is not None:
                        pPr.remove(e)
    log.append('định dạng tiêu đề: ' + ', '.join(
        f'cấp {lv if isinstance(lv, int) else 0} "{nm}"' for lv, nm in chosen.items())
        + f' (cấp 1: in đậm 14 canh giữa; cấp 2: đậm; cấp 3: đậm nghiêng; cấp 4: nghiêng); '
        f'xoá {n} định dạng gán trực tiếp ở tiêu đề')


# ------------------------------------------------------------------ số trang
def _page_field(p, size):
    for kind, text in (('begin', None), ('instr', ' PAGE '), ('separate', None), ('text', '1'), ('end', None)):
        r = p.add_run()
        r.font.size = Pt(size)
        if kind == 'instr':
            it = OxmlElement('w:instrText')
            it.set(qn('xml:space'), 'preserve')
            it.text = text
            r._r.append(it)
        elif kind == 'text':
            r.text = text
        else:
            fc = OxmlElement('w:fldChar')
            fc.set(qn('w:fldCharType'), kind)
            r._r.append(fc)


def has_page_number(doc):
    for s in doc.sections:
        for hf in (s.header, s.footer, s.first_page_header):
            if hf.is_linked_to_previous and s is not doc.sections[0]:
                continue
            for x in hf._element.iter(qn('w:instrText')):
                if 'PAGE' in (x.text or ''):
                    return True
            for x in hf._element.iter(qn('w:fldSimple')):
                if 'PAGE' in (x.get(qn('w:instr')) or ''):
                    return True
    return False


def step_page_number(doc):
    """Số trang giữa lề trên, không hiện ở trang đầu (NĐ 30). Bỏ qua nếu đã có số trang."""
    if has_page_number(doc):
        log.append('số trang: đã có, giữ nguyên')
        return
    s = doc.sections[0]
    if any(p.text.strip() for p in s.header.paragraphs):
        log.append('!! số trang: header đang có nội dung, không tự chèn; chèn tay Insert > Page Number > Top')
        return
    s.different_first_page_header_footer = True
    p = s.header.paragraphs[0] if s.header.paragraphs else s.header.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _page_field(p, SIZES['so_trang'][0])
    log.append('số trang: chèn giữa lề trên, ẩn ở trang đầu')


# ------------------------------------------------------------------ kiểm tra (cho audit_docx)
def audit_layout(doc, R):
    roles = detect_layout(doc)
    admin = bool({'quoc_hieu', 'tieu_ngu', 'noi_nhan', 'so_ky_hieu'} & set(roles))
    if admin:
        for k in ('quoc_hieu', 'tieu_ngu', 'co_quan', 'so_ky_hieu', 'dia_danh', 'noi_nhan'):
            if k not in roles:
                R['the_thuc_thieu'].append(LABEL[k])
        for k, ps in roles.items():
            if k.startswith('_') or k not in SIZES:
                continue
            size, bold, italic, upper = SIZES[k]
            lo, hi = RANGES.get(k, (size, size))
            for p in ps:
                pr = para_props(p)
                if not pr:
                    continue
                errs = []
                if any(not lo <= s <= hi for s in pr['size']):
                    errs.append(f'cỡ {"/".join(f"{s:g}" for s in pr["size"])} (cần {lo}–{hi})' if lo != hi
                                else f'cỡ {"/".join(f"{s:g}" for s in pr["size"])} (cần {lo})')
                if bold and not pr['bold']:
                    errs.append('cần in đậm')
                if not bold and pr['any_bold']:
                    errs.append('không in đậm')
                if italic and not pr['italic']:
                    errs.append('cần in nghiêng')
                if not italic and pr['any_italic']:
                    errs.append('không in nghiêng')
                t = p.text.strip()
                if upper and t != t.upper():
                    errs.append('cần IN HOA')
                if k not in ('noi_nhan', 'noi_nhan_ds', 'kinh_gui') and eff_align(p) != WD_ALIGN_PARAGRAPH.CENTER:
                    errs.append('cần canh giữa')
                if errs:
                    R['the_thuc_dinh_dang'].append(f'{LABEL[k]} "{t[:40]}": ' + ', '.join(errs))
        for p in roles.get('quoc_hieu', []):
            if p.text.strip() != QUOC_HIEU:
                R['the_thuc_noi_dung'].append(f'Quốc hiệu ghi "{p.text.strip()}" → "{QUOC_HIEU}"')
        for p in roles.get('tieu_ngu', []):
            if p.text.strip() != TIEU_NGU:
                R['the_thuc_noi_dung'].append(f'Tiêu ngữ ghi "{p.text.strip()}" → "{TIEU_NGU}"')
            if not has_rule_near(p):
                R['the_thuc_noi_dung'].append('chưa có đường kẻ dưới tiêu ngữ (dài bằng dòng chữ)')
        if roles.get('co_quan') and not has_rule_near(roles['co_quan'][-1]):
            R['the_thuc_noi_dung'].append('chưa có đường kẻ dưới tên cơ quan ban hành (1/3–1/2 dòng chữ)')
        for p in roles.get('so_ky_hieu', []):
            m = re.match(r'\s*Số\s*:?\s*(\d+)\s*/', p.text)
            if m and len(m.group(1)) < 2:
                R['the_thuc_noi_dung'].append(f'"{p.text.strip()[:30]}": số nhỏ hơn 10 phải thêm 0 (Số: 05/...)')
            if not re.match(r'\s*Số:\s', p.text):
                R['the_thuc_noi_dung'].append(f'"{p.text.strip()[:30]}": viết "Số: " (chữ thường, dấu hai chấm)')
        for p in roles.get('dia_danh', []):
            m = re.search(r'ngày\s+(\d{1,2})\s+tháng\s+(\d{1,2})', p.text)
            if m and (len(m.group(1)) < 2 or (int(m.group(2)) <= 2 and len(m.group(2)) < 2)):
                R['the_thuc_noi_dung'].append(f'"{p.text.strip()}": ngày < 10 và tháng 1, 2 phải thêm 0')
        ds = roles.get('noi_nhan_ds', [])
        for i, p in enumerate(ds):
            t = p.text.rstrip()
            end = '.' if i == len(ds) - 1 else ';'
            if t and t[-1] != end:
                R['the_thuc_noi_dung'].append(f'nơi nhận "{t.strip()[:40]}": cuối dòng cần "{end}"')
        if roles.get('noi_nhan') and not roles.get('chuc_vu'):
            R['the_thuc_noi_dung'].append('không thấy khối chữ ký (chức vụ, họ tên) trong bảng 2 cột cùng hàng '
                                          '"Nơi nhận:"; nên đặt nơi nhận bên trái, chữ ký bên phải')
    # ảnh
    maxw = text_width_emu(doc)
    ps = doc.paragraphs
    from audit_docx import CAPTION_RE
    for i, p in enumerate(ps):
        anc = p._p.findall('.//{%s}anchor' % WP)
        inl = p._p.findall('.//{%s}inline' % WP)
        if anc:
            R['anh'].append(f'[{i}] {len(anc)} ảnh/khung nổi (wrap text): nên đặt "In line with text" để không trôi')
        for d in inl:
            cx = int(d.find('{%s}extent' % WP).get('cx'))
            if cx > maxw * 1.01:
                R['anh'].append(f'[{i}] ảnh rộng {Emu(cx).mm:.0f} mm > vùng chữ {Emu(maxw).mm:.0f} mm')
        if inl and not p.text.strip():
            if eff_align(p) != WD_ALIGN_PARAGRAPH.CENTER:
                R['anh'].append(f'[{i}] ảnh chưa canh giữa')
            nxt = next((q for q in ps[i + 1:i + 4] if q.text.strip()), None)
            prv = next((q for q in reversed(ps[max(0, i - 3):i]) if q.text.strip()), None)
            is_cap = lambda q: q is not None and CAPTION_RE.match(q.text.strip()) and \
                not q.text.strip().startswith('Bảng')
            if not is_cap(nxt):
                R['anh'].append(f'[{i}] ảnh không có chú thích "Hình/Biểu đồ x.x." ngay bên dưới'
                                + (' (đang đặt phía trên)' if is_cap(prv) else ''))
    # bảng
    no_cap = []
    for ti, tbl in enumerate(doc.tables):
        if is_layout_table(tbl):
            continue
        el = tbl._tbl
        first = re.sub(r'\s+', ' ', tbl.rows[0].cells[0].text.strip())[:25] if tbl.rows else ''
        pos = table_caption(tbl)
        if pos is None:
            no_cap.append(str(ti + 1))
            continue
        if pos == 'below':
            R['bang'].append(f'bảng {ti + 1} ("{first}"): tên bảng đang đặt phía dưới, cần đặt phía trên')
        if not has_borders(tbl):
            R['bang'].append(f'bảng {ti + 1} ("{first}"): không có đường viền')
        trPr = tbl.rows[0]._tr.trPr if tbl.rows else None
        if trPr is None or trPr.find(qn('w:tblHeader')) is None:
            R['bang'].append(f'bảng {ti + 1} ("{first}"): dòng tiêu đề chưa đặt lặp lại khi sang trang')
        grid = el.find(qn('w:tblGrid'))
        if grid is not None:
            tw = sum(int(g.get(qn('w:w'), 0)) for g in grid) / TW_PER_MM
            if tw > Emu(maxw).mm + 2:
                R['bang'].append(f'bảng {ti + 1} ("{first}"): rộng {tw:.0f} mm > vùng chữ {Emu(maxw).mm:.0f} mm')
        sizes = {eff(p, r, 'size') for row in tbl.rows for c in row.cells for p in c.paragraphs for r in p.runs
                 if r.text.strip()}
        bad = sorted(s for s in sizes if not 10 <= s <= 13)
        if bad:
            R['bang'].append(f'bảng {ti + 1} ("{first}"): cỡ chữ {bad} (nên 11–13, thống nhất các bảng)')
    if no_cap:
        R['bang'].append(f'{len(no_cap)} bảng không có tên "Bảng x.x." (bảng số {", ".join(no_cap[:15])}): '
                         'nếu là bảng số liệu thì thêm tên phía trên; bảng bìa/phiếu phụ lục thì bỏ qua')
    if not has_page_number(doc):
        R['trang_giay'].append('chưa có số trang (giữa lề trên, ẩn trang đầu)')
    return roles
