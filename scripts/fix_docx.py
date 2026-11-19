#!/usr/bin/env python3
"""Chuẩn hoá thể thức file .docx theo Nghị định 30/2020 và quy ước báo cáo khoa học.

Không ghi đè file nguồn. Chỉ sửa định dạng và dấu câu của số thứ tự; KHÔNG tự đánh lại số chương/mục/bảng
(số trùng, số thiếu là lỗi nội dung, cần người viết quyết định; xem audit_docx.py).

    python fix_docx.py IN.docx -o OUT.docx --all
    python fix_docx.py IN.docx -o OUT.docx --headings --captions --toc

Các bước (chạy theo thứ tự dưới đây):
  --page       khổ A4, lề trên/dưới/trái/phải theo --margins (mặc định 20,20,30,15 mm)
  --font       Times New Roman cho mọi style và docDefaults; --font-runs: cả phông gán trực tiếp ở run
  --body       style Normal: cỡ --size, canh đều, thụt đầu dòng 1,27 cm, cách sau 6 pt; --line: giãn dòng
  --headings   gán style tiêu đề theo số thứ tự (Chương, 1.1., 1.1.1., 1.1.1.1.) để lên mục lục
  --chapter    arabic|roman: thống nhất "Chương 1." hoặc "Chương I." (dấu chấm sau số)
  --heading-dot thêm dấu chấm sau số mục: "3.1 Tên" -> "3.1. Tên"
  --captions   chú thích Bảng/Biểu đồ/Hình: style riêng từng loại, dấu "." sau số, ghép đoạn bị tách
  --toc        chèn/thay trường Mục lục, Danh mục bảng/biểu đồ/hình sau các tiêu đề tương ứng
  --heading-format  đồng bộ cỡ/đậm/nghiêng/canh lề từng cấp tiêu đề (kéo theo --headings)
  --admin      thể thức hành chính NĐ 30: quốc hiệu, tiêu ngữ, tên cơ quan, số ký hiệu, địa danh ngày tháng,
               tên loại, trích yếu, V/v, kính gửi, nơi nhận, khối chữ ký; đường kẻ; thêm 0 cho ngày/số
  --images     ảnh: canh giữa, thu nhỏ cho vừa lề, giữ liền với chú thích
  --tables     bảng số liệu: rộng bằng vùng chữ, viền đơn, dòng tiêu đề đậm + lặp lại, cỡ --table-size
  --page-number số trang giữa lề trên, ẩn trang đầu (nếu chưa có)
  --all        = tất cả các bước trên (trừ --chapter, --line)
"""
import argparse, copy, os, re, sys
from collections import Counter
import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Mm, Pt, Cm

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from audit_docx import classify_heading, CAPTION_RE, is_bold  # noqa: E402
import the_thuc  # noqa: E402
from the_thuc import replace_span  # noqa: E402

BODY_STYLES = ('Normal', 'Normal (Web)', 'No Spacing', 'Body Text', 'List Paragraph')
FRONT_TITLES = ('LỜI CAM ĐOAN', 'LỜI CẢM ƠN', 'DANH MỤC', 'MỤC LỤC', 'TÓM TẮT', 'LỜI NÓI ĐẦU')
TOC_TARGETS = [  # (tiêu đề, loại)
    ('MỤC LỤC', 'toc'), ('DANH MỤC BẢNG', 'Bảng'), ('DANH MỤC CÁC BẢNG', 'Bảng'),
    ('DANH MỤC BIỂU ĐỒ', 'Biểu đồ'), ('DANH MỤC CÁC BIỂU ĐỒ', 'Biểu đồ'),
    ('DANH MỤC HÌNH', 'Hình'), ('DANH MỤC HÌNH ẢNH', 'Hình'), ('DANH MỤC CÁC HÌNH', 'Hình'),
    ('DANH MỤC SƠ ĐỒ', 'Sơ đồ'),
]
PLACEHOLDER = 'Nhấn Ctrl+A rồi F9 trong Word để cập nhật mục lục / danh mục.'
ROMAN = [(10, 'X'), (9, 'IX'), (5, 'V'), (4, 'IV'), (1, 'I')]

log = []


# ---------------------------------------------------------------- tiện ích
def to_roman(n):
    out = ''
    for v, s in ROMAN:
        while n >= v:
            out, n = out + s, n - v
    return out


def from_roman(s):
    vals = {'I': 1, 'V': 5, 'X': 10, 'L': 50, 'C': 100}
    total = 0
    for i, c in enumerate(s):
        v = vals[c]
        total += -v if i + 1 < len(s) and vals[s[i + 1]] > v else v
    return total


def sub_prefix(p, pattern, repl):
    """Áp regex (neo đầu đoạn) lên nội dung run, chỉ sửa phần khớp."""
    text = ''.join(r.text for r in p.runs)
    lead = len(text) - len(text.lstrip())
    m = re.compile(pattern).match(text, lead)
    if not m:
        return False
    new = m.expand(repl)
    if new == m.group(0):
        return False
    replace_span(p, m.start(), m.end(), new)
    return True


def new_style(doc, name, base=None, outline=None, bold=None, center=False):
    """Tạo style đoạn mới (nhân bản `base` nếu có), bỏ cờ mặc định."""
    if name in [s.name for s in doc.styles]:
        st = doc.styles[name]
    else:
        if base is not None:
            el = copy.deepcopy(base.element)
            el.set(qn('w:styleId'), re.sub(r'\W', '', name, flags=re.ASCII) or 'S%d' % len(doc.styles))
            el.find(qn('w:name')).set(qn('w:val'), name)
            for attr in ('w:default',):
                if el.get(qn(attr)) is not None:
                    del el.attrib[qn(attr)]
            for tag in ('w:link', 'w:aliases'):
                x = el.find(qn(tag))
                if x is not None:
                    el.remove(x)
            doc.styles.element.append(el)
            st = doc.styles[name]
        else:
            st = doc.styles.add_style(name, 1)
            st.base_style = doc.styles['Normal']
        st.quick_style = True
    if bold is not None:
        st.font.bold = bold
    if center:
        st.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        st.paragraph_format.first_line_indent = Cm(0)
    if outline is not None:
        set_outline(st.element.get_or_add_pPr(), outline)
    return st


def set_outline(pPr, level):
    o = pPr.find(qn('w:outlineLvl'))
    if o is None:
        o = OxmlElement('w:outlineLvl')
        # đúng thứ tự schema: outlineLvl đứng trước divId, cnfStyle, rPr, sectPr, pPrChange
        nxt = next((c for c in pPr if c.tag in {qn(t) for t in
                    ('w:divId', 'w:cnfStyle', 'w:rPr', 'w:sectPr', 'w:pPrChange')}), None)
        nxt.addprevious(o) if nxt is not None else pPr.append(o)
    o.set(qn('w:val'), str(level))


def is_toc_para(p):
    n = p.style.name.lower()
    return n.startswith('toc') or n.startswith('table of figures')


def body_start(ps):
    """Chỉ số đoạn bắt đầu nội dung chính (tiêu đề chương/đặt vấn đề đầu tiên không nằm trong mục lục)."""
    for i, p in enumerate(ps):
        t = p.text.strip().upper()
        if is_toc_para(p) or '\t' in p.text.rstrip():
            continue
        if re.match(r'(CHƯƠNG|PHẦN)\s+[0-9IVX]+|ĐẶT VẤN ĐỀ|MỞ ĐẦU|LỜI MỞ ĐẦU|I\.\s|PHẦN THỨ NHẤT', t):
            # bỏ qua dòng mục lục gõ tay: tiêu đề thật thường có style tiêu đề hoặc đậm
            if p.style.name not in BODY_STYLES or is_bold(p):
                return i
    return 0


def list_sep(doc):
    """Dấu phân cách danh sách trong \\t của field TOC theo field có sẵn (Word phụ thuộc cài đặt vùng)."""
    for x in doc.element.body.iter(qn('w:instrText')):
        m = re.search(r'\\t\s+"[^",;]+([,;])\d', x.text or '')
        if m:
            return m.group(1)
    return ','


# ---------------------------------------------------------------- các bước
def step_page(doc, margins):
    t, b, l, r = margins
    for s in doc.sections:
        portrait = s.page_width <= s.page_height
        s.page_width, s.page_height = (Mm(210), Mm(297)) if portrait else (Mm(297), Mm(210))
        s.top_margin, s.bottom_margin, s.left_margin, s.right_margin = Mm(t), Mm(b), Mm(l), Mm(r)
    log.append(f'trang: A4, lề trên {t}, dưới {b}, trái {l}, phải {r} mm cho {len(doc.sections)} section')


def _set_rfonts(rPr, font):
    rf = rPr.find(qn('w:rFonts'))
    if rf is None:
        rf = OxmlElement('w:rFonts')
        rs = rPr.find(qn('w:rStyle'))
        rs.addnext(rf) if rs is not None else rPr.insert(0, rf)
    for a in ('w:asciiTheme', 'w:hAnsiTheme', 'w:cstheme', 'w:eastAsiaTheme'):
        if rf.get(qn(a)) is not None:
            del rf.attrib[qn(a)]
    for a in ('w:ascii', 'w:hAnsi', 'w:cs', 'w:eastAsia'):
        rf.set(qn(a), font)


def step_font(doc, font, runs_too):
    dd = doc.styles.element.find(qn('w:docDefaults'))
    if dd is not None:
        rpd = dd.find(qn('w:rPrDefault'))
        if rpd is not None and rpd.find(qn('w:rPr')) is not None:
            _set_rfonts(rpd.find(qn('w:rPr')), font)
    n = 0
    for st in doc.styles.element.findall(qn('w:style')):
        rPr = st.find(qn('w:rPr'))
        if rPr is not None and rPr.find(qn('w:rFonts')) is not None:
            _set_rfonts(rPr, font)
            n += 1
    m = 0
    if runs_too:
        keep = re.compile(r'(?i)symbol|wingdings|webdings|math')
        for rf in doc.element.body.iter(qn('w:rFonts')):
            if not keep.search(rf.get(qn('w:ascii')) or '') and rf.getparent().tag == qn('w:rPr'):
                _set_rfonts(rf.getparent(), font)
                m += 1
    log.append(f'phông: {font} cho docDefaults, {n} style' + (f', {m} run' if runs_too else ''))


def step_body(doc, size, line):
    st = doc.styles['Normal']
    st.font.size = Pt(size)
    pf = st.paragraph_format
    pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    pf.first_line_indent = Cm(1.27)
    pf.space_after = Pt(6)
    if line:
        pf.line_spacing = line
    # bỏ Tab/dấu cách dùng để lùi đầu dòng
    n = 0
    for p in doc.paragraphs:
        if p.style.name in BODY_STYLES and p.runs and p.runs[0].text[:1] in ('\t', ' ') and len(p.text) > 60:
            p.runs[0].text = p.runs[0].text.lstrip('\t ')
            n += 1
    log.append(f'thân bài: Normal cỡ {size}, canh đều, thụt 1,27 cm, cách sau 6 pt'
               + (f', giãn dòng {line}' if line else '') + f'; bỏ Tab/dấu cách đầu dòng ở {n} đoạn')


def detect_headings(doc):
    ps = doc.paragraphs
    start = body_start(ps)
    out = []
    for i, p in enumerate(ps):
        if is_toc_para(p) or CAPTION_RE.match(p.text.strip()):
            continue
        c = classify_heading(p)
        if not c:
            continue
        lv = c[0]
        t = p.text.strip().upper()
        if i < start:
            if any(t.startswith(x) for x in FRONT_TITLES):
                out.append((i, 0, p))         # 0 = tiêu đề phần đầu
            continue
        if lv in (1, 2, 3, 4):
            out.append((i, lv, p))
    return out


def step_headings(doc):
    heads = detect_headings(doc)
    if not heads:
        log.append('tiêu đề: không phát hiện tiêu đề nào')
        return {}
    chosen = {}
    used = set()
    for lv in (1, 2, 3, 4):
        cnt = Counter(p.style.name for _, l, p in heads if l == lv and p.style.name not in BODY_STYLES)
        top = cnt.most_common(1)[0][0] if cnt else None
        pick = top if top and top not in used else None
        if pick is None and any(l == lv for _, l, _ in heads):
            # style trội đã dùng cho cấp trên (hoặc chưa có): nhân bản để giữ hình thức, tách cấp
            src = top or chosen.get(lv - 1)
            base = doc.styles[src] if src else None
            pick = new_style(doc, f'Muc cap {lv}', base=base, bold=True).name
            log.append(f'tạo style "{pick}" cho tiêu đề cấp {lv}' + (f' (nhân bản "{src}")' if src else ''))
        if pick:
            chosen[lv] = pick
            used.add(pick)
            set_outline(doc.styles[pick].element.get_or_add_pPr(), lv - 1)
    extra = []   # style riêng sẵn có của tiêu đề phần đầu (giữ nguyên, chỉ đưa vào mục lục)
    for _, l, p in heads:
        if l == 0 and p.text.strip().upper() != 'MỤC LỤC' and p.style.name not in BODY_STYLES \
                and p.style.name not in chosen.values() and p.style.name not in extra:
            extra.append(p.style.name)
    if any(l == 0 and p.style.name in BODY_STYLES for _, l, p in heads):
        pick = new_style(doc, 'Tieu de phan dau', base=doc.styles['Normal'], outline=0, bold=True, center=True)
        chosen[0] = pick.name
    changes = Counter()
    for i, lv, p in heads:
        if lv not in chosen or p.style.name == chosen[lv]:
            continue
        if lv == 0 and (p.text.strip().upper() == 'MỤC LỤC' or p.style.name not in BODY_STYLES):
            continue
        p.style = doc.styles[chosen[lv]]
        changes[lv] += 1
    log.append('tiêu đề: style theo cấp ' + ', '.join(f'{k}→"{v}"' for k, v in sorted(chosen.items(), key=lambda kv: str(kv[0])))
               + (f'; giữ style phần đầu {extra}' if extra else '')
               + '; đã gán lại ' + ', '.join(f'cấp {k}: {v}' for k, v in sorted(changes.items())))
    for k, name in enumerate(extra):
        chosen[f'0{k}'] = name
    return chosen


def step_chapter(doc, mode):
    n = 0
    for p in doc.paragraphs[body_start(doc.paragraphs):]:
        m = re.match(r'\s*(Chương|CHƯƠNG)\s+([0-9]+|[IVXLC]+)\s*[.:\-–]?\s*', p.text)
        if not m or is_toc_para(p):
            continue
        k = int(m.group(2)) if m.group(2).isdigit() else from_roman(m.group(2))
        num = to_roman(k) if mode == 'roman' else str(k)
        if sub_prefix(p, r'(Chương|CHƯƠNG)\s+(?:[0-9]+|[IVXLC]+)\s*[.:\-–]?\s*', rf'\g<1> {num}. '):
            n += 1
    log.append(f'tên chương: thống nhất "Chương {"I" if mode == "roman" else "1"}." ({n} chỗ sửa)')


def step_heading_dot(doc):
    n = 0
    for p in doc.paragraphs[body_start(doc.paragraphs):]:
        if is_toc_para(p) or not classify_heading(p):
            continue
        if sub_prefix(p, r'(\d+(?:\.\d+)+)(?:\s*[:\-–]\s*|\s+)(?=\S)', r'\1. '):
            n += 1
    log.append(f'số mục: thêm dấu chấm sau số ở {n} tiêu đề')


def step_captions(doc):
    ps = doc.paragraphs
    start = body_start(ps)
    found = {}
    for i, p in enumerate(ps):
        if i < start or is_toc_para(p):
            continue
        m = CAPTION_RE.match(p.text.strip())
        if m and len(p.text) < 300:
            kind = 'Biểu đồ' if m.group('kind') == 'Biều đồ' else m.group('kind')
            found.setdefault(kind, []).append(p)
    styles = {}
    stats = Counter()
    for kind, items in found.items():
        cnt = Counter(p.style.name for p in items if p.style.name not in BODY_STYLES)
        pick = next((s for s, _ in cnt.most_common() if s not in styles.values()), None)
        if pick is None:
            st = new_style(doc, f'Chu thich {kind}', base=doc.styles['Normal'], bold=True, center=True)
            if kind == 'Bảng':
                st.paragraph_format.keep_with_next = True
            pick = st.name
            log.append(f'tạo style "{pick}"')
        styles[kind] = pick
        for p in items:
            # ghép đoạn bị tách (Enter giữa câu)
            nxt = p._p.getnext()
            if nxt is not None and nxt.tag == qn('w:p'):
                np_ = docx.text.paragraph.Paragraph(nxt, p._parent)
                nt = np_.text.strip()
                if nt and len(nt) < 120 and not CAPTION_RE.match(nt) and not p.text.rstrip().endswith('.') \
                        and not re.match(r'(\d+(\.\d+)*|[IVXLC]+|[a-zđ])[.):]?\s', nt) \
                        and np_.style.name in (p.style.name, 'Normal') and not nt.lower().startswith(('nguồn', 'nhận xét')) \
                        and not np_._p.findall('.//' + qn('w:drawing')):
                    if p.runs and not p.runs[-1].text.endswith(' '):
                        p.runs[-1].text += ' '
                    for r in np_.runs:
                        p._p.append(copy.deepcopy(r._r))
                    nxt.getparent().remove(nxt)
                    stats['ghép đoạn'] += 1
            for br in p._p.findall('.//' + qn('w:br')):
                if br.get(qn('w:type')) in (None, 'textWrapping'):
                    sp = OxmlElement('w:t')
                    sp.set(qn('xml:space'), 'preserve')
                    sp.text = ' '
                    br.addprevious(sp)
                    br.getparent().remove(br)
                    stats['bỏ ngắt dòng'] += 1
            if sub_prefix(p, r'Biều', 'Biểu'):
                stats['sửa "Biều"'] += 1
            if sub_prefix(p, r'(Bảng|Biểu đồ|Hình|Sơ đồ|Ảnh)\s+(\d+(?:\.\d+)?)\s*(?:[:\-–.]\s*|\s+)(?=\S)', r'\1 \2. '):
                stats['dấu sau số'] += 1
            # dọn khoảng trắng kép do ghép
            for r in p.runs:
                r.text = re.sub(r' {2,}', ' ', r.text)
            if p.style.name != pick:
                p.style = doc.styles[pick]
                stats['gán style'] += 1
    log.append('chú thích: ' + ', '.join(f'{k}→"{v}" ({len(found[k])})' for k, v in styles.items())
               + ('; ' + ', '.join(f'{k}: {v}' for k, v in stats.items()) if stats else ''))
    return styles


def _field_paragraph(instr):
    p = OxmlElement('w:p')
    for kind, text in (('begin', None), (None, instr), ('separate', None), ('text', PLACEHOLDER), ('end', None)):
        r = OxmlElement('w:r')
        if kind == 'text':
            t = OxmlElement('w:t')
            t.text = text
            r.append(t)
        elif kind is None:
            it = OxmlElement('w:instrText')
            it.set(qn('xml:space'), 'preserve')
            it.text = instr
            r.append(it)
        else:
            fc = OxmlElement('w:fldChar')
            fc.set(qn('w:fldCharType'), kind)
            r.append(fc)
        p.append(r)
    return p


def _is_manual_toc_line(p):
    t = p.text
    return is_toc_para(p) or 'Error! Bookmark' in t or bool(re.search(r'\t\s*[0-9ivxlc]+\s*$', t, re.I))


def _has_field(p):
    return p._p.find('.//' + qn('w:fldChar')) is not None


def clean_manual_toc(doc):
    """Xoá dòng mục lục gõ tay ở phần đầu: dòng có Tab + số trang, hoặc trùng chữ với một tiêu đề phía sau."""
    ps = doc.paragraphs
    start = body_start(ps)
    later = Counter(re.sub(r'\s+', ' ', p.text.strip().upper()) for p in ps if p.text.strip())
    seen = Counter()
    removed = []
    for p in ps[:start]:
        t = re.sub(r'\s+', ' ', p.text.strip().upper())
        if not t or _has_field(p) or is_toc_para(p):
            continue
        seen[t] += 1
        if (later[t] - seen[t] > 0 and len(t) < 120) or re.search(r'\t\s*[0-9ivxlc]+\s*$', p.text, re.I):
            removed.append(p.text.strip()[:40])
            p._p.getparent().remove(p._p)
    if removed:
        log.append(f'xoá {len(removed)} dòng mục lục gõ tay ở phần đầu: {removed}')


def _field_end(ps, j):
    """Chỉ số đoạn chứa fldChar 'end' khớp với 'begin' đầu tiên ở đoạn j."""
    depth, started = 0, False
    for k in range(j, len(ps)):
        for fc in ps[k]._p.iter(qn('w:fldChar')):
            t = fc.get(qn('w:fldCharType'))
            if t == 'begin':
                depth, started = depth + 1, True
            elif t == 'end' and started:
                depth -= 1
                if depth == 0:
                    return k
    return None


def step_toc(doc, head_styles, cap_styles):
    sep = list_sep(doc)
    ps = doc.paragraphs
    start = body_start(ps)
    # suy ra style tiêu đề/chú thích nếu không chạy --headings/--captions
    if not head_styles:
        head_styles = {}
        cnt = Counter((lv, p.style.name) for _, lv, p in detect_headings(doc) if p.style.name not in BODY_STYLES)
        for lv in sorted({lv for lv, _ in cnt}):
            names = [n for (l, n), _ in cnt.most_common() if l == lv]
            free = [n for n in names if n not in head_styles.values()]
            if free:
                head_styles[lv] = free[0]
            if len(names) > 1 or not free:
                log.append(f'!! tiêu đề cấp {lv} dùng nhiều style {names} hoặc trùng style cấp trên: '
                           'chạy kèm --headings để tách cấp')
    if not cap_styles:
        cap_styles = {}
        for p in ps[start:]:
            m = CAPTION_RE.match(p.text.strip())
            if m and p.style.name not in BODY_STYLES:
                cap_styles.setdefault('Biểu đồ' if m.group('kind') == 'Biều đồ' else m.group('kind'), p.style.name)
    long_doc = len(ps) > 120
    done = set()
    i = -1
    while True:
        i += 1
        if i >= start:
            break
        p = ps[i]
        t = re.sub(r'\s+', ' ', p.text.strip().upper())
        target = next((k for title, k in TOC_TARGETS if t == title), None)
        if target is None or target in done:
            continue
        if target == 'toc':
            order = sorted(head_styles.items(), key=lambda kv: str(kv[0]))
            parts = [f'{s}{sep}{lv if isinstance(lv, int) and lv else 1}' for lv, s in order]
            if not parts:
                log.append('!! mục lục: không có style tiêu đề nào, bỏ qua (chạy kèm --headings)')
                continue
            instr = f' TOC \\h \\z \\t "{sep.join(parts)}" '
        else:
            if target not in cap_styles:
                instr = f' TOC \\h \\z \\c "{target}" '
                log.append(f'!! {t}: không có chú thích "{target}" nào trong văn bản, chèn field rỗng')
            else:
                instr = f' TOC \\h \\z \\t "{cap_styles[target]}{sep}1" '
        # tìm field TOC có sẵn ngay sau tiêu đề (bỏ qua đoạn trống)
        j = i + 1
        while j < start and not ps[j].text.strip() and not _has_field(ps[j]):
            j += 1
        is_toc = j < start and any((x.text or '').strip().startswith('TOC')
                                   for x in ps[j]._p.iter(qn('w:instrText')))
        newp = _field_paragraph(instr)
        if is_toc and (k := _field_end(ps, j)) is not None:
            ps[j]._p.addprevious(newp)
            for q in ps[j:k + 1]:
                q._p.getparent().remove(q._p)
            log.append(f'{t}: thay field cũ ({k - j + 1} đoạn) bằng {instr.strip()}')
        else:
            p._p.addnext(newp)
            log.append(f'{t}: chèn field {instr.strip()}')
        ps = doc.paragraphs
        start = body_start(ps)
        done.add(target)
    for k in dict.fromkeys(k for _, k in TOC_TARGETS) if long_doc else []:
        if k != 'toc' and k in cap_styles and k not in done:
            log.append(f'!! có chú thích "{k}" nhưng không thấy tiêu đề "DANH MỤC {k.upper()}" ở phần đầu')
    if 'toc' not in done and long_doc:
        log.append('!! không thấy tiêu đề "MỤC LỤC" ở phần đầu; chưa chèn mục lục')
    st = doc.settings.element
    uf = st.find(qn('w:updateFields'))
    if uf is None:
        uf = OxmlElement('w:updateFields')
        st.append(uf)
    uf.set(qn('w:val'), 'true')


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('src')
    ap.add_argument('-o', '--out', required=True)
    for f in ('page', 'font', 'font-runs', 'body', 'headings', 'heading-dot', 'heading-format', 'captions', 'toc',
              'admin', 'images', 'tables', 'page-number', 'all'):
        ap.add_argument('--' + f, action='store_true')
    ap.add_argument('--margins', default='20,20,30,15', help='trên,dưới,trái,phải (mm)')
    ap.add_argument('--font-name', default='Times New Roman')
    ap.add_argument('--size', type=float, default=13)
    ap.add_argument('--line', type=float, help='giãn dòng, 1.0–1.5')
    ap.add_argument('--chapter', choices=['arabic', 'roman'])
    ap.add_argument('--tables-all', action='store_true', help='chuẩn hoá cả bảng không có tên "Bảng"')
    ap.add_argument('--table-size', type=float, default=12, help='cỡ chữ trong bảng (11–13)')
    a = ap.parse_args()
    if os.path.abspath(a.src) == os.path.abspath(a.out):
        sys.exit('Không ghi đè file nguồn: chọn -o khác.')
    if a.all:
        a.page = a.font = a.font_runs = a.body = a.headings = a.heading_dot = a.captions = a.toc = True
        a.heading_format = a.admin = a.images = a.tables = a.page_number = True
    if a.heading_format:
        a.headings = True
    if a.line and not 1.0 <= a.line <= 1.5:
        log.append(f'!! giãn dòng {a.line} ngoài khoảng 1–1,5 của NĐ 30/2020')
    doc = docx.Document(a.src)
    if a.page:
        step_page(doc, [float(x) for x in a.margins.split(',')])
    if a.font or a.font_runs:
        step_font(doc, a.font_name, a.font_runs)
    if a.body:
        step_body(doc, a.size, a.line)
    if a.toc:
        clean_manual_toc(doc)
    heads = step_headings(doc) if a.headings else {}
    if a.chapter:
        step_chapter(doc, a.chapter)
    if a.heading_dot:
        step_heading_dot(doc)
    caps = step_captions(doc) if a.captions else {}
    if a.heading_format and heads:
        the_thuc.format_heading_styles(doc, heads, a.size)
    if a.admin:
        the_thuc.step_admin(doc)
    if a.images:
        the_thuc.step_images(doc)
    if a.tables:
        the_thuc.step_tables(doc, a.table_size, a.tables_all)
    if a.page_number:
        the_thuc.step_page_number(doc)
    if a.toc:
        step_toc(doc, heads, caps)
    doc.save(a.out)
    print('\n'.join(log + the_thuc.log))
    print(f'\nĐã lưu: {a.out}')


if __name__ == '__main__':
    main()
