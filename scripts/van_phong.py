"""Văn phong kỹ thuật: kiểm tra (khẩu ngữ, chủ quan, câu dài, viết tắt, thuật ngữ, chính tả i/y, dấu thanh,
dấu câu, số và đơn vị) và sửa an toàn (khoảng trắng, dấu câu, đơn vị, dấu thanh, lặp hư từ).

Phần "kiểm tra" chỉ báo cáo; nội dung câu chữ do người viết quyết định. Phần "sửa" chỉ đụng tới lỗi trình bày
có quy tắc rõ ràng, giữ định dạng run.
"""
import re
from collections import Counter, defaultdict
from docx.oxml.ns import qn

import the_thuc as T

log = []

# ------------------------------------------------------------------ từ điển
# (mẫu regex, gợi ý)
KHAU_NGU = [
    (r'rất là', 'rất'), (r'khá là', 'khá'), (r'cực kỳ|cực kì|vô cùng|siêu (?:tốt|đẹp|nhanh|rẻ|hay|lớn|nhỏ|nhiều|cao|thấp)', 'bỏ, hoặc nêu số liệu cụ thể'),
    (r'quá trời|quá chừng|hết sức là', 'bỏ, hoặc nêu số liệu'), (r'nhiều lắm|lắm lắm', 'nhiều'),
    (r'tầm khoảng|cỡ khoảng|tầm(?= \d)', 'khoảng'), (r'hồi trước|hồi đó|dạo trước', 'trước đây'),
    (r'bây giờ|giờ đây', 'hiện nay'), (r'tụi|bọn(?= trẻ)', 'nhóm / các'), (r'cái này|cái đó|mấy cái', 'nội dung này / các …'),
    (r'vân vân|v\.v\.\.\.|v\.v…', 'v.v.'), (r'nói chung là', 'nhìn chung'), (r'kiểu như', 'chẳng hạn'),
    (r'thật sự là|thực sự là', 'bỏ'), (r'đủ thứ|đủ kiểu', 'nhiều loại'), (r'okay|oke|ok(?!\w)', 'bỏ'),
    (r'chẳng hạn như là', 'chẳng hạn'), (r'(?<=\w) (nhé|nha|ạ|đấy|đó nha)(?=[.!?]|$)', 'bỏ tiểu từ'),
]
CHU_QUAN = [
    (r'tôi nghĩ|tôi cho rằng|chúng tôi nghĩ|theo ý tôi|theo cảm nhận', 'nêu căn cứ: "kết quả cho thấy…", "theo [tài liệu]…"'),
    (r'có vẻ như|hình như|chắc là|chắc chắn là|dường như là', 'khẳng định theo số liệu hoặc nêu mức độ bằng thống kê'),
    (r'ai cũng biết|như mọi người đều biết|hiển nhiên là', 'dẫn nguồn'),
    (r'tuyệt vời|hoàn hảo|tốt nhất thế giới|đáng kinh ngạc', 'từ đánh giá cảm tính: thay bằng nhận xét định lượng'),
]
HU_TU = ('của', 'và', 'các', 'những', 'được', 'cho', 'với', 'trong', 'thì', 'mà', 'này', 'để', 'về', 'tại', 'bị')
# không đưa đơn vị 1 chữ dễ nhầm (A, V, W, D, s, l, L: "lớp 1A", "3D") vào danh sách
UNITS = r'mmHg|mm|cm|dm|km|nm|µm|μm|mg|kg|ml|mL|kPa|kcal|kW|kV|mA|kHz|MHz|Hz|GB|MB|KB|ms|m|g'
VOWEL_TONE = {  # dấu thanh kiểu cũ (oà) <-> kiểu mới (òa)
    'oà': 'òa', 'oá': 'óa', 'oả': 'ỏa', 'oã': 'õa', 'oạ': 'ọa',
    'oè': 'òe', 'oé': 'óe', 'oẻ': 'ỏe', 'oẽ': 'õe', 'oẹ': 'ọe',
    'uỳ': 'ùy', 'uý': 'úy', 'uỷ': 'ủy', 'uỹ': 'ũy', 'uỵ': 'ụy',
}
Y_I = {'y': 'i', 'ý': 'í', 'ỳ': 'ì', 'ỷ': 'ỉ', 'ỹ': 'ĩ', 'ỵ': 'ị'}
SKIP_STYLE = re.compile(r'(?i)^(toc|table of figures|endnote|bibliography|caption)')
REF_TITLE = re.compile(r'^(TÀI LIỆU THAM KHẢO|DANH MỤC TÀI LIỆU THAM KHẢO|REFERENCES)\s*$', re.I)


# ------------------------------------------------------------------ tiện ích
def paragraphs(doc):
    """(nhãn, Paragraph) cho đoạn cần xét văn phong: bỏ mục lục; dừng ở Tài liệu tham khảo (và phụ lục phía sau)."""
    idx = {p._p: i for i, p in enumerate(doc.paragraphs)}   # khoá là phần tử lxml (giữ sống proxy)
    for p in T.iter_paragraphs(doc):
        t = p.text.strip()
        if not t or SKIP_STYLE.match(p.style.name) or re.search(r'\t\s*\d+\s*$', p.text):   # dòng mục lục gõ tay
            continue
        if REF_TITLE.match(t):
            return
        if re.fullmatch(r'PHỤ LỤC\s*\d*', t) and i_of(idx, p) is not None and i_of(idx, p) > len(idx) * 0.5:
            return
        if any((x.text or '').strip().startswith('TOC') for x in p._p.iter(qn('w:instrText'))):
            continue
        i = idx.get(p._p)
        yield (f'[{i}]' if i is not None else '[bảng]'), p


def i_of(idx, p):
    return idx.get(p._p)


def ctx(text, m, w=25):
    a, b = max(0, m.start() - w), min(len(text), m.end() + w)
    return ('…' if a else '') + text[a:b].replace('\t', ' ') + ('…' if b < len(text) else '')


def syllables(text):
    return re.findall(r'[\wÀ-ỹ]+', text)


def sub_runs(p, pattern, repl, flags=0):
    """Thay theo regex trên nội dung các run của đoạn, giữ định dạng; trả về số chỗ thay."""
    text = T.run_text(p)
    ms = list(re.finditer(pattern, text, flags))
    for m in reversed(ms):
        new = m.expand(repl) if isinstance(repl, str) else repl(m)
        if new != m.group(0):
            T.replace_span(p, m.start(), m.end(), new)
    return sum(1 for m in ms if (m.expand(repl) if isinstance(repl, str) else repl(m)) != m.group(0))


def tone_variant(word, to_new):
    """Đổi vị trí dấu thanh oa/oe/uy ở âm tiết mở. 'qu' là phụ âm nên 'quý' giữ nguyên."""
    src = {v: k for k, v in VOWEL_TONE.items()} if not to_new else VOWEL_TONE
    lw = word.lower()
    for a, b in src.items():
        if lw.endswith(a) and not (a.startswith('u') and lw[:-len(a)].endswith('q')):
            out = word[:-len(a)] + (b.upper() if word[-len(a):].isupper() else
                                     b[0].upper() + b[1:] if word[-len(a):][0].isupper() else b)
            return out
    return None


# ------------------------------------------------------------------ kiểm tra
def audit(doc, R, max_ex=8):
    items = list(paragraphs(doc))
    all_text = '\n'.join(p.text for _, p in items)

    def hit(key, label, p, m, extra=''):
        if len(R[key]) < max_ex:
            R[key].append(f'{label} "{ctx(p.text, m)}"{extra}')
        else:
            R['_n_' + key] = R.get('_n_' + key, 0) + 1

    # 1. khẩu ngữ, chủ quan
    for key, table in (('vp_khau_ngu', KHAU_NGU), ('vp_chu_quan', CHU_QUAN)):
        for pat, sug in table:
            rx = re.compile(r'(?<![\wÀ-ỹ])(' + pat + r')(?![\wÀ-ỹ])', re.I)
            for lb, p in items:
                for m in rx.finditer(p.text):
                    hit(key, lb, p, m, f' → {sug}')
    for lb, p in items:
        for m in re.finditer(r'[!]{1,}|\?{2,}', p.text):
            hit('vp_cam_than', lb, p, m, ' → văn bản kỹ thuật không dùng dấu chấm than / "??"')
        for m in re.finditer(r'[\U0001F300-\U0001FAFF☀-➿]', p.text):
            hit('vp_cam_than', lb, p, m, ' → bỏ biểu tượng cảm xúc')

    # 2. câu, đoạn quá dài
    for lb, p in items:
        t = p.text.strip()
        if len(t) < 60:
            continue
        for s in re.split(r'(?<=[.?!;])\s+(?=[A-ZÀ-ỸĐ0-9(])', t):
            n = len(syllables(s))
            if n > 70:
                R['vp_cau_dai'].append(f'{lb} câu {n} âm tiết: "{s[:70]}…"') if len(R['vp_cau_dai']) < max_ex else None
        n = len(syllables(t))
        if n > 350:
            R['vp_doan_dai'].append(f'{lb} đoạn {n} âm tiết: "{t[:50]}…" → tách theo ý')

    # 3. lặp hư từ liền nhau
    rx = re.compile(r'(?<![\wÀ-ỹ])(' + '|'.join(HU_TU) + r')\s+\1(?![\wÀ-ỹ])', re.I)
    for lb, p in items:
        for m in rx.finditer(p.text):
            hit('vp_lap_tu', lb, p, m)

    # 4. dấu câu, khoảng trắng
    rules = [
        (r'[  ]+[,.;:?!)\]](?!\d)', 'thừa khoảng trắng trước dấu câu'),
        (r'[,;?!](?=[A-Za-zÀ-ỹĐđ])', 'thiếu khoảng trắng sau dấu câu'),
        (r'(?<=[A-Za-zÀ-ỹ]):(?=[A-Za-zÀ-ỹĐđ])', 'thiếu khoảng trắng sau dấu hai chấm'),
        (r'[(\[][  ]+', 'thừa khoảng trắng sau dấu mở ngoặc'),
        (r'(?<=\S) {2,}(?=\S)', 'nhiều khoảng trắng liền nhau'),
        (r'(?<![.\w])\.{2}(?!\.)', 'hai dấu chấm liền'),
    ]
    cnt = Counter()
    for pat, label in rules:
        for lb, p in items:
            if T.NGAY_RE.match(p.text):        # "ngày    tháng    năm": chỗ trống để điền ngày
                continue
            for m in re.finditer(pat, p.text):
                cnt[label] += 1
                if cnt[label] <= 3:
                    R['vp_dau_cau'].append(f'{label}: {lb} "{ctx(p.text, m, 15)}"')
    for label, n in cnt.items():
        if n > 3:
            R['vp_dau_cau'].append(f'… {label}: tổng {n} chỗ')
    q_straight = all_text.count('"')
    q_curly = all_text.count('“') + all_text.count('”')
    if q_straight and q_curly:
        R['vp_dau_cau'].append(f'trộn ngoặc kép thẳng (") ×{q_straight} và cong (“ ”) ×{q_curly}: dùng một kiểu')
    rng = Counter()
    for m in re.finditer(r'\d\s*([-–—])\s*\d', all_text):
        s = m.group(0)
        rng[('có cách' if ' ' in s else 'liền') + ' ' + m.group(1)] += 1
    if len(rng) > 1:
        R['vp_dau_cau'].append('khoảng giá trị ghi nhiều kiểu: ' + ', '.join(f'"{k}"×{v}' for k, v in rng.most_common())
                               + ' (khuyến nghị "6–12" hoặc "6 - 12", thống nhất)')

    # 5. số và đơn vị
    rx = re.compile(r'(?<![\w.,])(\d+(?:[.,]\d+)?)(' + UNITS + r')(?![\wÀ-ỹ/²³])')
    unit_attached = unit_spaced = 0
    for lb, p in items:
        for m in rx.finditer(p.text):
            unit_attached += 1
            if unit_attached <= 4:
                R['vp_don_vi'].append(f'{lb} "{ctx(p.text, m, 12)}" → "{m.group(1)} {m.group(2)}"')
        unit_spaced += len(re.findall(r'\d (?:' + UNITS + r')(?![\wÀ-ỹ])', p.text))
    if unit_attached > 4:
        R['vp_don_vi'].append(f'… tổng {unit_attached} chỗ số viết liền đơn vị (đã có {unit_spaced} chỗ viết cách)')

    # 6. viết tắt: chưa giải thích, chưa có trong danh mục
    abbr_rx = re.compile(r'(?<![\wÀ-ỹ])([A-ZĐ][A-ZĐ0-9]{1,6}(?:/[A-Z]{1,4})?)(?![\wÀ-ỹ])')
    uses = Counter()
    first = {}
    for lb, p in items:
        if p.text.strip().isupper() or T.SO_RE.match(p.text):   # tiêu đề in hoa, dòng "Số: 05/BC-…"
            continue
        for m in abbr_rx.finditer(p.text):
            a = m.group(1)
            if re.fullmatch(r'[IVXLC]+', a) or a.isdigit():
                continue
            uses[a] += 1
            first.setdefault(a, (lb, p.text, m))
    defined = set(re.findall(r'\(\s*([A-ZĐ][A-ZĐ0-9]{1,6}(?:/[A-Z]{1,4})?)\s*[),;]', all_text)) | \
        set(re.findall(r'([A-ZĐ][A-ZĐ0-9]{1,6})\s*[:(–-]\s*[a-zà-ỹA-Z]', all_text))
    # danh mục chữ viết tắt (bảng hoặc đoạn sau tiêu đề)
    listed = set()
    ps = list(T.iter_paragraphs(doc))
    for k, p in enumerate(ps):
        if re.match(r'^(DANH MỤC (CÁC )?(CHỮ|TỪ) VIẾT TẮT|BẢNG CHỮ VIẾT TẮT)', p.text.strip(), re.I):
            for q in ps[k + 1:k + 200]:
                t = q.text.strip()
                if re.match(r'^(MỤC LỤC|DANH MỤC|ĐẶT VẤN ĐỀ|MỞ ĐẦU|CHƯƠNG|Chương)', t):
                    break
                m = abbr_rx.match(t)
                if m:
                    listed.add(m.group(1))
            break
    whitelist = {'TP', 'UBND', 'HĐND', 'TW', 'VN', 'TS', 'BS', 'PGS', 'GS', 'ThS', 'CN', 'KT', 'TM', 'TL', 'TUQ',
                 'VT', 'STT', 'NXB', 'TCVN', 'ISO', 'PDF', 'URL', 'DOI', 'PMID', 'PMC', 'ISBN', 'BC', 'QĐ', 'NĐ', 'CP',
                 'TT', 'CV', 'KH', 'TTr', 'PHỤ', 'LỤC', 'MỤC', 'BẢNG', 'HÌNH'}
    vn_words = Counter(w.lower() for w in re.findall(r'[\wÀ-ỹ]+', all_text) if not w.isupper())
    for a in list(uses):
        if vn_words[a.lower()] >= 2 or not re.search(r'[A-Z]', a):   # "DANH", "SINH": từ in hoa, không phải viết tắt
            del uses[a]
    for a in list(uses):
        if '/' in a and all(x in defined for x in a.split('/')):
            defined.add(a)
    undef = [a for a, n in uses.most_common() if a not in defined and a not in whitelist
             and not first[a][1].strip().isupper()]
    for a in undef[:12]:
        lb, text, m = first[a]
        R['vp_viet_tat'].append(f'"{a}" ({uses[a]} lần) chưa giải thích lần đầu: {lb} "{ctx(text, m, 20)}"')
    if listed:
        miss = [a for a in uses if a not in listed and a not in whitelist and uses[a] >= 2]
        if miss:
            R['vp_viet_tat'].append('chưa có trong Danh mục chữ viết tắt: ' + ', '.join(sorted(miss)[:30]))
        unused = sorted(a for a in listed if uses[a] == 0)
        if unused:
            R['vp_viet_tat'].append('có trong danh mục nhưng không dùng trong bài: ' + ', '.join(unused[:30]))
    once = sorted(a for a, n in uses.items() if n == 1 and a not in whitelist and a in defined)
    if once:
        R['vp_viet_tat'].append('chỉ dùng 1 lần (nên viết đầy đủ, không cần viết tắt): ' + ', '.join(once[:20]))

    # 7. chính tả i/y, dấu thanh cũ/mới
    words = Counter(w for w in re.findall(r'[\wÀ-ỹ]+', all_text))
    lower = Counter()
    for w, n in words.items():
        lower[w.lower()] += n
    pairs = []
    for w, n in lower.items():
        m = re.fullmatch(r'([bcdđghklmnprstvx]+)([yýỳỷỹỵ])', w)
        if m and not w.startswith('qu'):
            alt = m.group(1) + Y_I[m.group(2)]
            if lower.get(alt):
                pairs.append(f'"{w}"×{n} / "{alt}"×{lower[alt]}')
    if pairs:
        R['vp_chinh_ta'].append('i/y không thống nhất: ' + ', '.join(sorted(pairs)[:15]))
    old = sum(n for w, n in lower.items() if tone_variant(w, True))
    new = sum(n for w, n in lower.items() if tone_variant(w, False))
    if old and new:
        ex_old = [w for w in lower if tone_variant(w, True)][:5]
        ex_new = [w for w in lower if tone_variant(w, False)][:5]
        R['vp_chinh_ta'].append(f'dấu thanh trộn kiểu cũ ×{old} ({", ".join(ex_old)}) và kiểu mới ×{new} '
                                f'({", ".join(ex_new)}): chạy fix_docx.py --dau-thanh moi')

    # 8. ngôi xưng
    m = re.search(r'^(ĐẶT VẤN ĐỀ|MỞ ĐẦU|LỜI MỞ ĐẦU|Chương 1|CHƯƠNG 1|Chương I|CHƯƠNG I|I\. )', all_text, re.M)
    body = all_text[m.start():] if m else all_text          # bỏ Lời cam đoan, Lời cảm ơn
    toi = len(re.findall(r'(?<![\wÀ-ỹ])tôi(?![\wÀ-ỹ])', body, re.I))
    ct = len(re.findall(r'(?<![\wÀ-ỹ])chúng tôi(?![\wÀ-ỹ])', body, re.I))
    if toi - ct > 0 and ct > 0:
        R['vp_ngoi'].append(f'phần nội dung trộn "tôi" ×{toi - ct} và "chúng tôi" ×{ct}: dùng thống nhất "chúng tôi" '
                            'hoặc câu khách quan ("nghiên cứu ghi nhận…"); "tôi" chỉ ở Lời cam đoan, Lời cảm ơn')

    for k in list(R):
        if k.startswith('_n_'):
            R[k[3:]].append(f'… và {R.pop(k)} chỗ khác')
    return R


# ------------------------------------------------------------------ sửa
def fix(doc, punct=True, units=False, tone=None):
    c = Counter()
    for lb, p in paragraphs(doc):
        if punct and not T.NGAY_RE.match(p.text):
            _fix_punct(p, c)
        if units:
            c['cách số và đơn vị'] += sub_runs(
                p, r'(?<![\w.,])(\d+(?:[.,]\d+)?)(' + UNITS + r')(?![\wÀ-ỹ/²³])', r'\1 \2')
        if tone:
            to_new = tone == 'moi'
            c[f'dấu thanh kiểu {"mới (hòa, thủy)" if to_new else "cũ (hoà, thuỷ)"}'] += sub_runs(
                p, r'(?<![\wÀ-ỹ])[\wÀ-ỹ]+(?![\wÀ-ỹ])', lambda m: tone_variant(m.group(0), to_new) or m.group(0))
    c = {k: v for k, v in c.items() if v}
    log.append('văn phong: ' + (', '.join(f'{k} {v}' for k, v in c.items()) if c else 'không có lỗi trình bày cần sửa'))


def _fix_punct(p, c):
    c['thừa khoảng trắng trước dấu câu'] += sub_runs(p, r'[  ]+([,.;:?!)\]])(?!\d)', r'\1')
    c['thêm khoảng trắng sau dấu câu'] += sub_runs(p, r'([,;?!])(?=[A-Za-zÀ-ỹĐđ])', r'\1 ')
    c['thêm khoảng trắng sau dấu hai chấm'] += sub_runs(p, r'(?<=[A-Za-zÀ-ỹ]):(?=[A-Za-zÀ-ỹĐđ])', ': ')
    c['bỏ khoảng trắng sau mở ngoặc'] += sub_runs(p, r'([(\[])[  ]+', r'\1')
    c['gộp khoảng trắng'] += sub_runs(p, r'(?<=\S) {2,}(?=\S)', ' ')
    c['v.v... → v.v.'] += sub_runs(p, r'v\.v(\.{2,}|…)', 'v.v.')
    c['± có cách hai bên'] += sub_runs(p, r'(?<=\d)(\s*)±(\s*)(?=\d)',
                                      lambda m: m.group(0) if m.group(1) and m.group(2) else ' ± ')
    c['bỏ hư từ lặp'] += sub_runs(p, r'(?<![\wÀ-ỹ])(' + '|'.join(HU_TU) + r')(\s+\1)(?![\wÀ-ỹ])', r'\1',
                                  re.I)
