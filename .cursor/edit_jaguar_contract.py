#!/usr/bin/env python3
"""Apply tracked changes and comments to ジャグー業務委託契約書 _2 file."""
from __future__ import annotations
import shutil
import zipfile
from copy import deepcopy
from pathlib import Path
from lxml import etree

# ---------- Paths ----------
SRC_ORIG = Path('/Users/kometaninaoki/Library/CloudStorage/GoogleDrive-n.kometani@re-startlaw.com/マイドライブ/共有用/01_事件記録/だ_大研バイオメディカル/260317_ジャグー株式会社/260427_【大研バイオメディカル株式会社御中】業務委託契約書_0423_米谷チェック済み.docx')
DST = SRC_ORIG.with_name(SRC_ORIG.stem + '_2.docx')

# ---------- Re-copy from original ----------
shutil.copy2(SRC_ORIG, DST)

# ---------- Constants ----------
W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
W14_NS = 'http://schemas.microsoft.com/office/word/2010/wordml'
W15_NS = 'http://schemas.microsoft.com/office/word/2012/wordml'
W16CID_NS = 'http://schemas.microsoft.com/office/word/2016/wordml/cid'
W16DU_NS = 'http://schemas.microsoft.com/office/word/2023/wordml/word16du'
NSMAP = {'w': W_NS, 'w14': W14_NS, 'w15': W15_NS, 'w16cid': W16CID_NS, 'w16du': W16DU_NS}

W = lambda t: f'{{{W_NS}}}{t}'
W14 = lambda t: f'{{{W14_NS}}}{t}'
W15 = lambda t: f'{{{W15_NS}}}{t}'
W16CID = lambda t: f'{{{W16CID_NS}}}{t}'

AUTHOR = '米谷尚起'
INITIALS = '米谷'
DATE = '2026-04-27T12:00:00Z'
DATE_UTC = '2026-04-27T03:00:00Z'

# Counters
_id_counter = [500]
_comment_id_counter = [300]
_paraid_counter = [0xA0000000]


def next_id() -> str:
    v = _id_counter[0]
    _id_counter[0] += 1
    return str(v)


def next_comment_id() -> str:
    v = _comment_id_counter[0]
    _comment_id_counter[0] += 1
    return str(v)


def next_paraid() -> str:
    _paraid_counter[0] += 1
    return f'{_paraid_counter[0]:08X}'


# ---------- Run helpers ----------
def make_rpr(font='ＭＳ 明朝', highlight=None, color=None):
    rpr = etree.Element(W('rPr'))
    rfonts = etree.SubElement(rpr, W('rFonts'))
    rfonts.set(W('ascii'), font)
    rfonts.set(W('hAnsi'), font)
    rfonts.set(W('hint'), 'eastAsia')
    if color:
        c = etree.SubElement(rpr, W('color'))
        c.set(W('val'), color)
    if highlight:
        h = etree.SubElement(rpr, W('highlight'))
        h.set(W('val'), highlight)
    sz = etree.SubElement(rpr, W('szCs'))
    sz.set(W('val'), '21')
    return rpr


def make_run(text: str, rpr_extra: dict | None = None):
    r = etree.Element(W('r'))
    rpr = make_rpr(**(rpr_extra or {}))
    r.append(rpr)
    t = etree.SubElement(r, W('t'))
    t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    t.text = text
    return r


def make_ins(child_runs):
    ins = etree.Element(W('ins'))
    ins.set(W('id'), next_id())
    ins.set(W('author'), AUTHOR)
    ins.set(W('date'), DATE)
    for r in child_runs:
        ins.append(r)
    return ins


def wrap_runs_in_del(runs_to_delete: list):
    """Wrap a list of <w:r> elements in <w:del> and convert their <w:t> to <w:delText>."""
    if not runs_to_delete:
        return None
    parent = runs_to_delete[0].getparent()
    idx = list(parent).index(runs_to_delete[0])
    del_elem = etree.Element(W('del'))
    del_elem.set(W('id'), next_id())
    del_elem.set(W('author'), AUTHOR)
    del_elem.set(W('date'), DATE)
    for r in runs_to_delete:
        # Convert w:t to w:delText
        for t in r.findall(W('t')):
            t.tag = W('delText')
        parent.remove(r)
        del_elem.append(r)
    parent.insert(idx, del_elem)
    return del_elem


def make_paragraph(runs, ind_first_line=False):
    p = etree.Element(W('p'))
    paraid = next_paraid()
    p.set(W14('paraId'), paraid)
    p.set(W14('textId'), '00000000')
    ppr = etree.SubElement(p, W('pPr'))
    if ind_first_line:
        ind = etree.SubElement(ppr, W('ind'))
        ind.set(W('firstLineChars'), '100')
        ind.set(W('firstLine'), '210')
    rpr_in_ppr = etree.SubElement(ppr, W('rPr'))
    ins_marker = etree.SubElement(rpr_in_ppr, W('ins'))
    ins_marker.set(W('id'), next_id())
    ins_marker.set(W('author'), AUTHOR)
    ins_marker.set(W('date'), DATE)
    rfonts = etree.SubElement(rpr_in_ppr, W('rFonts'))
    rfonts.set(W('ascii'), 'ＭＳ 明朝')
    rfonts.set(W('hAnsi'), 'ＭＳ 明朝')
    sz = etree.SubElement(rpr_in_ppr, W('szCs'))
    sz.set(W('val'), '21')
    p.append(make_ins(runs))
    return p, paraid


# ---------- Comment helpers ----------
def make_comment_xml(comment_id: str, text_lines: list[str]) -> tuple[etree._Element, str]:
    """Create a <w:comment> element. Returns (element, paraId of first paragraph)."""
    c = etree.Element(W('comment'))
    c.set(W('id'), comment_id)
    c.set(W('author'), AUTHOR)
    c.set(W('date'), DATE)
    c.set(W('initials'), INITIALS)
    first_paraid = None
    for i, line in enumerate(text_lines):
        p = etree.SubElement(c, W('p'))
        paraid = next_paraid()
        if i == 0:
            first_paraid = paraid
        p.set(W14('paraId'), paraid)
        p.set(W14('textId'), '77777777')
        ppr = etree.SubElement(p, W('pPr'))
        pstyle = etree.SubElement(ppr, W('pStyle'))
        pstyle.set(W('val'), 'ac')
        if i == 0:
            r = etree.SubElement(p, W('r'))
            rpr = etree.SubElement(r, W('rPr'))
            rstyle = etree.SubElement(rpr, W('rStyle'))
            rstyle.set(W('val'), 'ab')
            etree.SubElement(r, W('annotationRef'))
        if line:
            r = etree.SubElement(p, W('r'))
            rpr = etree.SubElement(r, W('rPr'))
            rfonts = etree.SubElement(rpr, W('rFonts'))
            rfonts.set(W('hint'), 'eastAsia')
            t = etree.SubElement(r, W('t'))
            t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
            t.text = line
    return c, first_paraid


def insert_comment_anchor_after(target_run: etree._Element, comment_id: str):
    """Insert commentRangeStart, commentRangeEnd, and commentReference after target_run."""
    parent = target_run.getparent()
    idx = list(parent).index(target_run)
    crs = etree.Element(W('commentRangeStart'))
    crs.set(W('id'), comment_id)
    cre = etree.Element(W('commentRangeEnd'))
    cre.set(W('id'), comment_id)
    cref_r = etree.Element(W('r'))
    cref_rpr = etree.SubElement(cref_r, W('rPr'))
    cref_style = etree.SubElement(cref_rpr, W('rStyle'))
    cref_style.set(W('val'), 'ab')
    cref = etree.SubElement(cref_r, W('commentReference'))
    cref.set(W('id'), comment_id)
    parent.insert(idx, crs)
    parent.insert(idx + 2, cre)
    parent.insert(idx + 3, cref_r)


# ---------- Load XML ----------
def load_xml_from_zip(zf, name) -> etree._Element:
    return etree.fromstring(zf.read(name))


with zipfile.ZipFile(DST) as zf:
    document_xml = load_xml_from_zip(zf, 'word/document.xml')
    comments_xml = load_xml_from_zip(zf, 'word/comments.xml')
    comments_ext_xml = load_xml_from_zip(zf, 'word/commentsExtended.xml')
    comments_ids_xml = load_xml_from_zip(zf, 'word/commentsIds.xml')


# ---------- Helper: find run by exact text ----------
def find_run_by_text(root, exact_text: str) -> etree._Element | None:
    for t in root.iter(W('t')):
        if t.text == exact_text:
            return t.getparent()  # the <w:r>
    return None


def find_runs_containing(root, substring: str) -> list[etree._Element]:
    out = []
    for t in root.iter(W('t')):
        if t.text and substring in t.text:
            out.append(t.getparent())
    return out


comments_added = []  # (comment_id, text, anchor_paraid)


# ============================================================
# Edit 1: 第3条4項「第◯条」→「第8条」
# ============================================================
print('Edit 1: 第3条4項 第◯条 → 第8条')
target = None
for t in document_xml.iter(W('t')):
    if t.text == '第◯条の定める検':
        target = t.getparent()
        break
assert target is not None, '第◯条の定める検 not found'

parent = target.getparent()
idx = list(parent).index(target)
# Replace with: w:r("第") + w:del("◯") + w:ins("８") + w:r("条の定める検")
rpr_template = target.find(W('rPr'))


def clone_rpr():
    if rpr_template is not None:
        return deepcopy(rpr_template)
    return None


def make_run_with_existing_rpr(text):
    r = etree.Element(W('r'))
    rpr = clone_rpr()
    if rpr is not None:
        r.append(rpr)
    t = etree.SubElement(r, W('t'))
    t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    t.text = text
    return r


# Build replacement
r_pre = make_run_with_existing_rpr('第')
del_elem = etree.Element(W('del'))
del_elem.set(W('id'), next_id())
del_elem.set(W('author'), AUTHOR)
del_elem.set(W('date'), DATE)
r_del = etree.Element(W('r'))
rpr2 = clone_rpr()
if rpr2 is not None:
    r_del.append(rpr2)
dt = etree.SubElement(r_del, W('delText'))
dt.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
dt.text = '◯'
del_elem.append(r_del)

ins_elem = etree.Element(W('ins'))
ins_elem.set(W('id'), next_id())
ins_elem.set(W('author'), AUTHOR)
ins_elem.set(W('date'), DATE)
r_ins = etree.Element(W('r'))
rpr3 = clone_rpr()
if rpr3 is not None:
    r_ins.append(rpr3)
it = etree.SubElement(r_ins, W('t'))
it.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
it.text = '８'
ins_elem.append(r_ins)

r_post = make_run_with_existing_rpr('条の定める検')

parent.remove(target)
parent.insert(idx, r_pre)
parent.insert(idx + 1, del_elem)
parent.insert(idx + 2, ins_elem)
parent.insert(idx + 3, r_post)

# ============================================================
# Edit 7 (do early before paragraph indices shift): 第18条 存続条項 修正
# Replace 第８条 → 第９条, 第１０条 → 第１１条, 第１５条 → 第１６条, 第１９条 → 第２０条
# ============================================================
print('Edit 7: 第18条 存続条項 番号修正')

# 第18条 ends with "本契約の解除又は期間満了後においても、第" + "８" + "条、" + "第１０条" + "、第" + "１５" + "条、第" + "１９" + "条..."
# Find each numeric run and replace.

def replace_run_text_with_tracked(root, old_text: str, new_text: str, occurrence: int = 1):
    """Find Nth occurrence of run with exact text == old_text, replace via w:del + w:ins."""
    matches = []
    for t in root.iter(W('t')):
        if t.text == old_text:
            matches.append(t.getparent())
    assert len(matches) >= occurrence, f'not enough matches for {old_text!r}: {len(matches)}'
    target = matches[occurrence - 1]
    par = target.getparent()
    i = list(par).index(target)
    rpr_t = target.find(W('rPr'))

    def clone():
        return deepcopy(rpr_t) if rpr_t is not None else None

    de = etree.Element(W('del'))
    de.set(W('id'), next_id())
    de.set(W('author'), AUTHOR)
    de.set(W('date'), DATE)
    rd = etree.Element(W('r'))
    rprx = clone()
    if rprx is not None:
        rd.append(rprx)
    dx = etree.SubElement(rd, W('delText'))
    dx.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    dx.text = old_text
    de.append(rd)

    ie = etree.Element(W('ins'))
    ie.set(W('id'), next_id())
    ie.set(W('author'), AUTHOR)
    ie.set(W('date'), DATE)
    ri = etree.Element(W('r'))
    rpry = clone()
    if rpry is not None:
        ri.append(rpry)
    ix = etree.SubElement(ri, W('t'))
    ix.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    ix.text = new_text
    ie.append(ri)

    par.remove(target)
    par.insert(i, de)
    par.insert(i + 1, ie)
    return target


# In document, the 存続条項 paragraph is at line ~3765 area. Use occurrence indexing.
# "８" appears multiple times in document. We need to find the one in 存続条項.
# Strategy: find paragraph containing "存続条項" structure, then operate within it.

surv_para = None
for p in document_xml.iter(W('p')):
    text = ''.join(t.text or '' for t in p.iter(W('t')))
    if '本契約の解除又は期間満了後においても、第' in text and 'の各条項は、それぞれに定めるところに従い引き続き有効とする' in text:
        surv_para = p
        break
assert surv_para is not None, '存続条項 paragraph not found'


def replace_in_paragraph_run(p, old_text: str, new_text: str):
    target = None
    for t in p.iter(W('t')):
        if t.text == old_text:
            target = t.getparent()
            break
    if target is None:
        # try contains (run might not be exact match)
        for t in p.iter(W('t')):
            if t.text and t.text.strip() == old_text:
                target = t.getparent()
                break
    assert target is not None, f'in surv_para: {old_text!r} not found'
    par = target.getparent()
    i = list(par).index(target)
    rpr_t = target.find(W('rPr'))

    def clone():
        return deepcopy(rpr_t) if rpr_t is not None else None

    de = etree.Element(W('del'))
    de.set(W('id'), next_id())
    de.set(W('author'), AUTHOR)
    de.set(W('date'), DATE)
    rd = etree.Element(W('r'))
    rprx = clone()
    if rprx is not None:
        rd.append(rprx)
    dx = etree.SubElement(rd, W('delText'))
    dx.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    dx.text = old_text
    de.append(rd)

    ie = etree.Element(W('ins'))
    ie.set(W('id'), next_id())
    ie.set(W('author'), AUTHOR)
    ie.set(W('date'), DATE)
    ri = etree.Element(W('r'))
    rpry = clone()
    if rpry is not None:
        ri.append(rpry)
    ix = etree.SubElement(ri, W('t'))
    ix.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    ix.text = new_text
    ie.append(ri)

    par.remove(target)
    par.insert(i, de)
    par.insert(i + 1, ie)


replace_in_paragraph_run(surv_para, '８', '９')
replace_in_paragraph_run(surv_para, '第１０条', '第１１条')
replace_in_paragraph_run(surv_para, '１５', '１６')
replace_in_paragraph_run(surv_para, '１９', '２０')


# ============================================================
# Edit 4: 第9条1項但書 — 利用許諾追加
# Find paragraph containing "等のデータを使用した場合の元データ" — that's the 但書 paragraph (inside <w:ins>).
# Insert new paragraph AFTER it with 案1 利用許諾 text, as tracked insertion by 米谷.
# ============================================================
print('Edit 4: 第9条1項 利用許諾追加')

target_p = None
for p in document_xml.iter(W('p')):
    text = ''.join(t.text or '' for t in p.iter(W('t')))
    if '等のデータを使用した場合の元データ' in text and '乙から甲に移転しないものとする。' in text:
        target_p = p
        break
assert target_p is not None, '但書 paragraph not found'

new_text_4_first = '　ただし、甲は、納品された成果物に組み込まれている範囲内において、乙が保有する元データ及び第三者ライセンス素材を、本契約の期間中および契約終了後も、無償かつ無期限で、本件業務目的を超えて自由に利用、複製、改変、譲渡することができる。'
new_text_4_second = '　なお、第三者ライセンス素材については、乙が当該第三者から取得しているライセンス範囲内に限るものとし、乙は契約締結時に当該ライセンス内容および制限事項を甲に書面で開示するものとする。'

new_p_4a, paraid_4a = make_paragraph([make_run(new_text_4_first)])
new_p_4b, paraid_4b = make_paragraph([make_run(new_text_4_second)])

# Insert after target_p
parent_p = target_p.getparent()
idx_p = list(parent_p).index(target_p)
parent_p.insert(idx_p + 1, new_p_4a)
parent_p.insert(idx_p + 2, new_p_4b)


# ============================================================
# Edit 5: 第9条2項本文削除
# Find three paragraphs and wrap each <w:r> within their <w:ins> in <w:del>
# ============================================================
print('Edit 5: 第9条2項本文削除')

texts_to_delete = [
    '２．前条第１項の成果物は、甲に限りその利用が許諾されるものとし、甲は乙の承諾なしに成果物を公表',
    '又は第三者に伝達することはできないものとする。',
    '　なお，甲は，前項但書の成果物の制作に使用した素材について，本契約の目的の範囲内でのみ利用できるものとする。',
]

for tx in texts_to_delete:
    found = False
    for t in document_xml.iter(W('t')):
        if t.text == tx:
            r = t.getparent()
            ins_parent = r.getparent()
            if ins_parent.tag != W('ins'):
                continue
            # Wrap r with w:del
            i = list(ins_parent).index(r)
            de = etree.Element(W('del'))
            de.set(W('id'), next_id())
            de.set(W('author'), AUTHOR)
            de.set(W('date'), DATE)
            ins_parent.remove(r)
            # convert w:t -> w:delText
            for tt in r.findall(W('t')):
                tt.tag = W('delText')
            de.append(r)
            ins_parent.insert(i, de)
            found = True
            break
    assert found, f'not found for delete: {tx!r}'


# ============================================================
# Edit 6: 第11条 機密保持 6/7/8項追加
# Insert after the paragraph containing "する。" that ends 第11条 5項.
# That's the para containing "する。" wrapped in <w:ins> by 田中, around line 2231.
# Identify it as the para with text content ending in "する。" that comes after "５．甲および乙が、法令"
# ============================================================
print('Edit 6: 第11条 機密保持 6/7/8項追加')

# Strategy: find the paragraph that contains the run "する。" inside <w:ins> just after the 5項 sequence
# Easier: walk paragraphs looking for "５．甲および乙が、法令" then find the next 段落 that ends with "する。"
target_5_para = None
paragraphs = list(document_xml.iter(W('p')))
for i, p in enumerate(paragraphs):
    text = ''.join(t.text or '' for t in p.iter(W('t')))
    if '５．甲および乙が、法令、官公庁または裁判所の処分・命令等により機密情報または個人情報の開示要求' in text:
        # The 5項 spans multiple paragraphs ending with one containing only "する。"
        # Look ahead for the paragraph containing exactly "する。" text only
        for j in range(i + 1, min(i + 8, len(paragraphs))):
            t_join = ''.join(tt.text or '' for tt in paragraphs[j].iter(W('t')))
            if t_join.strip() == 'する。':
                target_5_para = paragraphs[j]
                break
        break
assert target_5_para is not None, '5項 ending paragraph not found'

# Build new paragraphs
text_6 = '６．甲および乙は、本契約が終了した場合、又は相手方から要求があった場合は、理由の如何を問わず、速やかに受領した機密情報および個人情報（記録媒体、複製物等を含む。）を相手方の指示に従い返還、又は再生不能な状態で破棄若しくは消去しなければならない。この場合において、相手方から要求があったときは、情報を破棄又は消去した旨を証する書面を提出するものとする。'
text_7 = '７．乙は、機密情報又は個人情報の漏洩、滅失、毀損、改ざん又は不正アクセス等の事故（それらのおそれがある場合を含む。）が発生したことを知ったときは、直ちにその旨を甲に報告し、甲の指示に従わなければならない。この場合において、乙は、自己の責任と負担により、速やかに原因の究明および被害の拡大防止のための適切な措置を講じるとともに、事故の詳細な報告および再発防止策を書面により甲に提出し、その承認を得なければならない。'
text_8 = '８．前項の事故に起因して、甲が第三者（顧客を含むがこれに限られない。）から損害賠償請求等の異議申立てを受けた場合、又は甲に顧客対応費用、事実調査費用その他の損害（合理的な弁護士費用を含む。）が生じた場合、乙は、第１６条の規定にかかわらず、自己の責任と負担においてこれを解決し、甲が被った一切の損害を賠償しなければならない。'

new_p_6, _ = make_paragraph([make_run(text_6)])
new_p_7, _ = make_paragraph([make_run(text_7)])
new_p_8, _ = make_paragraph([make_run(text_8)])

parent_5 = target_5_para.getparent()
idx_5 = list(parent_5).index(target_5_para)
parent_5.insert(idx_5 + 1, new_p_6)
parent_5.insert(idx_5 + 2, new_p_7)
parent_5.insert(idx_5 + 3, new_p_8)


# ============================================================
# Edit 2: 第4条 月次レポート義務 — 提案挿入 + 依頼者向けコメント
# Insert new paragraph BEFORE 第4条第2項 ("本件業務における目標数値") with proposal text.
# ============================================================
print('Edit 2: 第4条 月次レポート提案 + コメント')

target_2 = None
for p in document_xml.iter(W('p')):
    text = ''.join(t.text or '' for t in p.iter(W('t')))
    if '本件業務における目標数値（KPI）' in text:
        target_2 = p
        break
assert target_2 is not None, '第4条2項 not found'

text_monthly = '２．乙は、毎月末日までに、当該月の業務遂行状況、成果、稼働時間実績および翌月の業務計画を記載した月次レポートを作成し、翌月５営業日以内に甲に提出する。'
new_p_monthly, paraid_monthly = make_paragraph([make_run(text_monthly)])

# We'll attach a comment anchored on this paragraph's run
comment_id_2 = next_comment_id()

parent_2 = target_2.getparent()
idx_2 = list(parent_2).index(target_2)
parent_2.insert(idx_2, new_p_monthly)

# Insert comment anchor: wrap the inserted run with commentRangeStart/End and reference
ins_in_new_p = new_p_monthly.find(W('ins'))
first_r = ins_in_new_p[0]
# Insert commentRangeStart before, End after, then Reference
i_in_ins = list(ins_in_new_p).index(first_r)
crs = etree.Element(W('commentRangeStart'))
crs.set(W('id'), comment_id_2)
cre = etree.Element(W('commentRangeEnd'))
cre.set(W('id'), comment_id_2)
cref_r = etree.Element(W('r'))
cref_rpr = etree.SubElement(cref_r, W('rPr'))
cref_style = etree.SubElement(cref_rpr, W('rStyle'))
cref_style.set(W('val'), 'ab')
cref = etree.SubElement(cref_r, W('commentReference'))
cref.set(W('id'), comment_id_2)
ins_in_new_p.insert(i_in_ins, crs)
ins_in_new_p.insert(i_in_ins + 2, cre)
ins_in_new_p.insert(i_in_ins + 3, cref_r)

comment_2_lines = [
    '【依頼者ご確認】月次レポート義務の復活を提案しています。前回米谷案にあった「月次の詳細報告」がジャグー側で削除されているため、業務状況の月次把握のため復活を要請する内容です。提出期限を「翌月5営業日以内」としていますが、御社のご希望に応じて調整可能です（例：翌月初日、翌月10日以内等）。不要であればこのコメントごと削除してください。',
]
c_elem_2, c_paraid_2 = make_comment_xml(comment_id_2, comment_2_lines)
comments_added.append((comment_id_2, c_elem_2, c_paraid_2))


# ============================================================
# Edit 3: 第8条1項 納品期日 — 依頼者向けコメント
# Anchor: the run "ィブ成果物を甲および乙が合意する納品期日までに納品する。"
# ============================================================
print('Edit 3: 第8条1項 納品期日コメント')

target_3 = None
for t in document_xml.iter(W('t')):
    if t.text == 'ィブ成果物を甲および乙が合意する納品期日までに納品する。':
        target_3 = t.getparent()
        break
assert target_3 is not None, '第8条1項 anchor not found'

comment_id_3 = next_comment_id()
# target_3 is inside <w:ins>. Insert range markers around it.
ins_parent_3 = target_3.getparent()
i3 = list(ins_parent_3).index(target_3)

crs3 = etree.Element(W('commentRangeStart'))
crs3.set(W('id'), comment_id_3)
cre3 = etree.Element(W('commentRangeEnd'))
cre3.set(W('id'), comment_id_3)
cref_r3 = etree.Element(W('r'))
cref_rpr3 = etree.SubElement(cref_r3, W('rPr'))
cref_style3 = etree.SubElement(cref_rpr3, W('rStyle'))
cref_style3.set(W('val'), 'ab')
cref3 = etree.SubElement(cref_r3, W('commentReference'))
cref3.set(W('id'), comment_id_3)

ins_parent_3.insert(i3, crs3)
ins_parent_3.insert(i3 + 2, cre3)
ins_parent_3.insert(i3 + 3, cref_r3)

comment_3_lines = [
    '【依頼者ご確認】第8条1項の修正点について2点ご判断ください。①納品期日：ジャグー側修正で「甲および乙が合意する納品期日」となっています（甲が一方的に指定する形ではない）。御社として「甲指定」に戻すか、「甲乙合意」を許容するかご判断ください。②納品物範囲：「画像またはバナー等のクリエイティブ成果物」と限定されていますが、コンサルレポート等も検収対象としたい場合は「本件業務にかかる成果物」へ変更要請が可能です。',
]
c_elem_3, c_paraid_3 = make_comment_xml(comment_id_3, comment_3_lines)
comments_added.append((comment_id_3, c_elem_3, c_paraid_3))


# ============================================================
# Edit 8: 代表者名 黄色ハイライト + コメント
# Find run with text "林 東慶" — replace its rPr to add highlight, and attach comment
# ============================================================
print('Edit 8: 代表者名 黄色ハイライト + コメント')

target_8 = None
for t in document_xml.iter(W('t')):
    if t.text and '林' in t.text and '東慶' in t.text:
        target_8 = t.getparent()
        break
assert target_8 is not None, '林 東慶 not found'

# Add highlight to rPr
rpr_8 = target_8.find(W('rPr'))
if rpr_8 is None:
    rpr_8 = etree.SubElement(target_8, W('rPr'))
    target_8.insert(0, rpr_8)
# Remove existing highlight if any
for h in rpr_8.findall(W('highlight')):
    rpr_8.remove(h)
hh = etree.SubElement(rpr_8, W('highlight'))
hh.set(W('val'), 'yellow')

comment_id_8 = next_comment_id()
parent_8 = target_8.getparent()
i8 = list(parent_8).index(target_8)

crs8 = etree.Element(W('commentRangeStart'))
crs8.set(W('id'), comment_id_8)
cre8 = etree.Element(W('commentRangeEnd'))
cre8.set(W('id'), comment_id_8)
cref_r8 = etree.Element(W('r'))
cref_rpr8 = etree.SubElement(cref_r8, W('rPr'))
cref_style8 = etree.SubElement(cref_rpr8, W('rStyle'))
cref_style8.set(W('val'), 'ab')
cref8 = etree.SubElement(cref_r8, W('commentReference'))
cref8.set(W('id'), comment_id_8)

parent_8.insert(i8, crs8)
parent_8.insert(i8 + 2, cre8)
parent_8.insert(i8 + 3, cref_r8)

comment_8_lines = [
    '【依頼者ご確認】大研バイオメディカル代表者名のご確認をお願いします。「林 東慶」と記載されていますが、「黄 楊」との表記もあるとのことです。全部事項証明書（240325取得分）に基づき、契約締結時点での正規の代表者名をご確認・ご指示ください。',
]
c_elem_8, c_paraid_8 = make_comment_xml(comment_id_8, comment_8_lines)
comments_added.append((comment_id_8, c_elem_8, c_paraid_8))


# ============================================================
# Append new comments to comments.xml, commentsExtended.xml, commentsIds.xml
# ============================================================
print('Appending comments to comments.xml')
for cid, c_elem, c_paraid in comments_added:
    comments_xml.append(c_elem)
    # commentsExtended
    cex = etree.SubElement(comments_ext_xml, W15('commentEx'))
    cex.set(W15('paraId'), c_paraid)
    cex.set(W15('done'), '0')
    # commentsIds: random 8-hex
    import secrets
    durable = secrets.token_hex(4).upper()
    cidn = etree.SubElement(comments_ids_xml, W16CID('commentId'))
    cidn.set(W16CID('paraId'), c_paraid)
    cidn.set(W16CID('durableId'), durable)


# ============================================================
# Save
# ============================================================
print(f'Writing back to {DST}')


def serialize(root):
    return etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)


# Read all entries, then write a new zip
import io
old = zipfile.ZipFile(DST)
buf = io.BytesIO()
with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as nz:
    for item in old.infolist():
        if item.filename == 'word/document.xml':
            nz.writestr(item, serialize(document_xml))
        elif item.filename == 'word/comments.xml':
            nz.writestr(item, serialize(comments_xml))
        elif item.filename == 'word/commentsExtended.xml':
            nz.writestr(item, serialize(comments_ext_xml))
        elif item.filename == 'word/commentsIds.xml':
            nz.writestr(item, serialize(comments_ids_xml))
        else:
            nz.writestr(item, old.read(item.filename))
old.close()
DST.write_bytes(buf.getvalue())
print('Done.')
