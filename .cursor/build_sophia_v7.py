"""v6→v7契約書生成。
- 第2条2項・第2条4項・第6条1項・第10条2項を指示通りに置換。
- v6の既存コメントは全消去。
- v2と異なる設計の主要条項に新規コメントを付与。
- 出力: /Users/kometaninaoki/Downloads/合弁会社設立契約書_株式会社Sophia_v7.docx
"""
from __future__ import annotations
import shutil
import zipfile
from pathlib import Path
from copy import deepcopy
from lxml import etree
from docx import Document

SRC = Path("/Users/kometaninaoki/Downloads/合弁会社設立契約書_株式会社Sophia_v6.docx")
DST = Path("/Users/kometaninaoki/Downloads/合弁会社設立契約書_株式会社Sophia_v7.docx")

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = f"{{{W_NS}}}"

PARA2_2_NEW = (
    "２．本会社の取締役の員数は、原則として2名以上の偶数とし、甲及び乙はそれぞれ"
    "同数の取締役を指名する権限を有する。本会社設立時の取締役は、甲指名の取締役"
    "として徳永秀和、乙指名の取締役として福井瑞希とする。"
)
PARA2_4_NEW = (
    "４．甲及び乙は、自らが指名した取締役についてのみ、その解任を決定することが"
    "できる。甲又は乙が当該解任を決定した場合、相手方は、当該解任に係る株主総会"
    "の議案に賛成の議決権を行使する義務を負う。"
)
PARA6_1_NEW = (
    "１．本会社における以下の各号に定める事項の決定（以下「重要事項」と総称する。）"
    "については、株主総会の決議によるものとし、当該決議は、議決権を行使すること"
    "ができる株主の議決権の過半数を有する株主が出席し、出席した当該株主の議決権"
    "の4分の3超に当たる多数をもって行うものとする。"
)
PARA10_2_NEW = (
    "２．配当を行うか否か、及びその配当額については、甲及び乙が協議の上、株主総会"
    "の決議によって決定する。"
)

# v6本文段落の旧テキスト判定（先頭一致でOK）
REPLACE_MAP = {
    # 第2条第2項
    "２．本会社の取締役の員数は、原則として2名以上の偶数とし、甲及び乙はそれぞれ同数の取締役を指名する権限を有する。本会社設立時の取締役は、甲指名の取締役として徳永秀和、乙指名の取締役として福井瑞希とする。なお、設立時及びその後の改選時における": PARA2_2_NEW,
    # 第2条第4項
    "４．甲又は乙は、取締役の解任を決定することができる。ただし、解任は、甲指名の取締役及び乙指名の取締役について同数ずつ同時に行わなければならず、": PARA2_4_NEW,
    # 第6条第1項
    "１．本会社における以下の各号に定める事項の決定（以下「重要事項」と総称する。）については、甲及び乙双方の事前の書面による同意を要するものとする。": PARA6_1_NEW,
    # 第10条第2項
    "２．配当を行うか否か、及びその配当額については、甲及び乙が協議の上、取締役会の全員一致の決議によって決定する。": PARA10_2_NEW,
}

# ────────────────────────────────────────────────────────────
# 新規コメント定義
# anchor_substr: 一意に特定できる段落本文の先頭部分文字列
# body: コメント本文（段落区切りで配列）
# ────────────────────────────────────────────────────────────
NEW_COMMENTS = [
    {
        "anchor_substr": "２．本会社の取締役の員数は、原則として2名以上の偶数とし",
        "title": "【第2条（取締役の選解任）について】",
        "body": [
            "v2では取締役の選任のみ「同数同時の原則」を定め、解任に関する規定が明示されていませんでした。",
            "v7では、会社法上の各株主の権利を尊重し、自己が指名した取締役は単独で解任できる仕組みに整理しています（第4項）。",
            "これにより、相手方取締役と抱合せで解任せざるを得ないという実務上の不都合（自社人事の硬直化）を解消する目的です。",
        ],
    },
    {
        "anchor_substr": "１．本会社における以下の各号に定める事項の決定（以下「重要事項」と総称する。）",
        "title": "【第6条（重要事項の決定）について】",
        "body": [
            "v2は第4条として「重要事項」の枠だけが置かれ、決定要件・対象事項の中身が定まっていませんでした。",
            "v7では会社法上の機関決定として機能させるため、株主総会の特別決議（出席要件・賛成要件加重）に揃え、別紙定款にも同旨を置く設計としています。",
            "単なる「甲乙の書面同意」では会社法上の機関決定にならず、後の登記・対外関係で支障が生じ得るため、契約条項と機関決定要件を一致させる方針に変更したものです。",
        ],
    },
    {
        "anchor_substr": "２．配当を行うか否か、及びその配当額については、甲及び乙が協議の上、株主総会",
        "title": "【第10条（剰余金の配当）について】",
        "body": [
            "v2は「取締役会の全員一致の決議」を要件としていましたが、会社法上、剰余金の配当は原則として株主総会決議事項であり、本契約のような株主構成（甲乙各50％）では取締役会への委任が機能しないため、v7では株主総会決議に置き換えました。",
            "また、v2には「甲及び乙以外の第三者が株主となった場合、当該第三者には配当を行わない」旨の規定が置かれていました（合弁会社性の維持を意図したもの）。",
            "もっとも、株主平等原則（会社法109条）との関係で当該規定の効力は否定されやすく、後日の紛争原因となり得るため、v7ではこの規定を採用していません。",
            "意図する効果（第三者株主の混入回避）は、第13条の譲渡制限・先買権及び第14条・第15条のコール／プットオプションによってリカバーする二段構えで実現する設計です。",
        ],
    },
    {
        "anchor_substr": "３．将来、本会社を第三者に譲渡する場合（いわゆるM&A）に備え",
        "title": "【第11条（解散時の知的財産の処理）について】",
        "body": [
            "v2第7条第3項は、解散時に知的財産権・顧客リストの利用権を甲乙の共有とすることを定めていました。",
            "しかし、共有とした場合、特許法73条・著作権法65条等により、他の共有者の同意なくしては第三者へのライセンスや持分の処分ができず、解散後に別方向に進む両当事者から同意を取り続けることが必要となり、知財・顧客資産が長期間活用できない「塩漬け」状態を招きかねません。",
            "また顧客リストには個人情報が含まれるため、個人情報保護法上、単に「共有」したのみでは各当事者が独自に利用することは認められません（利用目的の通知・同意、第三者提供制限等の問題）。",
            "そこでv7では、新規知的財産は原則として本会社に集約し、M&Aの際は株式譲渡で会社ごと売却する方向を優先する設計に整理しています。",
        ],
    },
    {
        "anchor_substr": "１．甲及び乙は、本契約期間中及び本契約終了後3年間、自ら、その関連会社、役員若しくは第三者をして、本事業と競合する事業",
        "title": "【第12条（競業禁止の範囲）について】",
        "body": [
            "v2第8条第1項は「完全に同一の事業」のみを競業禁止の対象としていました。",
            "もっとも、競業の範囲を「完全に同一」に限定すると、同種類似サービスを別法人で展開することで容易に潜脱され得るため、v7では「本事業と競合する事業」と広めに捕捉しています。",
            "将来の事業展開との関係で個別調整が必要な場合は、第27条の誠実協議又は別途の書面合意により対応する想定です。",
        ],
    },
    {
        "anchor_substr": "１．甲及び乙は、相手方の書面による事前の承諾がない限り、本会社の株式（以下「本株式」という。）を第三者に対して譲渡",
        "title": "【第13条（株式譲渡・先買権）について】",
        "body": [
            "v2第12条は譲渡を完全に禁止し、違反時の違約金（1,000万円）のみを定めていました。",
            "しかし、長期保有を一律強制すると、一方当事者が事業継続不能となった場合等に出口がなく、本会社・残存株主にも不利益が生じ得ます。",
            "v7では、設立後3年経過後については先買権を介した第三者譲渡を許容することで、退出経路を明文化しました。",
            "あわせて、第14条のコールオプション・第15条のプットオプションを設け、契約違反・倒産・反社該当等の場合の持分整理ルートを整備しています。",
        ],
    },
]

# ────────────────────────────────────────────────────────────
# Step 1: コピーしてpython-docxで本文置換
# ────────────────────────────────────────────────────────────
shutil.copy(SRC, DST)

doc = Document(str(DST))


def replace_paragraph_text(p, new_text: str) -> None:
    """段落本文を new_text で完全置換。runs を再構成し、元のフォントを最初のrunから継承。"""
    # 先頭runのrPrを保存
    first_r = None
    for r in p._p.findall(f"{W}r"):
        first_r = r
        break
    saved_rPr = None
    if first_r is not None:
        rPr = first_r.find(f"{W}rPr")
        if rPr is not None:
            saved_rPr = deepcopy(rPr)
    # 既存runを全削除（rangeマーカー類は段落内に残らないのでrのみ消す）
    for r in list(p._p.findall(f"{W}r")):
        p._p.remove(r)
    # 新run追加
    new_r = etree.SubElement(p._p, f"{W}r")
    if saved_rPr is not None:
        new_r.append(saved_rPr)
    t = etree.SubElement(new_r, f"{W}t")
    t.set(
        "{http://www.w3.org/XML/1998/namespace}space",
        "preserve",
    )
    t.text = new_text


replace_count = 0
for p in doc.paragraphs:
    txt = p.text
    for old_prefix, new_text in REPLACE_MAP.items():
        if txt.startswith(old_prefix):
            replace_paragraph_text(p, new_text)
            replace_count += 1
            print(f"[replace] '{old_prefix[:20]}...' -> done")
            break

assert replace_count == 4, f"Expected 4 replacements, got {replace_count}"
doc.save(str(DST))


# ────────────────────────────────────────────────────────────
# Step 2: zipfile直接編集
#   - word/document.xml: 既存コメントマーカー除去 → 新規コメントマーカー追加
#   - word/comments.xml: 新規コメント本体に置換
# ────────────────────────────────────────────────────────────


def make_comment_xml(comment_id: int, paragraphs: list[str]) -> etree._Element:
    c = etree.SubElement(
        etree.ElementBase(),
        f"{W}comment",
    )
    # 上の生成は使い物にならないので個別に組み立て直す
    raise NotImplementedError


def build_comments_root(comments: list[dict]) -> etree._Element:
    nsmap = {"w": W_NS}
    root = etree.Element(f"{W}comments", nsmap=nsmap)
    for cid, c in enumerate(comments):
        cmt = etree.SubElement(root, f"{W}comment")
        cmt.set(f"{W}id", str(cid))
        cmt.set(f"{W}author", "米谷尚起")
        cmt.set(f"{W}date", "2026-05-18T10:00:00Z")
        cmt.set(f"{W}initials", "NK")
        # タイトル段落
        if c.get("title"):
            p = etree.SubElement(cmt, f"{W}p")
            r = etree.SubElement(p, f"{W}r")
            rPr = etree.SubElement(r, f"{W}rPr")
            etree.SubElement(rPr, f"{W}lang").set(f"{W}eastAsia", "ja-JP")
            t = etree.SubElement(r, f"{W}t")
            t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            t.text = c["title"]
        for body in c["body"]:
            p = etree.SubElement(cmt, f"{W}p")
            r = etree.SubElement(p, f"{W}r")
            rPr = etree.SubElement(r, f"{W}rPr")
            etree.SubElement(rPr, f"{W}lang").set(f"{W}eastAsia", "ja-JP")
            t = etree.SubElement(r, f"{W}t")
            t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            t.text = body
    return root


def strip_existing_comment_markers(doc_root: etree._Element) -> int:
    """document.xml から既存の commentRangeStart/End/Reference を全削除。"""
    removed = 0
    targets = [
        f"{W}commentRangeStart",
        f"{W}commentRangeEnd",
        f"{W}commentReference",
    ]
    for tag in targets:
        for el in doc_root.iter(tag):
            parent = el.getparent()
            if parent is not None:
                parent.remove(el)
                removed += 1
    # commentReferenceを含む空の<w:r>も掃除（rPrのみ残ったもの）
    for r in list(doc_root.iter(f"{W}r")):
        # rの中にtもinstrTextもfldCharもtabもbrもpictもdrawingもなければ削除
        meaningful = False
        for child in r:
            local = etree.QName(child).localname
            if local in {"t", "instrText", "fldChar", "tab", "br", "pict", "drawing",
                          "object", "noBreakHyphen", "softHyphen", "sym"}:
                meaningful = True
                break
        if not meaningful:
            parent = r.getparent()
            if parent is not None:
                parent.remove(r)
    return removed


def attach_new_comments(doc_root: etree._Element, comments: list[dict]) -> None:
    """各コメントについて、anchorとなる段落を見つけて末尾にRangeStart→既存run→RangeEnd→Referenceを差し込む。

    実装簡略化のため、anchor段落の <w:r> 全体をRangeStartとRangeEndで囲み、最後にCommentReferenceを追加する。
    """
    body = doc_root.find(f"{W}body")
    paragraphs = list(body.iter(f"{W}p"))
    # 段落の全テキストを連結する関数
    def para_text(p):
        return "".join(t.text or "" for t in p.iter(f"{W}t"))

    for cid, c in enumerate(comments):
        anchor_substr = c["anchor_substr"]
        target = None
        for p in paragraphs:
            if anchor_substr in para_text(p):
                target = p
                break
        if target is None:
            raise RuntimeError(f"Anchor not found: {anchor_substr!r}")

        # commentRangeStart を先頭に挿入
        crs = etree.Element(f"{W}commentRangeStart")
        crs.set(f"{W}id", str(cid))
        # pPrの直後に置きたい
        pPr = target.find(f"{W}pPr")
        if pPr is not None:
            pPr.addnext(crs)
        else:
            target.insert(0, crs)
        # commentRangeEnd を末尾に
        cre = etree.SubElement(target, f"{W}commentRangeEnd")
        cre.set(f"{W}id", str(cid))
        # commentReference run を末尾に
        ref_r = etree.SubElement(target, f"{W}r")
        ref_rPr = etree.SubElement(ref_r, f"{W}rPr")
        # スタイル "a7" は v2 の参照スタイル名だが、v6 にあるかは不明なので付けない。
        cref = etree.SubElement(ref_r, f"{W}commentReference")
        cref.set(f"{W}id", str(cid))


# zip操作
TMP = DST.with_suffix(".tmp.docx")
shutil.move(DST, TMP)

with zipfile.ZipFile(TMP, "r") as zin:
    names = zin.namelist()
    contents = {n: zin.read(n) for n in names}

# document.xml 編集
parser = etree.XMLParser(remove_blank_text=False)
doc_root = etree.fromstring(contents["word/document.xml"], parser)
removed = strip_existing_comment_markers(doc_root)
print(f"[strip] removed {removed} legacy comment markers in document.xml")
attach_new_comments(doc_root, NEW_COMMENTS)
contents["word/document.xml"] = etree.tostring(
    doc_root, xml_declaration=True, encoding="UTF-8", standalone=True
)

# comments.xml 差し替え
comments_root = build_comments_root(NEW_COMMENTS)
contents["word/comments.xml"] = etree.tostring(
    comments_root, xml_declaration=True, encoding="UTF-8", standalone=True
)

# 書き出し
with zipfile.ZipFile(DST, "w", compression=zipfile.ZIP_DEFLATED) as zout:
    for n, data in contents.items():
        zout.writestr(n, data)

TMP.unlink(missing_ok=True)
print(f"[done] {DST}")
