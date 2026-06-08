"""v7_2→v7_3: 第11条・第16条の解説調/推奨表現を契約条項として書き直す。
- 第11条第2項末尾「なお〜十分ではない。」を削除
- 第11条第3項を契約義務として再構成
- 第11条第4項を明確な義務化
- 第16条第5項冒頭・段落125末尾の解説調を削除
- コメント本体は変更しない（コメント参照ランは温存）
"""
from __future__ import annotations
import shutil
import zipfile
from pathlib import Path
from copy import deepcopy
from lxml import etree

SRC = Path("/Users/kometaninaoki/Downloads/合弁会社設立契約書_株式会社Sophia_v7_2.docx")
DST = Path("/Users/kometaninaoki/Downloads/合弁会社設立契約書_株式会社Sophia_v7_3.docx")

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = f"{{{W_NS}}}"

# ────────────────────────────────────────────────────────────
# 置換内容（anchor: 先頭一致、text: 全文置換）
# ────────────────────────────────────────────────────────────
PARA_11_2_NEW = (
    "２．甲又は乙が本会社設立前から既に行っている事業に関するノウハウ、顧客リスト、"
    "商標、著作物その他の知的財産に係る権利（以下「既存知的財産権」という。）は、"
    "引き続きそれぞれの当事者に帰属するものとする。甲又は乙が、自らの既存知的財産権"
    "の本会社における利用を認める場合、その許諾の範囲、対価、期間その他の条件に"
    "ついては、甲又は乙と本会社との間で別途協議の上、決定する。"
)

PARA_11_3_NEW = (
    "３．甲及び乙は、本会社を第三者に譲渡する場合、新規知的財産権を本会社に帰属"
    "させたまま、本会社株式の譲渡の方法により移転するものとする。本会社が解散・"
    "清算した場合その他新規知的財産権を個別に処理する必要が生じた場合、当該権利の"
    "帰属及び処分の方法は、甲及び乙が協議の上、第20条に従い決定するものとし、"
    "原則として甲及び乙の共有としない。"
)

PARA_11_4_NEW = (
    "４．顧客リスト等に係る個人情報については、目的外利用、第三者提供並びに匿名"
    "加工情報及び仮名加工情報の取扱いに際し、甲及び乙は、個人情報の保護に関する"
    "法律その他関係法令に基づき必要な手続を履践するものとする。"
)

PARA_16_5_HEAD_NEW = (
    "５．甲又は乙は、前項各号に掲げる措置に加え、又はこれらに代えて、以下に定める"
    "オークション方式（甲及び乙がそれぞれ密封札により相手方保有株式の全部の取得を"
    "申し入れる買取提案単価を提示し、より高い単価を提示した当事者が相手方の保有"
    "株式の全部を取得する方式をいう。）により本株式の譲渡を申し出ることができる。"
)

PARA_16_5_3_NEW = (
    "（３）開札の結果、より高い買取提案単価を提示した当事者は、相手方の保有する"
    "本株式の全部を当該単価で取得し、相手方は当該譲渡に協力する義務を負う。単価が"
    "同額の場合は、再入札その他甲及び乙が協議で定める方法による。"
)

# anchor先頭文字列 → 新本文
REPLACE_MAP = {
    "２．甲又は乙が本会社設立前から既に行っている事業に関するノウハウ": PARA_11_2_NEW,
    "３．将来、本会社を第三者に譲渡する場合（いわゆるM&A）に備え": PARA_11_3_NEW,
    "４．前項のほか、顧客リスト等に係る個人情報については": PARA_11_4_NEW,
    "５．前項各号に掲げる措置に加え、又はこれらに代えて、デッドロック状態が継続": PARA_16_5_HEAD_NEW,
    "（３）開札の結果、より高い買取提案単価を提示した当事者は": PARA_16_5_3_NEW,
}


def replace_paragraph_text(p_elem: etree._Element, new_text: str) -> None:
    """段落本文を全置換。commentRangeStart/End/commentReferenceは温存する。"""
    # 既存runのうち、commentReferenceを含むものは保持
    runs_to_remove: list[etree._Element] = []
    saved_rPr = None
    first_text_run_found = False
    for r in p_elem.findall(f"{W}r"):
        has_cref = r.find(f"{W}commentReference") is not None
        if has_cref:
            continue  # 保持
        # フォーマットを最初のtext runから取得
        if not first_text_run_found:
            rPr = r.find(f"{W}rPr")
            if rPr is not None:
                saved_rPr = deepcopy(rPr)
            first_text_run_found = True
        runs_to_remove.append(r)
    for r in runs_to_remove:
        p_elem.remove(r)

    # 新しいrun作成
    new_r = etree.Element(f"{W}r")
    if saved_rPr is not None:
        new_r.append(saved_rPr)
    t = etree.SubElement(new_r, f"{W}t")
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t.text = new_text

    # 挿入位置: commentRangeEndの直前、commentRangeEndがなければcommentReferenceランの直前、なければ末尾
    cre = p_elem.find(f"{W}commentRangeEnd")
    if cre is not None:
        cre.addprevious(new_r)
        return
    # commentReferenceを含むrunを探す
    for r in p_elem.findall(f"{W}r"):
        if r.find(f"{W}commentReference") is not None:
            r.addprevious(new_r)
            return
    p_elem.append(new_r)


def para_text(p_elem: etree._Element) -> str:
    return "".join(t.text or "" for t in p_elem.iter(f"{W}t"))


# ────────────────────────────────────────────────────────────
# 実処理
# ────────────────────────────────────────────────────────────
shutil.copy(SRC, DST)
TMP = DST.with_suffix(".tmp.docx")
shutil.move(DST, TMP)

with zipfile.ZipFile(TMP, "r") as zin:
    contents = {n: zin.read(n) for n in zin.namelist()}

parser = etree.XMLParser(remove_blank_text=False)
doc_root = etree.fromstring(contents["word/document.xml"], parser)

body = doc_root.find(f"{W}body")
paragraphs = body.findall(f"{W}p")

remaining = dict(REPLACE_MAP)
for p in paragraphs:
    txt = para_text(p)
    for prefix in list(remaining.keys()):
        if txt.startswith(prefix):
            replace_paragraph_text(p, remaining[prefix])
            print(f"[replace] '{prefix[:24]}...' -> done")
            del remaining[prefix]
            break

assert not remaining, f"Unmatched anchors: {list(remaining.keys())}"

contents["word/document.xml"] = etree.tostring(
    doc_root, xml_declaration=True, encoding="UTF-8", standalone=True
)

with zipfile.ZipFile(DST, "w", compression=zipfile.ZIP_DEFLATED) as zout:
    for n, data in contents.items():
        zout.writestr(n, data)

TMP.unlink(missing_ok=True)
print(f"[done] {DST}")
