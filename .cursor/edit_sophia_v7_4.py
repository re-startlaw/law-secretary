"""Apply修正指示 to v7_4 docx.

修正項目:
1. 第6条第1項柱書: 「議決権の過半数」→「議決権の4分の3超」
2-1. 第13条第3項第2号: 「当該日」→「同日（当該直前の四半期末日）」
2-2. 第14条第2項: 同上
2-3. 第15条第2項: 同上
2-4. 第29条（存続条項）の新設（第28条本文と「以上、…」の間に挿入）
3-1. 第1条見出し: 「対象会社」→「本会社」
3-2. 第8条見出し: 「対象会社」→「本会社」
3-3. 第9条第1項: 「追加出資」初出位置で定義
3-4. 第9条第3項: 「追加出資」定義部分を削除し短縮
"""
from __future__ import annotations

import os
import shutil
import zipfile
from xml.etree import ElementTree as ET

DOCX = (
    "/Users/kometaninaoki/Library/CloudStorage/"
    "GoogleDrive-n.kometani@re-startlaw.com/マイドライブ/共有用/"
    "01_事件記録/え_えいすた/合弁契約(sophia)/"
    "合弁会社設立契約書_株式会社Sophia_v7_4.docx"
)

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = f"{{{W_NS}}}"
ET.register_namespace("w", W_NS)
ET.register_namespace("w14", "http://schemas.microsoft.com/office/word/2010/wordml")


def paragraph_text(p: ET.Element) -> str:
    return "".join(t.text or "" for t in p.iter(f"{W}t"))


def replace_in_runs(p: ET.Element, old: str, new: str) -> bool:
    """Replace `old` with `new` inside the concatenated text of paragraph p.

    Strategy: if `old` lies entirely within a single `w:t`, do an in-place
    replacement of that run's text. Otherwise we fall back to clearing the
    runs and putting the whole replaced text into the first run.
    """
    runs_t = list(p.iter(f"{W}t"))
    full = "".join(t.text or "" for t in runs_t)
    if old not in full:
        return False
    # Find the run that contains the entire `old` substring.
    for t in runs_t:
        if t.text and old in t.text:
            t.text = t.text.replace(old, new)
            return True
    # Spans multiple runs — rebuild text into the first run.
    new_full = full.replace(old, new)
    runs_t[0].text = new_full
    for t in runs_t[1:]:
        t.text = ""
    return True


def make_heading_paragraph(text: str) -> ET.Element:
    """Build a chapter heading paragraph matching existing style (bold, sz 22)."""
    p = ET.Element(f"{W}p")
    pPr = ET.SubElement(p, f"{W}pPr")
    spacing = ET.SubElement(pPr, f"{W}spacing")
    spacing.set(f"{W}before", "240")
    spacing.set(f"{W}after", "40")
    rPr_p = ET.SubElement(pPr, f"{W}rPr")
    lang_p = ET.SubElement(rPr_p, f"{W}lang")
    lang_p.set(f"{W}eastAsia", "ja-JP")

    r = ET.SubElement(p, f"{W}r")
    rPr = ET.SubElement(r, f"{W}rPr")
    ET.SubElement(rPr, f"{W}b")
    sz = ET.SubElement(rPr, f"{W}sz")
    sz.set(f"{W}val", "22")
    lang = ET.SubElement(rPr, f"{W}lang")
    lang.set(f"{W}eastAsia", "ja-JP")
    t = ET.SubElement(r, f"{W}t")
    t.text = text
    return p


def make_body_paragraph(text: str) -> ET.Element:
    """Build a normal body paragraph matching existing style (no indent)."""
    p = ET.Element(f"{W}p")
    pPr = ET.SubElement(p, f"{W}pPr")
    spacing = ET.SubElement(pPr, f"{W}spacing")
    spacing.set(f"{W}after", "40")
    rPr_p = ET.SubElement(pPr, f"{W}rPr")
    lang_p = ET.SubElement(rPr_p, f"{W}lang")
    lang_p.set(f"{W}eastAsia", "ja-JP")

    r = ET.SubElement(p, f"{W}r")
    rPr = ET.SubElement(r, f"{W}rPr")
    lang = ET.SubElement(rPr, f"{W}lang")
    lang.set(f"{W}eastAsia", "ja-JP")
    t = ET.SubElement(r, f"{W}t")
    t.text = text
    return p


def main() -> None:
    # Read document.xml
    with zipfile.ZipFile(DOCX) as zin:
        names = zin.namelist()
        contents = {n: zin.read(n) for n in names}

    tree = ET.ElementTree(ET.fromstring(contents["word/document.xml"]))
    root = tree.getroot()
    body = root.find(f"{W}body")
    assert body is not None

    paragraphs = list(body.iter(f"{W}p"))

    def find_by_substr(substr: str, *, headings_only: bool = False) -> ET.Element:
        hits = []
        for p in paragraphs:
            txt = paragraph_text(p)
            if substr in txt:
                hits.append(p)
        if not hits:
            raise SystemExit(f"Paragraph not found: {substr!r}")
        if headings_only and len(hits) > 1:
            # Pick the one whose text starts with 第 (heading).
            hits = [h for h in hits if paragraph_text(h).startswith("第")]
        if len(hits) != 1:
            raise SystemExit(
                f"Ambiguous paragraph: {substr!r} -> {[paragraph_text(h) for h in hits]}"
            )
        return hits[0]

    # 1. 第6条第1項: 議決権の過半数 -> 議決権の4分の3超
    p_art6 = find_by_substr(
        "本会社における以下の各号に定める事項の決定（以下「重要事項」と総称する。）"
    )
    assert replace_in_runs(p_art6, "議決権の過半数", "議決権の4分の3超"), "第6条修正失敗"

    # 2-1. 第13条第3項第2号
    p_art13_2 = find_by_substr("（２）譲渡希望通知日の直前の四半期末")
    assert replace_in_runs(
        p_art13_2,
        "当該日における本会社の発行済株式総数",
        "同日（当該直前の四半期末日）における本会社の発行済株式総数",
    ), "第13条修正失敗"

    # 2-2. 第14条第2項
    p_art14_2 = find_by_substr("買取請求がなされた日の直前の四半期末")
    # 「当該日」が単一runにある想定
    runs_t = list(p_art14_2.iter(f"{W}t"))
    full = "".join(t.text or "" for t in runs_t)
    assert "当該日における本会社の発行済株式総数" in full, "第14条 anchor not found"
    # Replace by run-rebuild because text spans multiple runs.
    new_full = full.replace(
        "当該日における本会社の発行済株式総数",
        "同日（当該直前の四半期末日）における本会社の発行済株式総数",
    )
    runs_t[0].text = new_full
    for t in runs_t[1:]:
        t.text = ""

    # 2-3. 第15条第2項
    p_art15_2 = find_by_substr("売渡請求がなされた日の直前の四半期末")
    runs_t = list(p_art15_2.iter(f"{W}t"))
    full = "".join(t.text or "" for t in runs_t)
    assert "当該日における本会社の発行済株式総数" in full, "第15条 anchor not found"
    new_full = full.replace(
        "当該日における本会社の発行済株式総数",
        "同日（当該直前の四半期末日）における本会社の発行済株式総数",
    )
    runs_t[0].text = new_full
    for t in runs_t[1:]:
        t.text = ""

    # 3-1. 第1条見出し
    p_art1_h = find_by_substr("第1条（対象会社の設立及び出資比率）", headings_only=True)
    assert replace_in_runs(p_art1_h, "対象会社", "本会社"), "第1条見出し修正失敗"

    # 3-2. 第8条見出し
    p_art8_h = find_by_substr("第8条（対象会社による情報提供）", headings_only=True)
    assert replace_in_runs(p_art8_h, "対象会社", "本会社"), "第8条見出し修正失敗"

    # 3-3. 第9条第1項
    p_art9_1 = find_by_substr(
        "本会社の運転資金その他事業運営に必要な資金が不足する場合"
    )
    assert replace_in_runs(
        p_art9_1,
        "出資比率に応じて追加出資を行うものとする。",
        "出資比率に応じて追加出資（本会社が新たに発行する株式の引受けをいう。以下同じ。）を行うものとする。",
    ), "第9条第1項修正失敗"

    # 3-4. 第9条第3項
    p_art9_3 = find_by_substr("甲又は乙が本条に基づく追加出資又は資金提供に関する協議開始後")
    runs_t = list(p_art9_3.iter(f"{W}t"))
    full = "".join(t.text or "" for t in runs_t)
    target_old = "日以内に、本会社が新たに発行する株式の引受け（以下「追加出資」という。）に応じない場合"
    target_new = "日以内に、追加出資に応じない場合"
    assert target_old in full, "第9条第3項 anchor not found"
    new_full = full.replace(target_old, target_new)
    runs_t[0].text = new_full
    for t in runs_t[1:]:
        t.text = ""

    # 2-4. 第29条 新設
    p_art28_h = find_by_substr("第28条（管轄裁判所）", headings_only=True)
    p_art28_b = find_by_substr(
        "本契約に関する一切の紛争については、大阪地方裁判所を第一審の専属的合意管轄裁判所とする。"
    )
    # 末尾の「以上、本契約の成立を証するため、本書2通を作成」段落の直前に挿入
    p_closing = find_by_substr("以上、本契約の成立を証するため、本書2通を作成")

    # body の直接の子としての位置を取得
    body_children = list(body)
    insert_idx = body_children.index(p_closing)
    new_h = make_heading_paragraph("第29条（存続条項）")
    new_b = make_body_paragraph(
        "本契約が解除、期間満了その他の事由により終了した場合であっても、"
        "第11条（知的財産権の処理）、第12条（競業禁止及び引抜きの禁止）、"
        "第14条（コールオプション）、第15条（プットオプション）、"
        "第18条（損害賠償）、第21条（秘密保持）、第28条（管轄裁判所）及び本条の規定は、"
        "本契約終了後も引き続き効力を有する。"
        "ただし、第12条（競業禁止及び引抜きの禁止）については本契約終了後3年間、"
        "第21条（秘密保持）については本契約終了後5年間に限り効力を有するものとする。"
    )
    body.insert(insert_idx, new_h)
    body.insert(insert_idx + 1, new_b)

    # Serialize document.xml
    xml_decl = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    new_xml = xml_decl + ET.tostring(root, encoding="utf-8")
    contents["word/document.xml"] = new_xml

    # Write back to docx (preserve structure & ordering)
    tmp = DOCX + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for name in names:
            zout.writestr(name, contents[name])
    os.replace(tmp, DOCX)
    print("done")


if __name__ == "__main__":
    main()
