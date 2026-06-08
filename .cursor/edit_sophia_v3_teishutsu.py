"""Apply 3 修正指示 to 合弁会社設立契約書_株式会社Sophia_v3_提出版_2.docx.

1. 第11条第3項・第20条第1項第5号 の循環参照を解消
2. 取締役会の廃止（第2条見出し/第1項、第4条第2項、第5条第2項、第17条第2項）
3. 「出資比率」→「持株比率」（第10条第3項の再定義も整理）
"""
from __future__ import annotations

import os
import zipfile
from xml.etree import ElementTree as ET

DOCX = (
    "/Users/kometaninaoki/Library/CloudStorage/"
    "GoogleDrive-n.kometani@re-startlaw.com/マイドライブ/共有用/"
    "01_事件記録/え_えいすた/合弁契約(sophia)/"
    "合弁会社設立契約書_株式会社Sophia_v3_提出版_2.docx"
)

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = f"{{{W_NS}}}"
ET.register_namespace("w", W_NS)
ET.register_namespace("w14", "http://schemas.microsoft.com/office/word/2010/wordml")


def paragraph_text(p: ET.Element) -> str:
    return "".join(t.text or "" for t in p.iter(f"{W}t"))


def replace_text(p: ET.Element, old: str, new: str, *, allow_multi_run: bool = True) -> bool:
    """Replace `old` -> `new` in paragraph text. Handles multi-run text."""
    runs_t = list(p.iter(f"{W}t"))
    full = "".join(t.text or "" for t in runs_t)
    if old not in full:
        return False
    # Try single-run replacement first.
    for t in runs_t:
        if t.text and old in t.text:
            t.text = t.text.replace(old, new)
            return True
    if not allow_multi_run:
        return False
    new_full = full.replace(old, new)
    runs_t[0].text = new_full
    for t in runs_t[1:]:
        t.text = ""
    return True


def replace_in_all(paragraphs: list[ET.Element], old: str, new: str) -> int:
    """Replace every occurrence of `old` -> `new` across paragraphs, including
    repeated occurrences in the same paragraph."""
    count = 0
    for p in paragraphs:
        # Loop until no more occurrences
        while True:
            runs_t = list(p.iter(f"{W}t"))
            full = "".join(t.text or "" for t in runs_t)
            if old not in full:
                break
            new_full = full.replace(old, new)
            runs_t[0].text = new_full
            for t in runs_t[1:]:
                t.text = ""
            count += full.count(old)
            break  # We replaced all in this paragraph in one go.
    return count


def main() -> None:
    with zipfile.ZipFile(DOCX) as zin:
        names = zin.namelist()
        contents = {n: zin.read(n) for n in names}

    root = ET.fromstring(contents["word/document.xml"])
    body = root.find(f"{W}body")
    assert body is not None
    paragraphs = list(body.iter(f"{W}p"))

    def find_by_substr(substr: str, *, unique: bool = True) -> ET.Element:
        hits = [p for p in paragraphs if substr in paragraph_text(p)]
        if not hits:
            raise SystemExit(f"Not found: {substr!r}")
        if unique and len(hits) != 1:
            raise SystemExit(
                f"Ambiguous: {substr!r} -> {[paragraph_text(h)[:40] for h in hits]}"
            )
        return hits[0]

    # ---- 1. 循環参照の解消 ----
    # 第11条第3項: "第20条に従い" を削除
    p_11_3 = find_by_substr("本会社株式の譲渡の方法により移転するもの")
    assert replace_text(
        p_11_3,
        "甲及び乙が協議の上、第20条に従い決定するものとし",
        "甲及び乙が協議の上決定するものとし",
    ), "第11条第3項修正失敗"

    # 第20条第1項第5号
    p_20_5 = find_by_substr("（５）本会社に帰属する新規知的財産権その他の無体財産の処理")
    assert replace_text(
        p_20_5,
        "（５）本会社に帰属する新規知的財産権その他の無体財産の処理は、第11条第3項及び第4項に定めるところに従う。",
        "（５）本会社に帰属する新規知的財産権その他の無体財産の処理は、解散時に甲及び乙が協議の上で決定するものとし、原則として甲及び乙の共有としない。",
    ), "第20条第1項第5号修正失敗"

    # ---- 2. 取締役会の廃止 ----
    # 第2条見出し
    p_art2_h = find_by_substr("第2条（取締役会の設置・取締役の指名権及び解任権）")
    assert replace_text(
        p_art2_h,
        "第2条（取締役会の設置・取締役の指名権及び解任権）",
        "第2条（取締役の指名権及び解任権）",
    ), "第2条見出し修正失敗"

    # 第2条第1項
    p_art2_1 = find_by_substr("１．本会社は、取締役会を設置する。")
    assert replace_text(
        p_art2_1,
        "１．本会社は、取締役会を設置する。",
        "１．本会社は、取締役会を設置しない。",
    ), "第2条第1項修正失敗"

    # 第4条第2項
    p_art4_2 = find_by_substr("本会社設立後2年を経過した後の代表取締役の選定")
    assert replace_text(
        p_art4_2,
        "甲及び乙が協議の上、取締役会の決議により定めるものとする。",
        "甲及び乙が協議の上、株主総会の決議により定めるものとする。",
    ), "第4条第2項修正失敗"

    # 第5条第2項
    p_art5_2 = find_by_substr("第4条第1項に定める者以外の者を代表取締役に選定する場合")
    assert replace_text(
        p_art5_2,
        "甲及び乙が協議の上、取締役会の決議により役員報酬を定めることができる。",
        "甲及び乙が協議の上、株主総会の決議により役員報酬を定めることができる。",
    ), "第5条第2項修正失敗"

    # 第17条第2項
    p_art17_2 = find_by_substr("前項の取引については、事前に取締役会に報告した上")
    assert replace_text(
        p_art17_2,
        "前項の取引については、事前に取締役会に報告した上、当該取引について特別の利害関係を有しない取締役の承認を要するものとし、取引後速やかに取締役会にも報告する。利害関係を有しない取締役が不在又は1名しかいない場合は、株主総会の承認を要するものとする。",
        "前項の取引については、事前に株主総会の承認を要するものとし、取引後速やかに甲及び乙に対し報告するものとする。",
    ), "第17条第2項修正失敗"

    # ---- 3. 「出資比率」→「持株比率」 ----
    # 第10条第3項: 再定義の括弧書きを除去（"以下「持株比率」という。"）
    p_art10_3 = find_by_substr(
        "配当を行う場合、その配当額は、配当基準日における各株主の保有株式数"
    )
    assert replace_text(
        p_art10_3,
        "配当基準日における各株主の保有株式数が発行済株式総数に占める比率（以下「持株比率」という。）に応じて按分する。",
        "配当基準日における各株主の持株比率に応じて按分する。",
    ), "第10条第3項修正失敗"

    # 全文一括置換: 出資比率 -> 持株比率
    n = replace_in_all(paragraphs, "出資比率", "持株比率")
    print(f"Replaced 出資比率 -> 持株比率: {n} occurrences (paragraph-level count)")

    # シリアライズ
    xml_decl = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    new_xml = xml_decl + ET.tostring(root, encoding="utf-8")
    contents["word/document.xml"] = new_xml

    tmp = DOCX + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for name in names:
            zout.writestr(name, contents[name])
    os.replace(tmp, DOCX)
    print("done")


if __name__ == "__main__":
    main()
