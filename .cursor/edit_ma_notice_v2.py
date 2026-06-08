"""馬さん案件・受任通知_2 への修正履歴付き編集スクリプト。

入出力:
- 入力: 260512電子内容証明案_2.docx（既に _2 へコピー済み）
- 修正は Word 変更履歴形式（<w:ins>/<w:del>）で本ファイルに直接適用する

修正項目（プロンプト指示の 9 項目 + 関連見出し）:
1+2. Para 24 見出し・Para 25 本文：時間表記・接触行為の表現
3.   Para 34：数学授業／離席の趣旨
4.   Para 40：精神的損害の表現
5.   Para 43：(5)妻関連損害の削除
6.   Para 63：G10 学費の請求事実関係加筆
7.   Para 59：学費相当額を 1,388,000 円に具体化
8.   Para 50：直ちに復学の位置付け
9.   Para 45 の後：新規損害項目を 3 項目追加
"""

from __future__ import annotations

import copy
from pathlib import Path

from docx import Document
from lxml import etree

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = f"{{{W_NS}}}"
NSMAP = {"w": W_NS}

AUTHOR = "米谷尚起"
DATE = "2026-05-12T00:00:00Z"

INPUT = Path(
    "/Users/kometaninaoki/Library/CloudStorage/"
    "GoogleDrive-n.kometani@re-startlaw.com/マイドライブ/"
    "共有用/01_事件記録/ま_馬さん/Wordファイル/260512電子内容証明案_2.docx"
)


class IdGen:
    def __init__(self) -> None:
        self.n = 1000

    def next(self) -> str:
        self.n += 1
        return str(self.n)


IDS = IdGen()


def make_run(text: str, base_rpr: etree._Element | None = None) -> etree._Element:
    r = etree.SubElement(etree.Element(f"{W}tmp"), f"{W}r")
    if base_rpr is not None:
        r.append(copy.deepcopy(base_rpr))
    t = etree.SubElement(r, f"{W}t")
    t.text = text
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    return r


def make_del_run(text: str, base_rpr: etree._Element | None = None) -> etree._Element:
    """delText を持つ run を作って返す。"""
    r = etree.SubElement(etree.Element(f"{W}tmp"), f"{W}r")
    if base_rpr is not None:
        r.append(copy.deepcopy(base_rpr))
    t = etree.SubElement(r, f"{W}delText")
    t.text = text
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    return r


def wrap_ins(runs: list[etree._Element]) -> etree._Element:
    ins = etree.Element(f"{W}ins")
    ins.set(f"{W}id", IDS.next())
    ins.set(f"{W}author", AUTHOR)
    ins.set(f"{W}date", DATE)
    for r in runs:
        ins.append(r)
    return ins


def wrap_del(runs: list[etree._Element]) -> etree._Element:
    d = etree.Element(f"{W}del")
    d.set(f"{W}id", IDS.next())
    d.set(f"{W}author", AUTHOR)
    d.set(f"{W}date", DATE)
    for r in runs:
        d.append(r)
    return d


def extract_first_rpr(p: etree._Element) -> etree._Element | None:
    """段落の最初の w:r/w:rPr を見つけて返す（書式継承用）。"""
    rpr = p.find(f".//{W}r/{W}rPr", namespaces=NSMAP)
    return rpr


def replace_paragraph_content(p: etree._Element, new_text: str) -> None:
    """段落のすべての run/ins/del 等を撤去し、
    旧テキストを <w:del>、新テキストを <w:ins> で並べる。"""
    base_rpr = extract_first_rpr(p)
    # 旧テキストを収集（run内のテキストを連結）
    old_text_parts: list[str] = []
    pPr = p.find(f"{W}pPr", namespaces=NSMAP)
    # 子要素を全部走査
    children_to_remove: list[etree._Element] = []
    for child in list(p):
        if child.tag == f"{W}pPr":
            continue
        if child.tag == f"{W}r":
            for t in child.findall(f"{W}t", namespaces=NSMAP):
                old_text_parts.append(t.text or "")
        elif child.tag == f"{W}ins":
            for r in child.findall(f"{W}r", namespaces=NSMAP):
                for t in r.findall(f"{W}t", namespaces=NSMAP):
                    old_text_parts.append(t.text or "")
        elif child.tag == f"{W}del":
            # 既存の削除は無視
            pass
        children_to_remove.append(child)
    for c in children_to_remove:
        p.remove(c)

    old_text = "".join(old_text_parts)
    # del と ins を追加
    if old_text:
        del_elem = wrap_del([make_del_run(old_text, base_rpr)])
        p.append(del_elem)
    if new_text:
        ins_elem = wrap_ins([make_run(new_text, base_rpr)])
        p.append(ins_elem)


def delete_paragraph(p: etree._Element) -> None:
    """段落全体を削除扱いにする：
    全 run を <w:del> でラップ + pPr に paragraph-end の del マーカーを置く。"""
    base_rpr = extract_first_rpr(p)
    # 旧テキスト収集
    old_text_parts: list[str] = []
    children_to_remove: list[etree._Element] = []
    pPr = p.find(f"{W}pPr", namespaces=NSMAP)
    for child in list(p):
        if child.tag == f"{W}pPr":
            continue
        if child.tag == f"{W}r":
            for t in child.findall(f"{W}t", namespaces=NSMAP):
                old_text_parts.append(t.text or "")
        children_to_remove.append(child)
    for c in children_to_remove:
        p.remove(c)
    old_text = "".join(old_text_parts)
    if old_text:
        del_elem = wrap_del([make_del_run(old_text, base_rpr)])
        p.append(del_elem)
    # paragraph-end 削除マーカー
    if pPr is None:
        pPr = etree.SubElement(p, f"{W}pPr")
        p.insert(0, pPr)
    rPr = pPr.find(f"{W}rPr", namespaces=NSMAP)
    if rPr is None:
        rPr = etree.SubElement(pPr, f"{W}rPr")
    del_marker = etree.Element(f"{W}del")
    del_marker.set(f"{W}id", IDS.next())
    del_marker.set(f"{W}author", AUTHOR)
    del_marker.set(f"{W}date", DATE)
    rPr.insert(0, del_marker)


def insert_paragraph_after(
    after_p: etree._Element, text: str, base_p: etree._Element
) -> etree._Element:
    """after_p の直後に新規段落を追加（挿入扱い）。
    base_p の pPr / rPr を複製して書式を揃える。"""
    new_p = etree.Element(f"{W}p")
    base_pPr = base_p.find(f"{W}pPr", namespaces=NSMAP)
    base_rpr = extract_first_rpr(base_p)
    if base_pPr is not None:
        new_pPr = copy.deepcopy(base_pPr)
        new_p.append(new_pPr)
        # paragraph-end が挿入であることを示す
        rPr = new_pPr.find(f"{W}rPr", namespaces=NSMAP)
        if rPr is None:
            rPr = etree.SubElement(new_pPr, f"{W}rPr")
        ins_marker = etree.Element(f"{W}ins")
        ins_marker.set(f"{W}id", IDS.next())
        ins_marker.set(f"{W}author", AUTHOR)
        ins_marker.set(f"{W}date", DATE)
        rPr.insert(0, ins_marker)
    # 本文 run を ins でラップして追加
    if text:
        new_p.append(wrap_ins([make_run(text, base_rpr)]))
    # 挿入位置
    after_p.addnext(new_p)
    return new_p


def main() -> None:
    doc = Document(str(INPUT))
    paragraphs = doc.paragraphs

    # Paragraph 24: 見出し「(1) 小牧氏が...襟元を掴んで引き移動させた行為」
    p24_old_check = "（１）　小牧氏がAngelina氏の襟元を掴んで引き移動させた行為"
    p24 = paragraphs[24]
    assert p24.text.strip() == p24_old_check, f"P24 mismatch: {p24.text!r}"
    replace_paragraph_content(
        p24._element,
        "（１）　小牧氏がAngelina氏の襟元付近に身体的接触を加えて移動させ、トラック脇で停止させた行為",
    )

    # Paragraph 25: 本文（時間 + 暴力性の表現）
    p25 = paragraphs[25]
    assert p25.text.startswith("　令和８年４月１３日午前９時１８分頃")
    replace_paragraph_content(
        p25._element,
        (
            "　令和８年４月１３日午前９時２０分頃から同９時３０分頃までの間、貴校体育館において、"
            "小牧氏は、Angelina氏に対し、後方からその襟元付近に身体的接触を加えてこれを掴み、"
            "Angelina氏の身体を伴って移動させた上、トラック脇においてこれを停止させるとともに、"
            "大声で叱責しました。当該事実は、Angelina氏の同日付書面陳述、当方が保有する事案"
            "当日の映像等から、優に認められるものです。"
        ),
    )

    # Paragraph 34: 数学授業・離席状況
    p34 = paragraphs[34]
    assert p34.text.startswith("　Angelina氏は、令和８年４月１３日の１時間目の算数の時間中")
    replace_paragraph_content(
        p34._element,
        (
            "　Angelina氏は、令和８年４月１３日の１時間目の数学の時間（午前８時４０分から"
            "同９時３５分まで）中、約４分間体育館に立ち寄ったことを認めています。これは、"
            "Angelina氏が自身の課題を比較的早く終えた後、長男であるGeorge氏の約４分間の"
            "クロスカントリー競技を観覧するため、体育館に向かったものです。当該時間は、"
            "前半が教員による講義の時間、後半が各生徒が課題を行う自習の時間として運用"
            "されていたものであり、また、当該自習時間中に生徒が一時的に教室を離れること"
            "についても、現場の教員において従前から黙認されてきたものです。"
        ),
    )

    # Paragraph 40: (2) うつ病悪化 → 精神状態悪化等
    p40 = paragraphs[40]
    assert p40.text.startswith("（２）　貴校から不適切な対応を受けたこと")
    replace_paragraph_content(
        p40._element,
        (
            "（２）　貴校から不適切な対応を受けたこと、登校できないこと、及び事実に反して"
            "Angelina氏が虚偽の主張をしている旨を周囲に言いふらされたことによりAngelina氏"
            "の精神状態が悪化し、投薬調整及び継続的な心理カウンセリングを要する状態となった"
            "ことに伴う精神科通院費用及び心理カウンセリング費用"
        ),
    )

    # Paragraph 43: (5) 妻の毎日学校立寄り損害 → 削除
    p43 = paragraphs[43]
    assert p43.text.startswith("（５）　George氏の看護者の校内立入りが何らの根拠なく拒絶")
    delete_paragraph(p43._element)

    # Paragraph 45: (7) 弁護士費用… の後に新規 (8)(9)(10) を挿入
    p45 = paragraphs[45]
    assert p45.text.startswith("（７）　弁護士費用その他本件への法的対応")
    base_p_element = p45._element
    new_items = [
        "（８）　Angelina氏の長期にわたる学籍中断により、Angelina氏の学業継続性が毀損されていることに関する損害",
        "（９）　Angelina氏の試験成績、GPA及びTranscriptへの不利益な影響に関する損害",
        "（１０）　Angelina氏の大学進学経路への現実的影響並びに教育機会の喪失及び将来的進学リスクに関する損害",
    ]
    anchor = base_p_element
    for txt in new_items:
        anchor = insert_paragraph_after(anchor, txt, base_p_element)

    # Paragraph 50: (2) 直ちに復学 → 公平・安全な教育環境確保
    p50 = paragraphs[50]
    assert p50.text.strip() == "（２）　Angelina氏を直ちに復学させていただくこと"
    replace_paragraph_content(
        p50._element,
        (
            "（２）　Angelina氏に対する公平かつ安全な教育環境を確保した上で、"
            "Angelina氏の復学を実現していただくこと（なお、当方として復学権自体を放棄"
            "するものではありません。）"
        ),
    )

    # Paragraph 59: (3) 学費相当額（金○○万円） → 1,388,000円具体化
    p59 = paragraphs[59]
    assert "学費相当額（金○○万円）" in p59.text, p59.text
    replace_paragraph_content(
        p59._element,
        (
            "（３）　もっとも、Angelina氏及び通知人は、貴校との在学契約の継続を強く"
            "希望しております。当方は、学費相当額（金１,３８８,０００円。Angelina氏の"
            "Ｇ１０学年第一期学費に相当します。）を当職の弁護士預り金口座において保全し、"
            "貴校による教育役務の提供再開と引換えに、直ちに同額をお支払いする用意がございます。"
        ),
    )

    # Paragraph 63: G10学費の支払方法… → 請求事実関係を加筆
    p63 = paragraphs[63]
    assert p63.text.startswith("　なお、Angelina氏のG10学年")
    replace_paragraph_content(
        p63._element,
        (
            "　なお、Angelina氏のＧ１０学年（令和８年度後期以降）の学費につきましては、"
            "貴校より令和８年３月９日に保護者宛て請求通知が発出されており、本書面送付日"
            "（令和８年５月１２日）にも貴校より学費全体に関する支払期限通知メールが送信"
            "されているところ、これらに対する具体的な支払方法及び時期等の取扱いに関する"
            "事項につきましては、通知人より、貴校に対し別途ご説明をいたします。"
        ),
    )

    doc.save(str(INPUT))
    print(f"saved: {INPUT}")


if __name__ == "__main__":
    main()
