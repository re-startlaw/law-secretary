"""差出人ブロックの w:firstLineChars 属性を XML 直接編集で削除し、
left_indent のみで右側配置するクリーンな版を作成する。

問題:
  python-docx の paragraph_format.first_line_indent = None は w:firstLine（EMU値）は
  消すが、w:firstLineChars（文字数指定）は消さない。テンプレ由来の w:firstLineChars
  が残ったため、各行先頭が大幅に右へ押し出されて表示が崩れていた。

解決:
  lxml で w:ind の firstLineChars / firstLine 属性を直接削除する。
  left_indent はテンプレオリジナル値 Emu(4140835) ≒ 11.5cm に戻す（米谷弁護士の
  image #2 ベース位置）。
"""

from __future__ import annotations

import shutil
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.shared import Emu, Pt

DRIVE = Path(
    "/Users/kometaninaoki/Library/CloudStorage/GoogleDrive-n.kometani@re-startlaw.com/マイドライブ"
)
SRC = DRIVE / "共有用/01_事件記録/ま_馬さん/Wordファイル/260512電子内容証明案.docx"
DEST = DRIVE / "共有用/01_事件記録/ま_馬さん/Wordファイル/260512電子内容証明案_2.docx"

SENDER_INDENT = Emu(4140835)  # テンプレ準拠


def clean_ind_attrs(p) -> None:
    """段落の w:ind から firstLineChars/firstLine/leftChars を削除（left のみ残す）。"""
    pPr = p._element.find(qn("w:pPr"))
    if pPr is None:
        return
    ind = pPr.find(qn("w:ind"))
    if ind is None:
        return
    for attr in (
        qn("w:firstLine"),
        qn("w:firstLineChars"),
        qn("w:leftChars"),
        qn("w:hangingChars"),
        qn("w:hanging"),
    ):
        if attr in ind.attrib:
            del ind.attrib[attr]


def main() -> None:
    shutil.copy2(SRC, DEST)
    doc = Document(DEST)
    paras = list(doc.paragraphs)

    # 全段落で firstLineChars 等の不可視属性を除去
    for p in paras:
        clean_ind_attrs(p)

    # 差出人ブロック（para 6-13）に左字下げ + 行間設定を上書き
    for i in range(6, 14):
        p = paras[i]
        pf = p.paragraph_format
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        pf.left_indent = SENDER_INDENT
        pf.right_indent = None
        pf.first_line_indent = None
        pf.space_before = Pt(0)
        pf.space_after = Pt(0)
        pf.line_spacing = 1.0
        pf.line_spacing_rule = WD_LINE_SPACING.SINGLE

    doc.save(DEST)
    print(f"保存: {DEST}")


if __name__ == "__main__":
    main()
