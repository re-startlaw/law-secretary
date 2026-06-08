"""_3の差出人ブロック行間／字下げを明示的に固定して _4 を作成。

_3で生じた問題:
- 差出人ブロックの行間が広く見える（Word既定のNormalスタイルにspace_after等が乗る）。
- 行折り返し懸念のため、left_indentも少し緩める。

修正方針:
- 差出人ブロック（para 6-13）に対し:
  - alignment = LEFT
  - left_indent を Emu(3200000) = 約8.9cm に縮小（column幅 ≒ 9.1cmで折り返しゼロ）
  - first_line_indent = None
  - space_before = 0, space_after = 0
  - line_spacing = 1.0 (single)
- 他の段落は _3 のまま維持。
"""

from __future__ import annotations

import shutil
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.shared import Emu, Pt

DRIVE = Path(
    "/Users/kometaninaoki/Library/CloudStorage/GoogleDrive-n.kometani@re-startlaw.com/マイドライブ"
)
SRC = DRIVE / "共有用/01_事件記録/ま_馬さん/Wordファイル/260512電子内容証明-下書き_3.docx"
DEST = DRIVE / "共有用/01_事件記録/ま_馬さん/Wordファイル/260512電子内容証明-下書き_4.docx"

SENDER_INDENT = Emu(3200000)  # ≒ 8.9cm（テンプレ4140835から少し緩めた）


def main() -> None:
    shutil.copy2(SRC, DEST)
    doc = Document(DEST)
    paras = list(doc.paragraphs)

    # 差出人ブロック paragraphs 6-13
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
