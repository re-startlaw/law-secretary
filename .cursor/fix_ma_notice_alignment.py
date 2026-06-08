"""馬さん通知書ドラフトの段落揃え・字下げを整える。

問題:
- テンプレート（鈴木七海）の段落6-18に left_indent=4140835 EMU（≒4.6cm）が入っており、
  上書きで差出人ブロック・タイトル・本文先頭まで右にずれていた。
- テンプレート段落33,34,37,(他)に first_line_indent=133350 EMU が混在し、章/項の
  字下げが揃わなかった。

修正方針:
- 差出人ブロック（段落6-13）: 左字下げをクリアし、alignment=RIGHT に統一（伝統的レイアウト）。
- タイトル・本文（段落14以降）: left_indent / first_line_indent を全て None に戻す。
  字下げは本文内の全角空白で表現済み。
- 「以上」はそのまま RIGHT。

別バージョンとして _2 で保存。
"""

from __future__ import annotations

import shutil
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH

DRIVE = Path(
    "/Users/kometaninaoki/Library/CloudStorage/GoogleDrive-n.kometani@re-startlaw.com/マイドライブ"
)
SRC = DRIVE / "共有用/01_事件記録/ま_馬さん/Wordファイル/260512電子内容証明-下書き.docx"
DEST = DRIVE / "共有用/01_事件記録/ま_馬さん/Wordファイル/260512電子内容証明-下書き_2.docx"


def main() -> None:
    assert SRC.exists(), SRC
    shutil.copy2(SRC, DEST)

    doc = Document(DEST)
    paras = list(doc.paragraphs)

    # 段落インデックス（_1版基準）
    # 0   日付 RIGHT
    # 1-5 宛先（寺井） LEFT
    # 6-13 差出人ブロック → RIGHT に変更、字下げクリア
    # 14-16 タイトル CENTER（字下げクリア）
    # 17 以降 本文 LEFT（字下げクリア）
    # 66 以上 RIGHT

    for i, p in enumerate(paras):
        pf = p.paragraph_format
        # 全段落の字下げをクリア
        pf.left_indent = None
        pf.right_indent = None
        pf.first_line_indent = None

        if i == 0:
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        elif 1 <= i <= 5:
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        elif 6 <= i <= 13:
            # 差出人ブロックは右寄せ
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        elif 14 <= i <= 16:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        elif i == len(paras) - 1:
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        else:
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT

    doc.save(DEST)
    print(f"保存: {DEST}")


if __name__ == "__main__":
    main()
