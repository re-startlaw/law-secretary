"""差出人ブロックを「左寄せ＋left_indentで右側配置（各行の左端が揃う）」に修正。

_2版ではRIGHT揃えにしてしまっていたが、米谷弁護士の好みは
「左寄せ＋left_indentでブロック全体を右側に寄せる」形。
"""

from __future__ import annotations

import shutil
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Emu

DRIVE = Path(
    "/Users/kometaninaoki/Library/CloudStorage/GoogleDrive-n.kometani@re-startlaw.com/マイドライブ"
)
SRC = DRIVE / "共有用/01_事件記録/ま_馬さん/Wordファイル/260512電子内容証明-下書き_2.docx"
DEST = DRIVE / "共有用/01_事件記録/ま_馬さん/Wordファイル/260512電子内容証明-下書き_3.docx"

SENDER_INDENT = Emu(4140835)  # 鈴木七海テンプレ準拠


def main() -> None:
    assert SRC.exists(), SRC
    shutil.copy2(SRC, DEST)

    doc = Document(DEST)
    paras = list(doc.paragraphs)

    # 差出人ブロックは paragraphs 6-13 (空行+（差出人）〜TEL)
    for i in range(6, 14):
        p = paras[i]
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.left_indent = SENDER_INDENT
        p.paragraph_format.first_line_indent = None

    doc.save(DEST)
    print(f"保存: {DEST}")


if __name__ == "__main__":
    main()
