#!/usr/bin/env python3
"""フェーズ0テストデータ生成（/tmp/kiroku_dev_case/）。

- (a) reportlab + HeiseiMin-W3 の日本語テキストPDF（埋め込みテキストあり）
- (b) PyMuPDF で日本語を画像化したテキスト層なしPDF（OCR経路用）
- (c) 200ページ超PDF（仮想化レンダリングの性能検証用）
- (d) ダミーmp4（メディア一覧表示用）

ファイル名は 甲・乙・弁・枝番・全角数字・符号無し を網羅する。
"""

from __future__ import annotations

import sys
from pathlib import Path

DEST = Path("/tmp/kiroku_dev_case")


def make_text_pdf(path: Path, lines: list[str]) -> None:
    from reportlab.lib.pagesizes import A4
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.cidfonts import UnicodeCIDFont
    from reportlab.pdfgen import canvas

    pdfmetrics.registerFont(UnicodeCIDFont("HeiseiMin-W3"))
    c = canvas.Canvas(str(path), pagesize=A4)
    width, height = A4
    c.setFont("HeiseiMin-W3", 14)
    y = height - 80
    for line in lines:
        c.drawString(72, y, line)
        y -= 28
        if y < 80:
            c.showPage()
            c.setFont("HeiseiMin-W3", 14)
            y = height - 80
    c.showPage()
    c.save()


def make_many_pages_pdf(path: Path, n_pages: int) -> None:
    from reportlab.lib.pagesizes import A4
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.cidfonts import UnicodeCIDFont
    from reportlab.pdfgen import canvas

    pdfmetrics.registerFont(UnicodeCIDFont("HeiseiMin-W3"))
    c = canvas.Canvas(str(path), pagesize=A4)
    width, height = A4
    for i in range(1, n_pages + 1):
        c.setFont("HeiseiMin-W3", 16)
        c.drawString(72, height - 100, f"長文資料 第{i}ページ")
        c.setFont("HeiseiMin-W3", 12)
        c.drawString(72, height - 140, f"これは性能検証用のページです。通し番号 {i}。")
        c.showPage()
    c.save()


def make_image_only_pdf(path: Path, text: str) -> None:
    """テキスト層を持たない画像PDF（OCR経路用）。

    PyMuPDF で空ページに文字を描画し、そのページをラスタライズした画像だけを
    別PDFに貼る。これにより抽出可能なテキスト層が存在しない。
    """
    import fitz

    src = fitz.open()
    page = src.new_page()
    page.insert_text((72, 120), text, fontsize=20, fontname="japan")
    pix = page.get_pixmap(dpi=200)
    src.close()

    out = fitz.open()
    rect = fitz.Rect(0, 0, pix.width, pix.height)
    opage = out.new_page(width=pix.width, height=pix.height)
    opage.insert_image(rect, pixmap=pix)
    out.save(str(path))
    out.close()


def make_dummy_mp4(path: Path) -> None:
    # ffmpeg 等に依存せず、識別可能な最小バイト列を書く（再生はQuickTime想定）。
    path.write_bytes(b"\x00\x00\x00\x18ftypmp42" + b"\x00" * 256)


def main() -> int:
    DEST.mkdir(parents=True, exist_ok=True)

    make_text_pdf(
        DEST / "甲1 被害届 2026-01-15.pdf",
        [
            "被害届",
            "令和６年１月１５日、被害者は供述調書のとおり申告した。",
            "実況見分調書の作成日は２０２６年１月１６日である。",
            "事件番号 R6 第123号。",
        ],
    )
    make_text_pdf(
        DEST / "甲2の1 実況見分調書.pdf",
        ["実況見分調書", "現場の状況を記録した。供述の裏付けとなる。"],
    )
    make_text_pdf(
        DEST / "乙3 員面調書＠田中太郎.pdf",
        ["員面調書", "田中太郎の供述を録取した書面である。"],
    )
    make_text_pdf(
        DEST / "弁１０ 意見書.pdf",
        ["意見書", "弁護人の意見を述べる。情状について。"],
    )
    make_text_pdf(
        DEST / "資料メモ.pdf",
        ["参考資料", "符号の付されていない一般資料。"],
    )
    make_many_pages_pdf(DEST / "甲20 長文資料.pdf", 210)
    make_image_only_pdf(DEST / "甲5 スキャン画像のみ.pdf", "スキャン文書 供述 被害")
    make_dummy_mp4(DEST / "弁2 取調べ録画.mp4")

    print(f"wrote test data: {DEST}")
    for p in sorted(DEST.iterdir()):
        print(f"  {p.name} ({p.stat().st_size} bytes)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
