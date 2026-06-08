"""乙フォルダ内の全PDFの1枚目を画像化する。
- フルページPNG（目視用、横1400pxほど）
- 右上クロップPNG（三角番号確認用、高解像度）
日本語パスはこのファイル内の文字列に閉じ込める（Bashコマンド行には出さない）。
"""
from pathlib import Path
import fitz

SRC_DIR = Path(Path(".cursor/shell_path_utf8.txt").read_text().strip().splitlines()[0])
OUT = Path(".cursor/otsu_thumbs")
OUT.mkdir(parents=True, exist_ok=True)

pdfs = sorted(SRC_DIR.glob("*.pdf"))
print(f"PDF count: {len(pdfs)}")

index = []
for i, p in enumerate(pdfs):
    doc = fitz.open(str(p))
    page = doc[0]
    rect = page.rect
    # フルページ（2倍）
    mat = fitz.Matrix(2.0, 2.0)
    pix = page.get_pixmap(matrix=mat)
    full = OUT / f"{i:02d}_full.png"
    pix.save(str(full))
    # 右上クロップ（右55%〜100%, 上0〜28%）高解像度4倍
    cw0, cw1 = rect.width * 0.52, rect.width * 1.0
    ch0, ch1 = rect.height * 0.0, rect.height * 0.30
    clip = fitz.Rect(cw0, ch0, cw1, ch1)
    pix2 = page.get_pixmap(matrix=fitz.Matrix(4.0, 4.0), clip=clip)
    crop = OUT / f"{i:02d}_topright.png"
    pix2.save(str(crop))
    # 本文冒頭帯（作成日・事件名・陳述者氏名）高解像度3倍
    bclip = fitz.Rect(rect.width * 0.0, rect.height * 0.22, rect.width * 1.0, rect.height * 0.40)
    pix3 = page.get_pixmap(matrix=fitz.Matrix(3.0, 3.0), clip=bclip)
    band = OUT / f"{i:02d}_band.png"
    pix3.save(str(band))
    npages = len(doc)
    index.append((i, p.name, npages))
    doc.close()
    print(f"{i:02d}\t{p.name}\t({npages}p)")

# インデックスを書き出し
idxfile = OUT / "index.tsv"
with idxfile.open("w") as f:
    for i, name, n in index:
        f.write(f"{i:02d}\t{name}\t{n}\n")
print(f"\nsaved to {OUT}")
