"""不鮮明な三角番号を超高解像度で再クロップする。"""
from pathlib import Path
import fitz

SRC_DIR = Path(Path(".cursor/shell_path_utf8.txt").read_text().strip().splitlines()[0])
OUT = Path(".cursor/otsu_thumbs")
pdfs = sorted(SRC_DIR.glob("*.pdf"))

targets = [4, 8, 12, 15, 16, 19]
for i in targets:
    p = pdfs[i]
    doc = fitz.open(str(p))
    page = doc[0]
    r = page.rect
    # 右肩三角だけ：右60-100%, 上0-18%
    clip = fitz.Rect(r.width * 0.60, r.height * 0.0, r.width * 1.0, r.height * 0.18)
    pix = page.get_pixmap(matrix=fitz.Matrix(8.0, 8.0), clip=clip)
    pix.save(str(OUT / f"{i:02d}_tri.png"))
    doc.close()
    print(f"{i:02d}\t{p.name}")
print("done")
