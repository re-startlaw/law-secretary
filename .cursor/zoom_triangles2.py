"""不鮮明な三角番号(08,12,15,16)をさらに狭く高倍率でクロップ。"""
from pathlib import Path
import fitz

SRC_DIR = Path(Path(".cursor/shell_path_utf8.txt").read_text().strip().splitlines()[0])
OUT = Path(".cursor/otsu_thumbs")
pdfs = sorted(SRC_DIR.glob("*.pdf"))

# (index, (rx0,ry0,rx1,ry1)) 個別に枠を微調整
specs = {
    8:  (0.66, 0.02, 0.92, 0.12),
    12: (0.70, 0.01, 0.95, 0.11),
    15: (0.66, 0.05, 0.90, 0.16),
    16: (0.66, 0.02, 0.92, 0.12),
}
for i, (rx0, ry0, rx1, ry1) in specs.items():
    p = pdfs[i]
    doc = fitz.open(str(p))
    page = doc[0]
    r = page.rect
    clip = fitz.Rect(r.width * rx0, r.height * ry0, r.width * rx1, r.height * ry1)
    pix = page.get_pixmap(matrix=fitz.Matrix(14.0, 14.0), clip=clip)
    pix.save(str(OUT / f"{i:02d}_tri2.png"))
    doc.close()
    print(f"{i:02d}\t{p.name}")
print("done")
