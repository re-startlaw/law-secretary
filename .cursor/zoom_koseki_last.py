from pathlib import Path
import fitz
SRC_DIR = Path(Path(".cursor/shell_path_utf8.txt").read_text().strip().splitlines()[0])
OUT = Path(".cursor/otsu_thumbs")
pdfs = sorted(SRC_DIR.glob("*.pdf"))
p = pdfs[15]
doc = fitz.open(str(p))
print("pages:", len(doc))
last = doc[len(doc)-1]
r = last.rect
pix = last.get_pixmap(matrix=fitz.Matrix(3.5,3.5))
pix.save(str(OUT/"15_lastpage.png"))
clip = fitz.Rect(0, r.height*0.55, r.width, r.height*1.0)
pix2 = last.get_pixmap(matrix=fitz.Matrix(6.0,6.0), clip=clip)
pix2.save(str(OUT/"15_last_bottom.png"))
doc.close()
print("done")
