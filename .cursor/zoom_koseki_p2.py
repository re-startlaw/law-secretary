from pathlib import Path
import fitz
SRC_DIR = Path(Path(".cursor/shell_path_utf8.txt").read_text().strip().splitlines()[0])
OUT = Path(".cursor/otsu_thumbs")
pdfs = sorted(SRC_DIR.glob("*.pdf"))
doc = fitz.open(str(pdfs[15]))
pg = doc[1]
r = pg.rect
clip = fitz.Rect(0, r.height*0.55, r.width, r.height*1.0)
pg.get_pixmap(matrix=fitz.Matrix(5.5,5.5), clip=clip).save(str(OUT/"15_p2_bottom.png"))
doc.close()
print("done")
