"""戸籍全部事項証明書(index15)の発行日スタンプ部分を拡大。"""
from pathlib import Path
import fitz

SRC_DIR = Path(Path(".cursor/shell_path_utf8.txt").read_text().strip().splitlines()[0])
OUT = Path(".cursor/otsu_thumbs")
pdfs = sorted(SRC_DIR.glob("*.pdf"))
p = pdfs[15]
doc = fitz.open(str(p))
page = doc[0]
r = page.rect
# 下部全幅（発行番号・発行日・区長印が並ぶ最下部 82-100%）
clip = fitz.Rect(r.width * 0.0, r.height * 0.80, r.width * 1.0, r.height * 1.0)
pix = page.get_pixmap(matrix=fitz.Matrix(5.0, 5.0), clip=clip)
pix.save(str(OUT / "15_bottom.png"))
# 右下の認証スタンプだけ拡大
clip2 = fitz.Rect(r.width * 0.62, r.height * 0.84, r.width * 1.0, r.height * 0.99)
pix2 = page.get_pixmap(matrix=fitz.Matrix(9.0, 9.0), clip=clip2)
pix2.save(str(OUT / "15_stamp.png"))
doc.close()
print("done")
