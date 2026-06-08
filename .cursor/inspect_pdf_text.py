"""甲1.pdfにテキストレイヤーがあるか調査する。"""
from pathlib import Path
import fitz

PATH_FILE = Path(".cursor/shell_path_utf8.txt")
src = Path(PATH_FILE.read_text().strip().splitlines()[0])

doc = fitz.open(str(src))
print(f"pages: {len(doc)}")
for i, page in enumerate(doc):
    text = page.get_text("text")
    print(f"--- page {i+1}: chars={len(text)} ---")
    if text.strip():
        print(text[:300])
    else:
        # 画像情報を確認
        imgs = page.get_images(full=True)
        print(f"  images: {len(imgs)}  (text empty)")
doc.close()
