"""ページ10/11のOCRテキストとword構造をダンプ。"""
from pathlib import Path
import fitz

src = Path(Path(".cursor/shell_path_utf8.txt").read_text().strip().splitlines()[0])
doc = fitz.open(str(src))

for pno in (10, 11):
    page = doc[pno - 1]
    print(f"\n===== page {pno} =====")
    print("--- text ---")
    print(page.get_text("text"))
    print("--- words (around IP/Nov) ---")
    for w in page.get_text("words"):
        t = w[4]
        if any(c in t for c in ("160", "237", "152", "151", "Nov", "2023", "08:00", "0600")):
            print(f"  bbox=({w[0]:.0f},{w[1]:.0f},{w[2]:.0f},{w[3]:.0f}) blk={w[5]} ln={w[6]} text={t!r}")
