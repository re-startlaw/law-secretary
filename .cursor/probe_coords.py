"""ページ10/11の対象テキスト座標を推定。
ページ10の「160.237.152.151」表示行と、ページ11上端の「15/Nov/2023:08:00:51」の
y座標範囲を、OCR wordsから推定する。
"""
from pathlib import Path
import fitz

src = Path(Path(".cursor/shell_path_utf8.txt").read_text().strip().splitlines()[0])
doc = fitz.open(str(src))

# ページ10: 「IPアドレスが」を含むwordのy座標を見つけて、その下の領域を特定
page = doc[9]
print(f"page10 size: {page.rect}")
words = page.get_text("words")
ip_kw_words = [w for w in words if "IPアドレス" in w[4] or "ＩＰアドレス" in w[4] or "1Pアドレス" in w[4]]
for w in ip_kw_words:
    print(f"  IP-kw: bbox=({w[0]:.0f},{w[1]:.0f},{w[2]:.0f},{w[3]:.0f}) text={w[4]!r}")
print("---")
# 「160.」「2 3 7」「1 5 2」「1 5 1」 single-digit/dot のwordをすべて表示
for w in words:
    t = w[4].strip()
    if t in ("160.", "2", "3", "7", "1", "5", "2 3 7", "1 5 2", "1 5 1") or "160" in t:
        print(f"  digit: bbox=({w[0]:.0f},{w[1]:.0f},{w[2]:.0f},{w[3]:.0f}) text={t!r}")

# ページ11: 「15/NOV」「Nov」「2023」「08」「0600」のwords
page = doc[10]
print(f"\npage11 size: {page.rect}")
words = page.get_text("words")
for w in words:
    t = w[4]
    if any(x in t for x in ("Nov", "NOV", "2023", "0600", "08", "51", "00")):
        print(f"  ts: bbox=({w[0]:.0f},{w[1]:.0f},{w[2]:.0f},{w[3]:.0f}) text={t!r}")
