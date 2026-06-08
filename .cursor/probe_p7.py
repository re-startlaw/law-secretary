from pathlib import Path
import unicodedata
import re
import fitz

src = Path(Path(".cursor/shell_path_utf8.txt").read_text().strip().splitlines()[0])
doc = fitz.open(str(src))


def nfkc_each(text):
    out = []
    for c in text:
        n = unicodedata.normalize("NFKC", c)
        out.append(n if len(n) == 1 else c)
    return "".join(out)


for pno in (7, 10):
    page = doc[pno - 1]
    words = page.get_text("words")
    print(f"===== page {pno}: showing 160/237/152/151 words =====")
    for w in words:
        if any(x in w[4] for x in ("160", "237", "152", "151")):
            print(f"  bbox=({w[0]:.0f},{w[1]:.0f},{w[2]:.0f},{w[3]:.0f}) blk={w[5]} ln={w[6]} text={w[4]!r}")
    # ページ7のIP行を再構築 (y_rows相当)
    print("--- y_rows analysis ---")
    y_rows = {}
    for w in words:
        y_center = (w[1] + w[3]) / 2
        y_key = round(y_center)
        y_rows.setdefault(y_key, []).append(w)
    # ip行候補
    for y_key in sorted(y_rows.keys()):
        row = y_rows[y_key]
        row.sort(key=lambda w: w[0])
        text = " ".join(w[4] for w in row)
        norm = nfkc_each(text)
        m = re.search(r"(\d{1,3})[.,\s\-]+(\d{1,3})[.,\s\-]+(\d{1,3})[.,\s\-]+(\d{1,3})", norm)
        if m:
            parts = [int(m.group(i)) for i in range(1, 5)]
            print(f"  y={y_key} match={m.group()!r} parts={parts}")
            print(f"     text={text!r}")
