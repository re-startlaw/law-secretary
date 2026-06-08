"""スキャナーからフォルダ内の判別困難PDFの先頭を抽出して表示する。"""
import sys
from pathlib import Path

try:
    from pdfminer.high_level import extract_text
except ImportError:
    print("pdfminer未インストール")
    sys.exit(1)

BASE = Path(
    "/Users/kometaninaoki/Library/CloudStorage/"
    "GoogleDrive-n.kometani@re-startlaw.com/マイドライブ/共有用/06_分類依頼/スキャナーから"
)

TARGETS = [
    "20260507_RECEIPT.pdf",
    "20260507_相談カード.pdf",
    "20260507_1番1号.pdf",
    "20260507.pdf",
    "20260413_本日朝、弟が体育館で9時20分からクロスカントリーの競技に参加しておりました。1時間目の.pdf",
    "20260507_本件は、KInternationalSchoolTokyoにおいて発生した事案である。.pdf",
]

for name in TARGETS:
    p = BASE / name
    print(f"\n=== {name} ===")
    if not p.exists():
        print("ファイルなし")
        continue
    try:
        txt = extract_text(str(p))
        head = (txt or "")[:800]
        print(head if head.strip() else "(本文抽出できず・画像PDFの可能性)")
    except Exception as e:
        print(f"err: {e}")
