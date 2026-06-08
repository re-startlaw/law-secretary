"""1番1号.pdf の中身確認（鈴木七海さん用に振り分け先サブフォルダを判断）。"""
from pathlib import Path
from pdfminer.high_level import extract_text

p = Path(
    "/Users/kometaninaoki/Library/CloudStorage/"
    "GoogleDrive-n.kometani@re-startlaw.com/マイドライブ/共有用/"
    "01_事件記録/ま_馬さん/06資料/20260507_1番1号.pdf"
)
print(extract_text(str(p))[:2500])
