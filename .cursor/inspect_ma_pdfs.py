"""馬さん06資料配下PDFの先頭テキストを確認して、適切なリネーム判断材料にする。"""
import sys
from pathlib import Path

from pdfminer.high_level import extract_text

BASE = Path(
    "/Users/kometaninaoki/Library/CloudStorage/"
    "GoogleDrive-n.kometani@re-startlaw.com/マイドライブ/共有用/"
    "01_事件記録/ま_馬さん/06資料"
)

NAMES = [
    "20260413_古目的(1)に加えて、事案発生時の背景および現場全体の状況を示すもの（中核事実十補助.pdf",
    "20260413_本資料は、当方が現在保有する書面陳述、学校との往復メール、面談録音およびその整理記録、動.pdf",
    "20260414_⑥二、主要証拠抜粋（時系列）.pdf",
    "20260414_⑥二、主要証拠抜粋（時系列）_001.pdf",
    "20260428_CGAClaDetaiIsEnrolmentProcess.pdf",
    "20260507_KISTKarenDonaldGodfrey.pdf",
    "20260507_KISTKarenDonaldGodfrey_001.pdf",
    "20260507_Stairs.pdf",
    "20260507_●收件人@ParentsofG9AAngelinaFeiliMaK2AGeorgeManTzar.pdf",
    "20260507_收件人OParentsofG9AAngelinaFeiliMaK2AGeorgeManTZar.pdf",
]

for n in NAMES:
    p = BASE / n
    print(f"\n=== {n} ===")
    if not p.exists():
        print("ファイルなし")
        continue
    try:
        txt = extract_text(str(p)) or ""
        head = txt[:1000]
        print(head if head.strip() else "(画像PDF・本文抽出できず)")
    except Exception as e:
        print(f"err: {e}")
