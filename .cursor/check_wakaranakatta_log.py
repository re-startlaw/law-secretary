"""分からなかったフォルダ内の各ファイルについて保存ログから情報を取得する。"""
import sys
sys.path.insert(0, '/Users/kometaninaoki/law-secretary')

from googleapiclient.discovery import build
from secretary import get_credentials, fetch_sheet_data, build_filename_index

def get_sheets_service():
    creds = get_credentials()
    return build("sheets", "v4", credentials=creds)

FILES = [
    "260507_ネット塾費用は5月1日に支払い済みです.jpg",
    "260508_2ﾂ‘2崎鰔紬7日.pdf",
    "260509_53_親族・知人の嘆願書（山口様）.docx",
    "260509_委任契約書（馬様） - 署名済み.pdf",
    "260511_000024640_Re-Start法律事務所様_Legalscape ご利用料金_2026-05-01ご請求分.pdf",
    "260511_20260508.pdf",
    "260511_KDDI_SEIKYU0510095114.pdf",
]

sheets = get_sheets_service()
data = fetch_sheet_data(sheets)
idx = build_filename_index(data)

for f in FILES:
    info = idx.get(f)
    if info:
        print(f"=== {f} ===")
        print(f"  row: {info['row']}")
        print(f"  subject: {info['subject']}")
        print(f"  sender: {info['sender']}")
        print(f"  dest: {info['dest']}")
    else:
        # Try partial match
        candidates = [k for k in idx.keys() if f.split("_", 1)[-1][:10] in k]
        print(f"=== {f} ===  [NOT FOUND in sheet]")
        if candidates:
            print(f"  candidates: {candidates[:5]}")
    print()
