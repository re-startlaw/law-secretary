"""分からなかったフォルダ内PDFの内容抽出。"""
import fitz

BASE = "/Users/kometaninaoki/Library/CloudStorage/GoogleDrive-n.kometani@re-startlaw.com/マイドライブ/共有用/06_分類依頼/分からなかった/"

PDFS = [
    "260508_2ﾂ‘2崎鰔紬7日.pdf",
    "260509_委任契約書（馬様） - 署名済み.pdf",
    "260511_000024640_Re-Start法律事務所様_Legalscape ご利用料金_2026-05-01ご請求分.pdf",
    "260511_20260508.pdf",
    "260511_KDDI_SEIKYU0510095114.pdf",
]

for p in PDFS:
    full = BASE + p
    print(f"===== {p} =====")
    try:
        doc = fitz.open(full)
        for i, page in enumerate(doc):
            text = page.get_text("text")
            print(f"--- page {i+1} ---")
            print(text[:2000])
            if i >= 2:
                print("...")
                break
        doc.close()
    except Exception as e:
        print(f"  ERR: {e}")
    print()
