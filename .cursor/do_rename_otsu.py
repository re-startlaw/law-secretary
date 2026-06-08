# -*- coding: utf-8 -*-
"""乙号証PDFを確定した新ファイル名へリネームする。
原ファイル名→新ファイル名の辞書で対応（インデックスずれ防止）。送付書は対象外。
"""
from pathlib import Path

SRC_DIR = Path(Path(".cursor/shell_path_utf8.txt").read_text().strip().splitlines()[0])

MAP = {
    "19920326_f上とうZE司司.pdf":        "乙30 供述調書 20260416 ＠木村広希.pdf",
    "20260222_供述調書.pdf":              "乙32 供述調書 20260330 ＠木村広希.pdf",
    "20260223_住と迂匠司司.pdf":          "乙39 供述調書 20260430 ＠三井隆将.pdf",
    "20260330_供述調書.pdf":              "乙25 供述調書 20260330 ＠熊野正起.pdf",
    "20260331_供述調書.pdf":              "乙7 供述調書 20260331 ＠田村正宣.pdf",
    "20260413_供述調書.pdf":              "乙31 供述調書 20260413 ＠木村広希.pdf",
    "20260414_供述調書.pdf":              "乙19 供述調書 20260414 ＠奥田和也.pdf",
    "20260416_佳進．ラゴミ司司薑壽.pdf":   "乙16 供述調書 20260416 ＠奥田和也.pdf",
    "20260416_蓼はこラZiさ言司.pdf":       "乙17 供述調書 20260416 ＠奥田和也.pdf",
    "20260417_供述。調書.pdf":             "乙40 供述調書 20260417 ＠三井隆将.pdf",
    "20260426_供述調書.pdf":              "乙34 供述調書 20260426 ＠木村広希.pdf",
    "20260427_供述調書.pdf":              "乙3 供述調書 20260427 ＠鈴木義明.pdf",
    "20260604_f生う重重司司.pdf":          "乙2 供述調書 20260430 ＠鈴木義明.pdf",
    "20260604_はとうzE司司.pdf":           "乙10 供述調書 20260417 ＠田村正宣.pdf",
    "20260604_イラ生う重重司司.pdf":       "乙33 供述調書 20260428 ＠木村広希.pdf",
    "20260604_且睾、ユ〆1弧盈工叫訂剥.pdf": "乙13 戸籍全部事項証明書 20260311 ＠大阪市住吉区長.pdf",
    "20260604_佳と五重司司言壽.pdf":       "乙12 供述調書 20260428 ＠田村正宣.pdf",
    "20260604_参はと五zE司司三妻.pdf":     "乙8 供述調書 20260416 ＠田村正宣.pdf",
    "20260604_性と式zE・‘司司二言.pdf":    "乙20 供述調書 20260428 ＠奥田和也.pdf",
    "20260604_産うZE司司.pdf":            "乙11 供述調書 20260417 ＠田村正宣.pdf",
    # 260525_送付書.pdf はリネームしない
}

# 重複チェック
news = list(MAP.values())
assert len(news) == len(set(news)), "新ファイル名に重複あり"

done, skipped = [], []
for old, new in MAP.items():
    src = SRC_DIR / old
    dst = SRC_DIR / new
    if not src.exists():
        skipped.append((old, "原ファイルなし"))
        continue
    if dst.exists():
        skipped.append((old, f"宛先が既存: {new}"))
        continue
    src.rename(dst)
    done.append((old, new))

print(f"=== リネーム完了 {len(done)}件 ===")
for old, new in done:
    print(f"  {old}\n   -> {new}")
if skipped:
    print(f"\n=== スキップ {len(skipped)}件 ===")
    for old, why in skipped:
        print(f"  {old}: {why}")
