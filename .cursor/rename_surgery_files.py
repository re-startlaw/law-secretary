# -*- coding: utf-8 -*-
"""
手術関係フォルダ内のファイル名を、入院許可申請書の添付資料リストに合わせてリネームする。
"""
from pathlib import Path

FOLDER = Path(
    "/Users/kometaninaoki/Library/CloudStorage/GoogleDrive-n.kometani@re-startlaw.com/"
    ".shortcut-targets-by-id/1pEJgbHuaj4L4ovFEISOuq-0LeHGInxYl/"
    "大中 忠生 : 詐欺未遂/資料/大中さんと共有/その他/手術関係"
)

RENAMES = {
    "町医者；にしじま眼科.png": "資料１　にしじま眼科 医院案内.png",
    "スキャン 6.jpeg": "資料２　大阪医科薬科大学病院 診察券.jpeg",
    "スキャン 1.jpeg": "資料３　外来診療明細書（令和８年４月２８日付）.jpeg",
    "スキャン 2.jpeg": "資料４　外来診療明細書（令和８年５月１日付）.jpeg",
    "スキャン 5.jpeg": "資料５　外来診療費領収書（令和８年４月２８日付・同年５月１日付）.jpeg",
    "スキャン 4.jpeg": "資料６－１　予約票（令和８年４月２８日発行）.jpeg",
    "スキャン 3.jpeg": "資料６－２　予約票（令和８年５月１日発行）.jpeg",
    "スキャン.jpeg": "資料７　眼科手術（入院）を受ける患者様へ.jpeg",
}


def main():
    print(f"対象フォルダ: {FOLDER}")
    print()
    for old_name, new_name in RENAMES.items():
        src = FOLDER / old_name
        dst = FOLDER / new_name
        if not src.exists():
            print(f"  [skip] 元ファイルが存在しない: {old_name}")
            continue
        if dst.exists():
            print(f"  [skip] 同名既存: {new_name}")
            continue
        src.rename(dst)
        print(f"  ✓ {old_name}")
        print(f"    → {new_name}")
    print()
    print("リネーム後の一覧:")
    for f in sorted(FOLDER.iterdir()):
        if f.name.startswith("."):
            continue
        print(f"  {f.name}")


if __name__ == "__main__":
    main()
