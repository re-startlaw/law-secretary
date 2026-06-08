"""2026/5/7 ファイル分類: スプレッドシート保存ログ更新と通知メール送信。

secretary.py の既存ヘルパーを再利用する。
- 既存行があれば D列(HYPERLINK) と E列（移動先）を更新
- なければ新規行を A:E に追加
- 全件分まとめて通知メール送信
"""
import os
import sys
from datetime import datetime

sys.path.insert(0, "/Users/kometaninaoki/law-secretary")

from secretary import (
    BASE_PATH,
    hyperlink_formula,
    build_drive_link_from_id,
    drive_search_url,
    resolve_drive_folder_id,
    fetch_drive_file_meta,
    fetch_sheet_data,
    build_filename_index,
    append_to_sheet,
    update_sheet_row,
    send_notification,
    get_credentials,
)
from googleapiclient.discovery import build

SHARE = os.path.join(BASE_PATH, "共有用")

# (filename, 移動先(共有用相対))
MOVES = [
    # 馬さん
    ("20260413_古目的(1)に加えて、事案発生時の背景および現場全体の状況を示すもの（中核事実十補助.pdf", "01_事件記録/ま_馬さん/06資料"),
    ("20260413_本日朝、弟が体育館で9時20分からクロスカントリーの競技に参加しておりました。1時間目の.pdf", "01_事件記録/ま_馬さん/06資料"),
    ("20260413_本資料は、当方が現在保有する書面陳述、学校との往復メール、面談録音およびその整理記録、動.pdf", "01_事件記録/ま_馬さん/06資料"),
    ("20260414_⑥二、主要証拠抜粋（時系列）.pdf", "01_事件記録/ま_馬さん/06資料"),
    ("20260414_⑥二、主要証拠抜粋（時系列）_001.pdf", "01_事件記録/ま_馬さん/06資料"),
    ("20260428_CGAClaDetaiIsEnrolmentProcess.pdf", "01_事件記録/ま_馬さん/06資料"),
    ("20260507.pdf", "01_事件記録/ま_馬さん/06資料"),
    ("20260507_1番1号.pdf", "01_事件記録/ま_馬さん/06資料"),
    ("20260507_KISTKarenDonaldGodfrey.pdf", "01_事件記録/ま_馬さん/06資料"),
    ("20260507_KISTKarenDonaldGodfrey_001.pdf", "01_事件記録/ま_馬さん/06資料"),
    ("20260507_RECEIPT.pdf", "01_事件記録/ま_馬さん/06資料"),
    ("20260507_Stairs.pdf", "01_事件記録/ま_馬さん/06資料"),
    ("20260507_●收件人@ParentsofG9AAngelinaFeiliMaK2AGeorgeManTzar.pdf", "01_事件記録/ま_馬さん/06資料"),
    ("20260507_宗在鴫？単.pdf", "01_事件記録/ま_馬さん/06資料"),
    ("20260507_收件人OParentsofG9AAngelinaFeiliMaK2AGeorgeManTZar.pdf", "01_事件記録/ま_馬さん/06資料"),
    ("20260507_本件は、KInternationalSchoolTokyoにおいて発生した事案である。.pdf", "01_事件記録/ま_馬さん/06資料"),
    ("20260507_相談カード.pdf", "01_事件記録/ま_馬さん/06資料"),
    # 田村正宣
    ("20260501_勾留通知.pdf", "01_事件記録/た_田村正宣/02_身体拘束関係"),
    ("20260502_弁護人選任届.pdf", "01_事件記録/た_田村正宣/01_選任関係"),
    # 経理
    ("20260502.pdf", "03_経理/工具器具備品"),
    ("20260502_001.pdf", "03_経理/工具器具備品"),
    ("20260502_PM161Q.pdf", "03_経理/工具器具備品"),
    ("260505_Receipt-2184-1374-4017.pdf", "03_経理/支払手数料 (1)"),
    ("5553711356.pdf", "03_経理/支払手数料 (1)"),
    # 不要
    ("20260502_故障かな？と思ったら.pdf", "06_分類依頼/pre_trash"),
    # ファムさん
    ("260502_63_配偶者の陳述書（ファム様）.docx", "01_事件記録/ふ_ファム・ティ・フォン"),
]


def main():
    creds = get_credentials()
    sheets_service = build("sheets", "v4", credentials=creds)
    gmail_service = build("gmail", "v1", credentials=creds)
    drive_service = build("drive", "v3", credentials=creds)

    sheet_data = fetch_sheet_data(sheets_service)
    filename_index = build_filename_index(sheet_data)
    print(f"既存スプレッドシートレコード: {len(filename_index)}件")

    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    file_results = []
    pending_new = []
    updated = 0

    for filename, rel_dest in MOVES:
        abs_dest_dir = os.path.join(SHARE, rel_dest)
        folder_id = resolve_drive_folder_id(drive_service, abs_dest_dir)
        meta = (
            fetch_drive_file_meta(drive_service, folder_id, filename)
            if folder_id
            else None
        )
        if meta and meta.get("webViewLink"):
            link_url = meta["webViewLink"]
            kind = "[直リンク]"
        elif meta and meta.get("id"):
            link_url = build_drive_link_from_id(meta["id"])
            kind = "[直リンク:ID]"
        else:
            link_url = drive_search_url(filename)
            kind = "[検索]"
        d_cell = hyperlink_formula(link_url, filename)
        dest_description = rel_dest + "/"

        existing = filename_index.get(filename)
        if existing:
            update_sheet_row(sheets_service, existing["row"], d_cell, dest_description)
            updated += 1
            print(f"  既存行更新 row={existing['row']} {kind}: {filename} → {dest_description}")
        else:
            # A:保存日時 B:タイトル C:送信者 D:HYPERLINK E:移動先 F-L:空
            pending_new.append([now, "", "スキャナー", d_cell, dest_description, "", "", "", "", "", "", ""])
            print(f"  新規追加予定 {kind}: {filename} → {dest_description}")

        file_results.append((filename, dest_description))

    if pending_new:
        append_to_sheet(sheets_service, pending_new)
        print(f"新規行追加: {len(pending_new)}件")
    print(f"既存行更新: {updated}件")

    # 通知メール
    send_notification(gmail_service, file_results)
    print("通知メール送信済み")


if __name__ == "__main__":
    main()
