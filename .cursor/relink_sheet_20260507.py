"""保存ログを新ファイル名・新パスで再リンクする。

旧ファイル名で行を検索 → 新HYPERLINK式と新dest descriptionに更新する。
"""
import os
import sys
import time

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
    update_sheet_row,
    get_credentials,
)
from googleapiclient.discovery import build

SHARE = os.path.join(BASE_PATH, "共有用")

# (旧ファイル名, 新ファイル名, 新移動先 共有用相対)
RELINKS = [
    # 鈴木七海さんへ
    ("20260507_1番1号.pdf",
     "1番1号書留通知（封筒）.pdf",
     "01_事件記録/す_鈴木七海/03連絡文書"),
    # 馬さん 260507受領 配下
    ("20260413_古目的(1)に加えて、事案発生時の背景および現場全体の状況を示すもの（中核事実十補助.pdf",
     "証拠一覧と各証拠の目的説明.pdf",
     "01_事件記録/ま_馬さん/06資料/260507受領"),
    ("20260413_本日朝、弟が体育館で9時20分からクロスカントリーの競技に参加しておりました。1時間目の.pdf",
     "Angelina書面陳述（4月13日付）.pdf",
     "01_事件記録/ま_馬さん/06資料/260507受領"),
    ("20260413_本資料は、当方が現在保有する書面陳述、学校との往復メール、面談録音およびその整理記録、動.pdf",
     "事案整理書面（依頼者作成）.pdf",
     "01_事件記録/ま_馬さん/06資料/260507受領"),
    ("20260414_⑥二、主要証拠抜粋（時系列）.pdf",
     "面談録音主要抜粋(4月14日学校面談).pdf",
     "01_事件記録/ま_馬さん/06資料/260507受領"),
    ("20260414_⑥二、主要証拠抜粋（時系列）_001.pdf",
     "面談録音主要抜粋(4月14日学校面談)_2.pdf",
     "01_事件記録/ま_馬さん/06資料/260507受領"),
    ("20260428_CGAClaDetaiIsEnrolmentProcess.pdf",
     "CGA入学プロセス資料.pdf",
     "01_事件記録/ま_馬さん/06資料/260507受領"),
    ("20260507.pdf",
     "証拠映像説明書面.pdf",
     "01_事件記録/ま_馬さん/06資料/260507受領"),
    ("20260507_KISTKarenDonaldGodfrey.pdf",
     "学校メール_KarenDonaldGodfreyスチューデントケアコーディネーター.pdf",
     "01_事件記録/ま_馬さん/06資料/260507受領"),
    ("20260507_KISTKarenDonaldGodfrey_001.pdf",
     "学校メール_KarenDonaldGodfreyスチューデントケアコーディネーター_2.pdf",
     "01_事件記録/ま_馬さん/06資料/260507受領"),
    ("20260507_RECEIPT.pdf",
     "医療費領収書(AmericanClinicTokyo_4月16日精神科).pdf",
     "01_事件記録/ま_馬さん/06資料/260507受領"),
    ("20260507_Stairs.pdf",
     "事案発生現場の位置関係図.pdf",
     "01_事件記録/ま_馬さん/06資料/260507受領"),
    ("20260507_●收件人@ParentsofG9AAngelinaFeiliMaK2AGeorgeManTzar.pdf",
     "学校メール_MarkCowe(4月13日事案当日).pdf",
     "01_事件記録/ま_馬さん/06資料/260507受領"),
    ("20260507_宗在鴫？単.pdf",
     "WeChat記録(証拠動画取得経緯).pdf",
     "01_事件記録/ま_馬さん/06資料/260507受領"),
    ("20260507_收件人OParentsofG9AAngelinaFeiliMaK2AGeorgeManTZar.pdf",
     "学校法務顧問メール_寺井隼人(4月20日).pdf",
     "01_事件記録/ま_馬さん/06資料/260507受領"),
    ("20260507_本件は、KInternationalSchoolTokyoにおいて発生した事案である。.pdf",
     "案件概要.pdf",
     "01_事件記録/ま_馬さん/06資料/260507受領"),
    ("20260507_相談カード.pdf",
     "相談カード.pdf",
     "01_事件記録/ま_馬さん/06資料/260507受領"),
]

# 注: 上のRELINKSの新ファイル名に書いた「（）」は、既にrefileスクリプトで全角括弧のまま
# 移動しているので、Drive側のファイル名と一致させるため全角括弧で再構成する。
RELINKS = [
    (a, b.replace("(", "（").replace(")", "）"), c)
    for (a, b, c) in RELINKS
]


def main():
    creds = get_credentials()
    sheets_service = build("sheets", "v4", credentials=creds)
    drive_service = build("drive", "v3", credentials=creds)

    sheet_data = fetch_sheet_data(sheets_service)
    filename_index = build_filename_index(sheet_data)

    not_found = []
    updated = 0
    for old, new, rel in RELINKS:
        existing = filename_index.get(old)
        if not existing:
            # 旧ファイル名のままの行が前回追加で残っているはずなので念のため新名でも探す
            existing = filename_index.get(new)
        if not existing:
            not_found.append(old)
            continue
        abs_dest = os.path.join(SHARE, rel)
        folder_id = resolve_drive_folder_id(drive_service, abs_dest)
        meta = fetch_drive_file_meta(drive_service, folder_id, new) if folder_id else None
        # 同期遅延に備えて1度リトライ
        if not meta:
            time.sleep(2)
            meta = fetch_drive_file_meta(drive_service, folder_id, new) if folder_id else None
        if meta and meta.get("webViewLink"):
            link_url = meta["webViewLink"]
            kind = "[直リンク]"
        elif meta and meta.get("id"):
            link_url = build_drive_link_from_id(meta["id"])
            kind = "[ID]"
        else:
            link_url = drive_search_url(new)
            kind = "[検索フォールバック]"
        d_cell = hyperlink_formula(link_url, new)
        dest_description = rel + "/"
        update_sheet_row(sheets_service, existing["row"], d_cell, dest_description)
        updated += 1
        print(f"  row={existing['row']} {kind}: {old}\n      -> {new} ({dest_description})")

    print(f"\n更新: {updated}件")
    if not_found:
        print(f"未検出: {len(not_found)}件")
        for n in not_found:
            print(f"  - {n}")


if __name__ == "__main__":
    main()
