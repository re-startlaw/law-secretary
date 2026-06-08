"""弁護士法人設立TODO スプレッドシートの整形

- A列にチェックボックス列「完了」を追加
- 全セルに折り返し（WRAP_TEXT）を適用
- 列幅を調整・行は自動リサイズ
"""
import sys, os
sys.path.insert(0, os.path.expanduser("~/law-secretary"))
from secretary import get_credentials  # noqa: E402
from googleapiclient.discovery import build  # noqa: E402

SHEET_ID = "1iUZgkTFR5FQg_ISiTSsVRTnaqgY2hX2jhuXy-P3xtmQ"


def main():
    creds = get_credentials()
    sheets = build("sheets", "v4", credentials=creds)

    meta = sheets.spreadsheets().get(spreadsheetId=SHEET_ID).execute()
    sheet0 = meta["sheets"][0]
    sheet_gid = sheet0["properties"]["sheetId"]
    row_count = sheet0["properties"]["gridProperties"]["rowCount"]
    col_count = sheet0["properties"]["gridProperties"]["columnCount"]
    print(f"sheet gid={sheet_gid} rows={row_count} cols={col_count}")

    # 1) A列にチェックボックス列を挿入
    requests = [
        {
            "insertDimension": {
                "range": {
                    "sheetId": sheet_gid,
                    "dimension": "COLUMNS",
                    "startIndex": 0,
                    "endIndex": 1,
                },
                "inheritFromBefore": False,
            }
        }
    ]
    sheets.spreadsheets().batchUpdate(
        spreadsheetId=SHEET_ID, body={"requests": requests}
    ).execute()

    # 2) A1ヘッダーに「完了」を書き込み
    sheets.spreadsheets().values().update(
        spreadsheetId=SHEET_ID,
        range="A1",
        valueInputOption="RAW",
        body={"values": [["完了"]]},
    ).execute()

    # 列数を再取得（挿入後）
    meta = sheets.spreadsheets().get(spreadsheetId=SHEET_ID).execute()
    sheet0 = meta["sheets"][0]
    col_count = sheet0["properties"]["gridProperties"]["columnCount"]
    row_count = sheet0["properties"]["gridProperties"]["rowCount"]
    print(f"after insert: rows={row_count} cols={col_count}")

    requests = []

    # 3) A2:A末尾にチェックボックスのデータ検証
    requests.append(
        {
            "setDataValidation": {
                "range": {
                    "sheetId": sheet_gid,
                    "startRowIndex": 1,
                    "endRowIndex": row_count,
                    "startColumnIndex": 0,
                    "endColumnIndex": 1,
                },
                "rule": {
                    "condition": {"type": "BOOLEAN"},
                    "strict": True,
                },
            }
        }
    )

    # 4) 全セル WRAP_TEXT・縦中央寄せ
    requests.append(
        {
            "repeatCell": {
                "range": {
                    "sheetId": sheet_gid,
                    "startRowIndex": 0,
                    "endRowIndex": row_count,
                    "startColumnIndex": 0,
                    "endColumnIndex": col_count,
                },
                "cell": {
                    "userEnteredFormat": {
                        "wrapStrategy": "WRAP",
                        "verticalAlignment": "TOP",
                    }
                },
                "fields": "userEnteredFormat(wrapStrategy,verticalAlignment)",
            }
        }
    )

    # 5) ヘッダー行を太字＋背景色＋中央寄せ＋固定
    requests.append(
        {
            "repeatCell": {
                "range": {
                    "sheetId": sheet_gid,
                    "startRowIndex": 0,
                    "endRowIndex": 1,
                    "startColumnIndex": 0,
                    "endColumnIndex": col_count,
                },
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": {
                            "red": 0.85,
                            "green": 0.92,
                            "blue": 0.83,
                        },
                        "textFormat": {"bold": True},
                        "horizontalAlignment": "CENTER",
                        "verticalAlignment": "MIDDLE",
                        "wrapStrategy": "WRAP",
                    }
                },
                "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,wrapStrategy)",
            }
        }
    )
    requests.append(
        {
            "updateSheetProperties": {
                "properties": {
                    "sheetId": sheet_gid,
                    "gridProperties": {"frozenRowCount": 1, "frozenColumnCount": 2},
                },
                "fields": "gridProperties.frozenRowCount,gridProperties.frozenColumnCount",
            }
        }
    )

    # 6) 列幅を調整
    # A: 完了(checkbox) 60px
    # B: 区分           150px
    # C: タスク内容    320px
    # D: 期限の数え方  260px
    # E: Re-Start具体日 300px
    # F: 提出先・窓口  200px
    # G: 根拠・参考リンク 240px
    # H: Re-Startメモ  360px
    widths = [60, 150, 320, 260, 300, 200, 240, 360]
    for i, w in enumerate(widths):
        if i >= col_count:
            break
        requests.append(
            {
                "updateDimensionProperties": {
                    "range": {
                        "sheetId": sheet_gid,
                        "dimension": "COLUMNS",
                        "startIndex": i,
                        "endIndex": i + 1,
                    },
                    "properties": {"pixelSize": w},
                    "fields": "pixelSize",
                }
            }
        )

    # 7) 行高さを自動リサイズ
    requests.append(
        {
            "autoResizeDimensions": {
                "dimensions": {
                    "sheetId": sheet_gid,
                    "dimension": "ROWS",
                    "startIndex": 0,
                    "endIndex": row_count,
                }
            }
        }
    )

    sheets.spreadsheets().batchUpdate(
        spreadsheetId=SHEET_ID, body={"requests": requests}
    ).execute()
    print("formatting complete")
    print(f"URL: https://docs.google.com/spreadsheets/d/{SHEET_ID}/edit")


if __name__ == "__main__":
    main()
