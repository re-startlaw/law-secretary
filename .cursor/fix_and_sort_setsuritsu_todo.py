"""列順を修復し、期限順に並び替える。

現状:
  0=期限の数え方  1=Re-Start具体日  2=完了  3=タスク内容
  4=提出先       5=根拠            6=メモ  7=区分
正しい順:
  0=完了 1=区分 2=タスク内容 3=期限の数え方 4=Re-Start具体日 5=提出先 6=根拠 7=メモ
"""
import sys, os, re
from datetime import date

sys.path.insert(0, os.path.expanduser("~/law-secretary"))
from secretary import get_credentials  # noqa: E402
from googleapiclient.discovery import build  # noqa: E402

SHEET_ID = "1iUZgkTFR5FQg_ISiTSsVRTnaqgY2hX2jhuXy-P3xtmQ"

# 現在のcol index -> 移動先のcol index
# 並べ替えマップ: target index -> source index
ORDER = [2, 7, 3, 0, 1, 4, 5, 6]

DATE_RE = re.compile(r"(20\d{2})[/年\-](\d{1,2})[/月\-](\d{1,2})")
YM_RE = re.compile(r"(20\d{2})[/年\-](\d{1,2})月?")
FAR = date(2099, 12, 31)


def earliest_date(s: str):
    if not s:
        return FAR
    candidates = []
    for m in DATE_RE.finditer(s):
        y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
        try:
            candidates.append(date(y, mo, d))
        except ValueError:
            pass
    if candidates:
        return min(candidates)
    for m in YM_RE.finditer(s):
        y, mo = int(m.group(1)), int(m.group(2))
        try:
            candidates.append(date(y, mo, 1))
        except ValueError:
            pass
    if candidates:
        return min(candidates)
    return FAR


def main():
    creds = get_credentials()
    sheets = build("sheets", "v4", credentials=creds)

    meta = sheets.spreadsheets().get(spreadsheetId=SHEET_ID).execute()
    sheet0 = meta["sheets"][0]
    sheet_gid = sheet0["properties"]["sheetId"]

    res = sheets.spreadsheets().values().get(
        spreadsheetId=SHEET_ID, range="A1:H200"
    ).execute()
    values = res.get("values", [])
    header = values[0] if values else []
    rows = values[1:] if values else []
    rows = [r for r in rows if any((c or "").strip() for c in r)]
    # 列パディング
    width = 8
    header = (header + [""] * width)[:width]
    rows = [(r + [""] * width)[:width] for r in rows]

    # 列を並び替え
    new_header = [header[i] for i in ORDER]
    new_rows = [[r[i] for i in ORDER] for r in rows]

    # 期限順（col4=Re-Start具体日）でソート
    new_rows.sort(key=lambda r: earliest_date(r[4]))

    # 書き戻し（先に全体をクリア）
    sheets.spreadsheets().values().clear(
        spreadsheetId=SHEET_ID, range="A1:H1000"
    ).execute()

    body = {"values": [new_header] + new_rows}
    sheets.spreadsheets().values().update(
        spreadsheetId=SHEET_ID,
        range=f"A1:H{1 + len(new_rows)}",
        valueInputOption="USER_ENTERED",
        body=body,
    ).execute()

    end_row = 1 + len(new_rows)

    # チェックボックス再設定・行高自動・列幅
    requests = [
        {
            "setDataValidation": {
                "range": {
                    "sheetId": sheet_gid,
                    "startRowIndex": 1,
                    "endRowIndex": end_row,
                    "startColumnIndex": 0,
                    "endColumnIndex": 1,
                },
                "rule": {"condition": {"type": "BOOLEAN"}, "strict": True},
            }
        },
    ]
    widths = [60, 160, 320, 260, 300, 200, 240, 360]
    for i, w in enumerate(widths):
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
    requests.append(
        {
            "repeatCell": {
                "range": {
                    "sheetId": sheet_gid,
                    "startRowIndex": 0,
                    "endRowIndex": end_row,
                    "startColumnIndex": 0,
                    "endColumnIndex": 8,
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
    requests.append(
        {
            "repeatCell": {
                "range": {
                    "sheetId": sheet_gid,
                    "startRowIndex": 0,
                    "endRowIndex": 1,
                    "startColumnIndex": 0,
                    "endColumnIndex": 8,
                },
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": {"red": 0.85, "green": 0.92, "blue": 0.83},
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
    requests.append(
        {
            "autoResizeDimensions": {
                "dimensions": {
                    "sheetId": sheet_gid,
                    "dimension": "ROWS",
                    "startIndex": 0,
                    "endIndex": end_row,
                }
            }
        }
    )
    sheets.spreadsheets().batchUpdate(
        spreadsheetId=SHEET_ID, body={"requests": requests}
    ).execute()

    print(f"rewritten {len(new_rows)} rows")
    # preview top 8
    print("\n=== sorted preview ===")
    for r in new_rows[:8]:
        print(f"{earliest_date(r[4])}  [{r[1]}]  {r[2][:40]}")
    print("...")
    for r in new_rows[-3:]:
        print(f"{earliest_date(r[4])}  [{r[1]}]  {r[2][:40]}")
    print(f"\nURL: https://docs.google.com/spreadsheets/d/{SHEET_ID}/edit")


if __name__ == "__main__":
    main()
