"""タスク内容（C列）が空の行を削除して、48行に整える。"""
import sys, os, re
from datetime import date

sys.path.insert(0, os.path.expanduser("~/law-secretary"))
from secretary import get_credentials  # noqa: E402
from googleapiclient.discovery import build  # noqa: E402

SHEET_ID = "1iUZgkTFR5FQg_ISiTSsVRTnaqgY2hX2jhuXy-P3xtmQ"

DATE_RE = re.compile(r"(20\d{2})[/年\-](\d{1,2})[/月\-](\d{1,2})")
YM_RE = re.compile(r"(20\d{2})[/年\-](\d{1,2})月?")
FAR = date(2099, 12, 31)


def earliest_date(s: str):
    if not s:
        return FAR
    for m in DATE_RE.finditer(s):
        try:
            return date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        except ValueError:
            pass
    for m in YM_RE.finditer(s):
        try:
            return date(int(m.group(1)), int(m.group(2)), 1)
        except ValueError:
            pass
    return FAR


def main():
    creds = get_credentials()
    sheets = build("sheets", "v4", credentials=creds)
    meta = sheets.spreadsheets().get(spreadsheetId=SHEET_ID).execute()
    sheet0 = meta["sheets"][0]
    sheet_gid = sheet0["properties"]["sheetId"]

    res = sheets.spreadsheets().values().get(
        spreadsheetId=SHEET_ID, range="A1:H300"
    ).execute()
    values = res.get("values", [])
    header = (values[0] + [""] * 8)[:8]
    rows = []
    for r in values[1:]:
        r = (r + [""] * 8)[:8]
        # タスク内容(C, index 2) が空の行は捨てる
        if (r[2] or "").strip():
            rows.append(r)
    rows.sort(key=lambda r: earliest_date(r[4]))
    print(f"valid rows: {len(rows)}")

    # 全クリア
    sheets.spreadsheets().values().clear(
        spreadsheetId=SHEET_ID, range="A1:H1000"
    ).execute()

    body = {"values": [header] + rows}
    end_row = 1 + len(rows)
    sheets.spreadsheets().values().update(
        spreadsheetId=SHEET_ID,
        range=f"A1:H{end_row}",
        valueInputOption="USER_ENTERED",
        body=body,
    ).execute()

    # チェックボックス再付与
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
        # 余り行のチェックボックスは外す
        {
            "setDataValidation": {
                "range": {
                    "sheetId": sheet_gid,
                    "startRowIndex": end_row,
                    "endRowIndex": 1000,
                    "startColumnIndex": 0,
                    "endColumnIndex": 1,
                },
            }
        },
        {
            "autoResizeDimensions": {
                "dimensions": {
                    "sheetId": sheet_gid,
                    "dimension": "ROWS",
                    "startIndex": 0,
                    "endIndex": end_row,
                }
            }
        },
    ]
    sheets.spreadsheets().batchUpdate(
        spreadsheetId=SHEET_ID, body={"requests": requests}
    ).execute()

    print("\n=== sorted preview (top 12) ===")
    for r in rows[:12]:
        print(f"{earliest_date(r[4])}  [{r[1]}]  {r[2][:50]}")
    print(f"\nURL: https://docs.google.com/spreadsheets/d/{SHEET_ID}/edit")


if __name__ == "__main__":
    main()
