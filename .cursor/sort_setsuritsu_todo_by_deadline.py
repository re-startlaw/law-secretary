"""弁護士法人設立TODO スプレッドシートを期限順に並び替え

E列「Re-Start具体日」から最も早い日付を抽出してソートキーにする。
日付が抽出できない行は末尾。
"""
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
    # 完全日付（年月日）優先
    candidates = []
    for m in DATE_RE.finditer(s):
        y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
        try:
            candidates.append(date(y, mo, d))
        except ValueError:
            pass
    if candidates:
        return min(candidates)
    # 次に年月（その月初を仮の日付に）
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

    # 全データ取得
    res = sheets.spreadsheets().values().get(
        spreadsheetId=SHEET_ID, range="A1:H200"
    ).execute()
    values = res.get("values", [])
    if not values:
        print("empty")
        return
    header = values[0]
    rows = values[1:]
    # 末尾の空行除去
    rows = [r for r in rows if any((c or "").strip() for c in r)]
    print(f"header cols={len(header)} data rows={len(rows)}")

    # E列(index 4)=Re-Start具体日 でソート
    def key(r):
        v = r[4] if len(r) > 4 else ""
        return earliest_date(v)

    # 行をパディングして列数を揃える
    width = max(len(header), max((len(r) for r in rows), default=0))
    header = header + [""] * (width - len(header))
    rows = [r + [""] * (width - len(r)) for r in rows]

    sorted_rows = sorted(rows, key=key)

    # 並び替え後のキー日付をプレビュー
    for r in sorted_rows[:5]:
        print(f"  {key(r)} | {r[1]} | {r[2][:30] if len(r)>2 else ''}")
    print("  ...")
    for r in sorted_rows[-3:]:
        print(f"  {key(r)} | {r[1]} | {r[2][:30] if len(r)>2 else ''}")

    # 書き戻し
    body = {"values": [header] + sorted_rows}
    end_row = 1 + len(sorted_rows)
    # 列番号→A1
    def col_letter(n):
        s = ""
        while n > 0:
            n, r = divmod(n - 1, 26)
            s = chr(65 + r) + s
        return s

    end_col = col_letter(width)
    rng = f"A1:{end_col}{end_row}"
    sheets.spreadsheets().values().update(
        spreadsheetId=SHEET_ID,
        range=rng,
        valueInputOption="USER_ENTERED",
        body=body,
    ).execute()

    # チェックボックス（A列）の検証を再適用、念のため
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

    print(f"sorted {len(sorted_rows)} rows")
    print(f"URL: https://docs.google.com/spreadsheets/d/{SHEET_ID}/edit")


if __name__ == "__main__":
    main()
