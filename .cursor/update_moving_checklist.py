"""引越しチェックリストに条件付き書式を追加: 完了→グレー取消線 / 期限超過→黄色"""

import json
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

TOKEN_PATH = "secrets/token.json"
SS_ID = "1XYVZhTOywNwwm3I-zSzXUH81tfxY8tm55tvir0C5gRg"

with open(TOKEN_PATH) as f:
    token_data = json.load(f)
creds = Credentials.from_authorized_user_info(token_data)
svc = build("sheets", "v4", credentials=creds)

ss = svc.spreadsheets().get(spreadsheetId=SS_ID).execute()
sheet_id = ss["sheets"][0]["properties"]["sheetId"]
row_count = ss["sheets"][0]["properties"]["gridProperties"]["rowCount"]

full_range = {
    "sheetId": sheet_id,
    "startRowIndex": 1,
    "endRowIndex": row_count,
    "startColumnIndex": 0,
    "endColumnIndex": 7,
}

requests = []

# 1) 済み行の静的書式をクリア（条件付き書式で代替するため）
requests.append({
    "repeatCell": {
        "range": {"sheetId": sheet_id, "startRowIndex": 1, "endRowIndex": 2, "startColumnIndex": 0, "endColumnIndex": 7},
        "cell": {
            "userEnteredFormat": {
                "textFormat": {
                    "strikethrough": False,
                    "foregroundColor": {"red": 0, "green": 0, "blue": 0},
                },
            }
        },
        "fields": "userEnteredFormat.textFormat(strikethrough,foregroundColor)",
    }
})

# 2) 完了チェック → 行全体グレー＋取消線（最優先: index 0）
requests.append({
    "addConditionalFormatRule": {
        "rule": {
            "ranges": [full_range],
            "booleanRule": {
                "condition": {
                    "type": "CUSTOM_FORMULA",
                    "values": [{"userEnteredValue": "=$G2=TRUE"}],
                },
                "format": {
                    "backgroundColor": {"red": 0.93, "green": 0.93, "blue": 0.93},
                    "textFormat": {
                        "strikethrough": True,
                        "foregroundColor": {"red": 0.6, "green": 0.6, "blue": 0.6},
                    },
                },
            },
        },
        "index": 0,
    }
})

# 3) 期限超過かつ未完了 → 行全体黄色（index 1）
requests.append({
    "addConditionalFormatRule": {
        "rule": {
            "ranges": [full_range],
            "booleanRule": {
                "condition": {
                    "type": "CUSTOM_FORMULA",
                    "values": [{"userEnteredValue": '=AND(ISNUMBER($E2), $E2<TODAY(), $G2<>TRUE)'}],
                },
                "format": {
                    "backgroundColor": {"red": 1.0, "green": 0.95, "blue": 0.6},
                    "textFormat": {
                        "bold": True,
                    },
                },
            },
        },
        "index": 1,
    }
})

svc.spreadsheets().batchUpdate(
    spreadsheetId=SS_ID, body={"requests": requests}
).execute()

print("Done: conditional formatting added")
print(f"  - Checkbox ON  -> grey + strikethrough")
print(f"  - Overdue + not done -> yellow + bold")
