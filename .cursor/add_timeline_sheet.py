#!/usr/bin/env python3
"""既存の鈴木七海整理表に「時系列」シートを追加する"""

import json
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

# 既存の鈴木七海整理表
SPREADSHEET_ID = "1FRfGsvId9Vyd0ioQ0P2DSxzYlz33KKRsbauEWOl1zk0"
# 不要になったコピー（削除対象）
COPY_ID = "1D14pBxFEfcwQAXVzzClTbOQF9ulxuy2-Lmg30ySiB-4"

TOKEN_PATH = "/Users/kometaninaoki/law-secretary/secrets/token.json"
CREDS_PATH = "/Users/kometaninaoki/law-secretary/secrets/oauth_credentials.json"

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]


def get_credentials():
    with open(TOKEN_PATH) as f:
        token_data = json.load(f)
    with open(CREDS_PATH) as f:
        creds_data = json.load(f)
    installed = creds_data.get("installed", creds_data.get("web", {}))
    creds = Credentials(
        token=token_data["token"],
        refresh_token=token_data["refresh_token"],
        token_uri=installed["token_uri"],
        client_id=installed["client_id"],
        client_secret=installed["client_secret"],
        scopes=SCOPES,
    )
    return creds


def main():
    creds = get_credentials()
    sheets_service = build("sheets", "v4", credentials=creds)
    drive_service = build("drive", "v3", credentials=creds)

    # 1. 「時系列」シートを追加
    add_sheet_request = {
        "requests": [
            {
                "addSheet": {
                    "properties": {
                        "title": "時系列",
                        "index": 0,
                    }
                }
            }
        ]
    }
    result = sheets_service.spreadsheets().batchUpdate(
        spreadsheetId=SPREADSHEET_ID, body=add_sheet_request
    ).execute()
    sheet_id = result["replies"][0]["addSheet"]["properties"]["sheetId"]
    print(f"時系列シート追加完了 (sheetId={sheet_id})")

    # 2. 時系列データ（大塚氏の新子誕生・婚姻等を含む）
    # 資料から判明した事実:
    # - 鈴木様メモ: 「2月に子どもが生まれることに伴い」
    # - 1/19に大塚氏から「配偶者の間に子が産まれる」連絡
    # - 2/10に大塚氏「弁護士に相談。妻の扶養での減額は不可。子の分のみ考慮」
    # - 3/23服部弁護士書面: 「大塚様の子どもが誕生したことに伴い」（過去形）
    # - 養育費計算書: 2026年1月分と2月以降で扶養数が変わる（1月:子1人、2月以降:子2人）
    timeline_data = [
        ["日付", "出来事", "備考"],
        ["令和2年12月24日", "鈴木陽翔くん出生", "鈴木七海様の長男"],
        ["令和3年6月末", "出産費用精算完了", "大塚氏が半額28万6575円を支払済（公正証書第2条）"],
        ["令和3年7月16日", "養育費支払等合意公正証書作成", "新宿御苑前公証役場・公証人青野洋士\n甲代理人：木村佐生弁護士（奥野法律特許事務所）\n乙代理人：服部咲弁護士"],
        ["令和3年7月16日〜23日", "大塚氏が認知届に署名押印→鈴木様が所沢市役所へ提出", "公正証書第1条2項"],
        ["令和4年2〜3月頃", "令和3年度の源泉徴収票に基づく養育費精算", "大塚氏 推定360万→実際384万／鈴木様 推定184万→実際?万\n推定養育費4.5万→実際4万（鈴木様が6万円返金）"],
        ["令和5年2〜3月頃", "令和4年度の源泉徴収票に基づく養育費精算", "大塚氏 推定500万→実際418万／鈴木様 推定320万→実際390万\n推定養育費4万→実際3万（鈴木様が12万円返金）"],
        ["令和6年2〜3月頃", "令和5年度の源泉徴収票に基づく養育費精算", "大塚氏 推定700万→実際698万／鈴木様 推定415万→実際?万\n推定養育費6万→実際6万"],
        ["令和7年2〜3月頃", "令和6年度の源泉徴収票に基づく養育費精算", "大塚氏 推定700万→実際1184万／鈴木様 推定415万→実際416万\n推定養育費6万→実際5.7万（計算ツール使用）\n大塚氏の年収が推定を大幅に上回り、差額を鈴木様に支払い"],
        ["令和7年3月18日", "大塚氏がサイトの計算ツールでの養育費算定を提案", "理由：「算定表のどこのマス目を見るかで額が変わるため」\n実際は計算ツール使用で月3000円減額が可能だったため\n鈴木様は拒否後、最終的に譲歩"],
        ["令和8年1月19日", "大塚氏から「配偶者との間に子が産まれる」旨の連絡", "「2月に子どもが生まれることに伴い、1月から妻も仕事をしておらず、扶養する人が2人になるため養育費が1月から減額になる」と通告\n鈴木様は「まだ産まれていない1月からの減額は納得できない。弁護士を通して」と回答"],
        ["令和8年1月19日〜2月9日", "大塚氏が弁護士を通さず直接交渉を継続", "鈴木様は毎回「弁護士を通して」と求めるが、大塚氏がスルー\n大塚氏「育休のお金は非課税で所得にカウントされない。妻は養育費計算の子に相当する」と主張"],
        ["令和8年2月頃", "大塚氏の新子（第2子）誕生", "大塚氏と配偶者との間の子。養育費計算上、2月以降は扶養家族が2人（子2名）に"],
        ["令和8年2月10日", "大塚氏が弁護士に相談後「妻の扶養での減額は不可」と連絡", "「子どもの分のみを考慮して養育費を決める」と方針転換\n約1か月間の不要なやり取りの末、当初から弁護士に相談していれば解決した事案"],
        ["令和8年3月13日", "米谷弁護士が鈴木様と接見（初回面談？）", "18:30〜20:00（1時間30分）・日当10万円\n※整理表の日当シートに記録あり"],
        ["令和8年3月23日", "服部咲弁護士（大塚氏代理人）から鈴木様宛「御連絡」書面送付", "内容：\n①養育費月額 月11万0358円→月7万9813円へ減額通告（大塚氏に新子誕生のため）\n②精算差額57万0362円（R7.4〜R8.1月分52万6210円＋R8.2・3月分4万4152円）を3月末支払\n③今後は遡って精算しない\n④次回協議は令和11年2月頃（陽翔くん満9歳時）"],
        ["令和8年3月29日", "鈴木七海様の事件フォルダ作成", "公正証書写真・服部弁護士書面等の資料を保存"],
        ["令和8年4月22日", "初回法律相談（対面・沖田弁護士オンライン参加）", "相談内容：\n①養育費見直し時期を2〜3年毎に変更\n②大学進学条項の明確化\n③その他不利にならない箇所の修正\n方針：減額交渉への対応と公正証書改訂を一体で進める"],
        ["令和8年4月23日", "委任契約書締結（Adobe Sign電子署名）", "事件名：養育費請求の交渉（相手方：大塚眞司）\n着手金20万円（税別）・報酬金20万円（税別）\n鈴木様18:41署名→沖田弁護士20:41署名"],
        ["令和8年4月27日", "鈴木七海様からLINEで追加要望", "①2〜3年毎の見直し・進学条項の明確化\n②扶養家族増による減額制限の可否\n③係争中の支払継続・未払い分の遡及請求"],
        ["令和8年4月28日", "受任通知及び御連絡（内容証明）起案", "米谷起案→沖田加筆→米谷再検討→最終版の過程で複数バージョン作成\n内容証明フォルダに未提出版（v1〜v4）、07_ワードファイルに電子内容証明案を保存"],
        ["令和8年5月1日", "電子内容証明第1通発送（服部咲弁護士宛）", "受付通番G02191879・郵便料金2,677円\n内容：受任通知＋減額拒否＋公正証書解釈＋改訂提案（7項目）＋回答期限5/12"],
        ["令和8年5月7日", "電子内容証明第2通発送（服部咲弁護士宛）", "受付通番G02195487・郵便料金2,677円\n最終送付版（260507電子内容証明-送付.docx）"],
        ["令和8年5月9日", "第1通が服部咲弁護士に配達（配達証明受領）", "追跡番号132-78-68759-0"],
        ["令和8年5月12日", "書面記載の回答期限（書面到達後2週間以内）", ""],
        ["令和8年5月15日", "第2通が寺井隼人弁護士宛に配達（配達証明受領）", "追跡番号132-78-94617-0\n※寺井隼人法律事務所は山岡総合法律事務所の沖田弁護士の関係先か"],
    ]

    # 3. データ書き込み
    body = {"values": timeline_data}
    sheets_service.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range="時系列!A1",
        valueInputOption="USER_ENTERED",
        body=body,
    ).execute()
    print(f"時系列データ書き込み完了（{len(timeline_data)-1}件）")

    # 4. 書式設定
    format_requests = [
        {
            "repeatCell": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": 0,
                    "endRowIndex": 1,
                },
                "cell": {
                    "userEnteredFormat": {
                        "textFormat": {"bold": True},
                        "backgroundColor": {"red": 0.9, "green": 0.9, "blue": 0.9},
                    }
                },
                "fields": "userEnteredFormat(textFormat,backgroundColor)",
            }
        },
        {
            "updateDimensionProperties": {
                "range": {
                    "sheetId": sheet_id,
                    "dimension": "COLUMNS",
                    "startIndex": 0,
                    "endIndex": 1,
                },
                "properties": {"pixelSize": 200},
                "fields": "pixelSize",
            }
        },
        {
            "updateDimensionProperties": {
                "range": {
                    "sheetId": sheet_id,
                    "dimension": "COLUMNS",
                    "startIndex": 1,
                    "endIndex": 2,
                },
                "properties": {"pixelSize": 500},
                "fields": "pixelSize",
            }
        },
        {
            "updateDimensionProperties": {
                "range": {
                    "sheetId": sheet_id,
                    "dimension": "COLUMNS",
                    "startIndex": 2,
                    "endIndex": 3,
                },
                "properties": {"pixelSize": 600},
                "fields": "pixelSize",
            }
        },
        {
            "repeatCell": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": 0,
                    "endRowIndex": len(timeline_data),
                    "startColumnIndex": 0,
                    "endColumnIndex": 3,
                },
                "cell": {
                    "userEnteredFormat": {
                        "wrapStrategy": "WRAP",
                        "verticalAlignment": "TOP",
                    }
                },
                "fields": "userEnteredFormat(wrapStrategy,verticalAlignment)",
            }
        },
        {
            "updateSheetProperties": {
                "properties": {
                    "sheetId": sheet_id,
                    "gridProperties": {"frozenRowCount": 1},
                },
                "fields": "gridProperties.frozenRowCount",
            }
        },
    ]

    sheets_service.spreadsheets().batchUpdate(
        spreadsheetId=SPREADSHEET_ID,
        body={"requests": format_requests},
    ).execute()
    print("書式設定完了")

    # 5. 不要なコピーを削除
    drive_service.files().delete(fileId=COPY_ID).execute()
    print(f"不要なコピー（{COPY_ID}）を削除しました")

    print(
        f"\n完了: https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/edit"
    )


if __name__ == "__main__":
    main()
