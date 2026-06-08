"""引越しやることリスト スプレッドシート作成"""

import json
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

TOKEN_PATH = "secrets/token.json"
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

with open(TOKEN_PATH) as f:
    token_data = json.load(f)
creds = Credentials.from_authorized_user_info(token_data, SCOPES)

sheets_svc = build("sheets", "v4", credentials=creds)
drive_svc = build("drive", "v3", credentials=creds)

# ---------- データ ----------
rows = [
    # [カテゴリ, #, やること, 担当, 期限, 備考, 完了]
    ["1. 契約・手配", "1-1", "賃貸解約予告", "尚起", "済み", "済み", "TRUE"],
    ["1. 契約・手配", "1-2", "引越し業者の見積もり・手配（2〜3社相見積もり）", "尚起", "2026-06-07", "7月中旬は混みやすい。早めに", ""],
    ["1. 契約・手配", "1-3", "新居の契約手続き完了・初期費用支払い", "尚起", "2026-06-07", "鍵の受取日も確認", ""],
    ["1. 契約・手配", "1-4", "新居の火災保険加入", "尚起", "2026-06-07", "賃貸契約時セットの場合もあるが確認", ""],
    ["1. 契約・手配", "1-5", "新居のインターネット回線申込み", "尚起", "2026-06-07", "開通工事に2〜4週間。早めに", ""],
    ["1. 契約・手配", "1-6", "梱包資材の準備", "尚起", "2026-06-14", "業者提供分を確認、不足分は購入", ""],
    ["2. 新居準備", "2-1", "新居の採寸・家具配置検討", "両方", "2026-06-14", "カーテン・家電サイズ。エアコン設置可否", ""],
    ["2. 新居準備", "2-2", "新居の入居前清掃・害虫駆除", "両方", "2026-07-06", "バルサン等。赤ちゃんがいるので搬入前に", ""],
    ["3. 不用品処分", "3-1", "不用品の仕分け・処分計画", "両方", "2026-06-14", "粗大ごみは予約制（1〜2週間待ち）", ""],
    ["3. 不用品処分", "3-2", "粗大ごみの予約・回収", "両方", "2026-07-06", "引越し前に完了させる", ""],
    ["4. ライフライン", "4-1", "電気の停止/開始手続き", "尚起", "2026-06-30", "Webで可。旧居停止7/14、新居開始7/14", ""],
    ["4. ライフライン", "4-2", "ガスの停止/開始手続き", "尚起", "2026-06-30", "新居は開栓立会い必要。早めに予約", ""],
    ["4. ライフライン", "4-3", "水道の停止/開始手続き", "尚起", "2026-06-30", "Webまたは電話", ""],
    ["4. ライフライン", "4-4", "旧居のインターネット回線の解約手続き", "尚起", "2026-06-30", "違約金・撤去工事の有無を確認", ""],
    ["5. 届出（引越し前）", "5-1", "転出届の提出（豊島区）", "尚起", "2026-07-07", "マイナポータルでオンライン可。14日前から", ""],
    ["5. 届出（引越し前）", "5-2", "郵便局に転居届（e転居 or 窓口）", "尚起", "2026-06-30", "反映まで数日。早めに", ""],
    ["5. 届出（引越し前）", "5-3", "NHKの住所変更", "尚起", "2026-07-07", "Webで手続き可能", ""],
    ["5. 届出（引越し前）", "5-4", "各種定期配達サービスの住所変更・解約", "両方", "2026-07-07", "生協・ウォーターサーバー等", ""],
    ["5. 届出（引越し前）", "5-5", "優花さんの勤務先に住所変更を連絡", "優花", "2026-07-07", "育休中でも届出必要（通勤手当・社保）", ""],
    ["6. 荷造り", "6-1", "荷造り開始（季節外れの服・本等）", "両方", "2026-06-28", "赤ちゃん用品・日用品は最後に", ""],
    ["6. 荷造り", "6-2", "貴重品・重要書類をまとめる", "両方", "2026-07-13", "通帳・印鑑・保険証・母子手帳・マイナカード", ""],
    ["7. 引越し前日", "7-1", "引越し業者との前日最終確認", "尚起", "2026-07-13", "時間・住所・エレベーター・駐車スペース", ""],
    ["7. 引越し前日", "7-2", "冷蔵庫の水抜き・霜取り", "尚起", "2026-07-13", "前日夜に電源OFF", ""],
    ["7. 引越し前日", "7-3", "洗濯機の水抜き", "尚起", "2026-07-13", "前日に実施", ""],
    ["7. 引越し前日", "7-4", "赤ちゃんの当日の過ごし方を確保", "優花", "2026-07-13", "一時預かり or 実家 or 外出先を決めておく", ""],
    ["8. 引越し当日", "8-1", "旧居の状態を写真撮影", "尚起", "2026-07-14", "原状回復トラブル防止。壁・床・水回り", ""],
    ["8. 引越し当日", "8-2", "旧居の近隣への挨拶", "両方", "2026-07-14", "搬出時の騒音のお詫びも兼ねて", ""],
    ["8. 引越し当日", "8-3", "旧居の最終清掃", "両方", "2026-07-14", "", ""],
    ["8. 引越し当日", "8-4", "旧居の退去立会い（鍵返却）", "尚起", "2026-07-14", "管理会社と日時調整。当日 or 後日", ""],
    ["8. 引越し当日", "8-5", "新居の傷・汚れチェック（写真撮影）", "尚起", "2026-07-14", "搬入前に入居時の状態を記録", ""],
    ["8. 引越し当日", "8-6", "新居でガス開栓立会い", "尚起", "2026-07-14", "時間指定しておく", ""],
    ["8. 引越し当日", "8-7", "新居の電気・水道の開通確認", "尚起", "2026-07-14", "", ""],
    ["8. 引越し当日", "8-8", "搬入後の荷物・家具の破損確認", "両方", "2026-07-14", "業者がいる間に確認。破損は当日申告", ""],
    ["8. 引越し当日", "8-9", "引越し料金の精算", "尚起", "2026-07-14", "現金払いの場合は事前に用意", ""],
    ["8. 引越し当日", "8-10", "新居で近隣への挨拶", "両方", "2026-07-14", "赤ちゃんがいる旨を一言添える", ""],
    ["9. 役所手続き", "9-1", "転入届の提出（板橋区）", "尚起", "2026-07-28", "窓口のみ。転出証明書 or マイナカード持参", ""],
    ["9. 役所手続き", "9-2", "マイナンバーカードの住所変更", "両方", "2026-07-28", "転入届と同時。各自のカード＋暗証番号", ""],
    ["9. 役所手続き", "9-3", "国民健康保険の住所変更", "尚起", "2026-07-28", "世帯主として。家族分も", ""],
    ["9. 役所手続き", "9-4", "乳幼児医療証（マル乳）の届出", "優花", "2026-07-28", "板橋区で新規交付申請。春来くん分", ""],
    ["9. 役所手続き", "9-5", "児童手当の届出", "優花", "2026-07-29", "豊島区で消滅届→板橋区で認定請求。15日以内", ""],
    ["9. 役所手続き", "9-6", "印鑑登録（必要な場合）", "尚起", "2026-07-28", "豊島区の登録は転出で自動失効", ""],
    ["10. 弁護士業務", "10-1", "東京弁護士会への自宅住所変更届", "尚起", "2026-07-28", "事務所住所は不変だが自宅住所を確認", ""],
    ["10. 弁護士業務", "10-2", "弁護士法人の代表社員住所変更（法務局）", "尚起", "2026-07-28", "登記事項に含まれる場合。2週間以内", ""],
    ["10. 弁護士業務", "10-3", "税務署・都税事務所への異動届", "尚起", "2026-07-28", "法人代表者住所変更として必要か確認", ""],
    ["11. 住所変更（各種）", "11-1", "銀行口座の住所変更", "尚起", "2026-07-28", "GMOあおぞら等。ネットバンキングで可", ""],
    ["11. 住所変更（各種）", "11-2", "クレジットカードの住所変更", "両方", "2026-07-28", "各自のカード分", ""],
    ["11. 住所変更（各種）", "11-3", "生命保険・損害保険の住所変更", "尚起", "2026-07-28", "", ""],
    ["11. 住所変更（各種）", "11-4", "携帯電話の住所変更", "両方", "2026-07-28", "各自で", ""],
    ["11. 住所変更（各種）", "11-5", "運転免許証の住所変更", "両方", "2026-07-31", "最寄りの警察署。新住所の住民票持参", ""],
    ["11. 住所変更（各種）", "11-6", "Amazon・EC・サブスクの配送先変更", "両方", "2026-07-28", "各自で", ""],
    ["11. 住所変更（各種）", "11-7", "ふるさと納税ワンストップ特例の届出", "尚起", "2026-07-31", "申請済み自治体へ変更届出書を送付", ""],
    ["11. 住所変更（各種）", "11-8", "敷金返還の確認・精算", "尚起", "2026-09-14", "退去後1〜2ヶ月で返還。明細確認", ""],
    ["12. 赤ちゃん関連", "12-1", "新居周辺の小児科・病院を探す", "優花", "2026-06-30", "予防接種スケジュールとの兼ね合い", ""],
    ["12. 赤ちゃん関連", "12-2", "板橋区の保育園情報を収集", "優花", "2026-06-30", "育休明けに向けて保活事情を把握", ""],
    ["12. 赤ちゃん関連", "12-3", "予防接種の引き継ぎ準備", "優花", "2026-07-28", "母子手帳＋接種済み証明を新しい小児科へ", ""],
    ["12. 赤ちゃん関連", "12-4", "板橋区の予防接種予診票の取得", "優花", "2026-07-28", "転入届時に窓口で受け取れる場合あり", ""],
    ["12. 赤ちゃん関連", "12-5", "板橋区の子育て支援サービスを確認", "優花", "2026-07-31", "一時預かり・子育て広場・ファミサポ等", ""],
    ["12. 赤ちゃん関連", "12-6", "保育園の見学予約", "優花", "2026-08-31", "秋の申込みに向けて早めに", ""],
    ["13. 引越し後片付け", "13-1", "梱包資材（ダンボール等）の返却", "尚起", "2026-07-31", "業者に回収依頼。無料回収期限を確認", ""],
]

# 期限でソート（「済み」は先頭に）
def sort_key(row):
    d = row[4]
    if d == "済み":
        return "0000-00-00"
    return d

rows.sort(key=sort_key)

# ---------- スプレッドシート作成 ----------
body = {
    "properties": {"title": "引越しやることリスト（7/14）"},
    "sheets": [
        {
            "properties": {
                "title": "チェックリスト",
                "gridProperties": {"rowCount": len(rows) + 2, "columnCount": 7},
            }
        }
    ],
}
ss = sheets_svc.spreadsheets().create(body=body).execute()
ss_id = ss["spreadsheetId"]
sheet_id = ss["sheets"][0]["properties"]["sheetId"]

# ---------- ヘッダー + データ書き込み ----------
header = ["カテゴリ", "#", "やること", "担当", "期限", "備考", "完了"]
values = [header] + rows

sheets_svc.spreadsheets().values().update(
    spreadsheetId=ss_id,
    range="チェックリスト!A1",
    valueInputOption="USER_ENTERED",
    body={"values": values},
).execute()

# ---------- 書式設定 ----------
requests = []

# 1) ヘッダー行: 太字・背景色・固定
requests.append({
    "repeatCell": {
        "range": {"sheetId": sheet_id, "startRowIndex": 0, "endRowIndex": 1},
        "cell": {
            "userEnteredFormat": {
                "backgroundColor": {"red": 0.2, "green": 0.4, "blue": 0.65},
                "textFormat": {"bold": True, "foregroundColor": {"red": 1, "green": 1, "blue": 1}},
                "horizontalAlignment": "CENTER",
            }
        },
        "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)",
    }
})

requests.append({
    "updateSheetProperties": {
        "properties": {"sheetId": sheet_id, "gridProperties": {"frozenRowCount": 1}},
        "fields": "gridProperties.frozenRowCount",
    }
})

# 2) 列幅
col_widths = [
    (0, 160),  # カテゴリ
    (1, 60),   # #
    (2, 380),  # やること
    (3, 70),   # 担当
    (4, 100),  # 期限
    (5, 350),  # 備考
    (6, 50),   # 完了
]
for col_idx, px in col_widths:
    requests.append({
        "updateDimensionProperties": {
            "range": {"sheetId": sheet_id, "dimension": "COLUMNS", "startIndex": col_idx, "endIndex": col_idx + 1},
            "properties": {"pixelSize": px},
            "fields": "pixelSize",
        }
    })

# 3) 担当列にプルダウン（データバリデーション）
requests.append({
    "setDataValidation": {
        "range": {"sheetId": sheet_id, "startRowIndex": 1, "endRowIndex": len(rows) + 1, "startColumnIndex": 3, "endColumnIndex": 4},
        "rule": {
            "condition": {
                "type": "ONE_OF_LIST",
                "values": [
                    {"userEnteredValue": "尚起"},
                    {"userEnteredValue": "優花"},
                    {"userEnteredValue": "両方"},
                ],
            },
            "showCustomUi": True,
            "strict": True,
        },
    }
})

# 4) 完了列にチェックボックス
requests.append({
    "setDataValidation": {
        "range": {"sheetId": sheet_id, "startRowIndex": 1, "endRowIndex": len(rows) + 1, "startColumnIndex": 6, "endColumnIndex": 7},
        "rule": {
            "condition": {"type": "BOOLEAN"},
            "showCustomUi": True,
        },
    }
})

# 5) 期限列を日付書式に
requests.append({
    "repeatCell": {
        "range": {"sheetId": sheet_id, "startRowIndex": 1, "endRowIndex": len(rows) + 1, "startColumnIndex": 4, "endColumnIndex": 5},
        "cell": {
            "userEnteredFormat": {
                "numberFormat": {"type": "DATE", "pattern": "M/d（ddd）"},
            }
        },
        "fields": "userEnteredFormat.numberFormat",
    }
})

# 6) 交互の背景色
requests.append({
    "addBanding": {
        "bandedRange": {
            "range": {"sheetId": sheet_id, "startRowIndex": 0, "endRowIndex": len(rows) + 1, "startColumnIndex": 0, "endColumnIndex": 7},
            "rowProperties": {
                "headerColor": {"red": 0.2, "green": 0.4, "blue": 0.65},
                "firstBandColor": {"red": 1, "green": 1, "blue": 1},
                "secondBandColor": {"red": 0.92, "green": 0.94, "blue": 0.97},
            },
        }
    }
})

# 7) 済み行（1行目データ）をグレーアウト＋取消線
requests.append({
    "repeatCell": {
        "range": {"sheetId": sheet_id, "startRowIndex": 1, "endRowIndex": 2, "startColumnIndex": 0, "endColumnIndex": 7},
        "cell": {
            "userEnteredFormat": {
                "textFormat": {
                    "strikethrough": True,
                    "foregroundColor": {"red": 0.6, "green": 0.6, "blue": 0.6},
                },
            }
        },
        "fields": "userEnteredFormat.textFormat(strikethrough,foregroundColor)",
    }
})

# 8) 担当列の条件付き書式（色分け）
# 尚起 = 青
requests.append({
    "addConditionalFormatRule": {
        "rule": {
            "ranges": [{"sheetId": sheet_id, "startRowIndex": 1, "endRowIndex": len(rows) + 1, "startColumnIndex": 3, "endColumnIndex": 4}],
            "booleanRule": {
                "condition": {"type": "TEXT_EQ", "values": [{"userEnteredValue": "尚起"}]},
                "format": {
                    "backgroundColor": {"red": 0.85, "green": 0.92, "blue": 1.0},
                    "textFormat": {"foregroundColor": {"red": 0.1, "green": 0.3, "blue": 0.6}},
                },
            },
        },
        "index": 0,
    }
})
# 優花 = ピンク
requests.append({
    "addConditionalFormatRule": {
        "rule": {
            "ranges": [{"sheetId": sheet_id, "startRowIndex": 1, "endRowIndex": len(rows) + 1, "startColumnIndex": 3, "endColumnIndex": 4}],
            "booleanRule": {
                "condition": {"type": "TEXT_EQ", "values": [{"userEnteredValue": "優花"}]},
                "format": {
                    "backgroundColor": {"red": 1.0, "green": 0.88, "blue": 0.9},
                    "textFormat": {"foregroundColor": {"red": 0.6, "green": 0.1, "blue": 0.2}},
                },
            },
        },
        "index": 1,
    }
})
# 両方 = 緑
requests.append({
    "addConditionalFormatRule": {
        "rule": {
            "ranges": [{"sheetId": sheet_id, "startRowIndex": 1, "endRowIndex": len(rows) + 1, "startColumnIndex": 3, "endColumnIndex": 4}],
            "booleanRule": {
                "condition": {"type": "TEXT_EQ", "values": [{"userEnteredValue": "両方"}]},
                "format": {
                    "backgroundColor": {"red": 0.85, "green": 0.95, "blue": 0.85},
                    "textFormat": {"foregroundColor": {"red": 0.1, "green": 0.45, "blue": 0.1}},
                },
            },
        },
        "index": 2,
    }
})

sheets_svc.spreadsheets().batchUpdate(
    spreadsheetId=ss_id, body={"requests": requests}
).execute()

url = f"https://docs.google.com/spreadsheets/d/{ss_id}"
print(f"Created: {url}")
