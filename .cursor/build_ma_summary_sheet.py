"""
ま_馬さん 案件整理票スプレッドシート作成スクリプト
- 人物整理／時系列／要求事項 の3シート
- 参照ファイルは Google Drive 直リンク
"""

import os
import sys
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]
CREDENTIALS_PATH = os.path.expanduser("~/law-secretary/secrets/oauth_credentials.json")
TOKEN_PATH = os.path.expanduser("~/law-secretary/secrets/token.json")

MA_FOLDER_PATH_PARTS = ["共有用", "01_事件記録", "ま_馬さん"]

SPREADSHEET_TITLE = "260511_案件整理票（馬様）"


def get_creds():
    creds = None
    if os.path.exists(TOKEN_PATH):
        creds = Credentials.from_authorized_user_file(TOKEN_PATH, SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_PATH, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(TOKEN_PATH, "w") as f:
            f.write(creds.to_json())
    return creds


def find_folder_id(drive, path_parts):
    """マイドライブ配下のフォルダパスをたどってIDを返す。"""
    parent = "root"
    for part in path_parts:
        q = (
            f"name = '{part}' and '{parent}' in parents "
            f"and mimeType = 'application/vnd.google-apps.folder' "
            f"and trashed = false"
        )
        res = drive.files().list(q=q, fields="files(id,name)", pageSize=10).execute()
        files = res.get("files", [])
        if not files:
            raise RuntimeError(f"フォルダが見つかりません: {part} (parent={parent})")
        parent = files[0]["id"]
    return parent


def list_files_recursive(drive, folder_id):
    """フォルダ配下のファイルを再帰的に取得し、名前→IDのdictを返す。"""
    result = {}
    stack = [folder_id]
    while stack:
        fid = stack.pop()
        page_token = None
        while True:
            res = drive.files().list(
                q=f"'{fid}' in parents and trashed = false",
                fields="nextPageToken, files(id,name,mimeType,parents)",
                pageSize=200,
                pageToken=page_token,
            ).execute()
            for f in res.get("files", []):
                if f["mimeType"] == "application/vnd.google-apps.folder":
                    stack.append(f["id"])
                # 同名重複は最初に見つかったものを採用
                result.setdefault(f["name"], f)
            page_token = res.get("nextPageToken")
            if not page_token:
                break
    return result


def file_link(fid):
    return f"https://drive.google.com/file/d/{fid}/view"


def folder_link(fid):
    return f"https://drive.google.com/drive/folders/{fid}"


def hyperlink_formula(url, label):
    safe_label = (label or url).replace('"', '""')
    return f'=HYPERLINK("{url}","{safe_label}")'


def create_spreadsheet(drive, sheets, parent_folder_id, title):
    file_metadata = {
        "name": title,
        "mimeType": "application/vnd.google-apps.spreadsheet",
        "parents": [parent_folder_id],
    }
    f = drive.files().create(body=file_metadata, fields="id,webViewLink").execute()
    return f["id"], f["webViewLink"]


def setup_sheets(sheets, sid):
    # デフォルトのシートを取得し、3シート構成にする
    meta = sheets.spreadsheets().get(spreadsheetId=sid).execute()
    existing = {s["properties"]["title"]: s["properties"]["sheetId"]
                for s in meta["sheets"]}
    requests = []
    target_sheets = ["人物整理", "時系列", "要求事項"]
    # 既存の最初のシート名を「人物整理」に変更
    first_id = list(existing.values())[0]
    requests.append({
        "updateSheetProperties": {
            "properties": {"sheetId": first_id, "title": "人物整理"},
            "fields": "title",
        }
    })
    # 残り2シート追加
    for t in target_sheets[1:]:
        requests.append({"addSheet": {"properties": {"title": t}}})
    sheets.spreadsheets().batchUpdate(
        spreadsheetId=sid, body={"requests": requests}
    ).execute()
    # 改めて全シートIDを取得
    meta = sheets.spreadsheets().get(spreadsheetId=sid).execute()
    return {s["properties"]["title"]: s["properties"]["sheetId"]
            for s in meta["sheets"]}


def write_values(sheets, sid, sheet_name, rows):
    sheets.spreadsheets().values().update(
        spreadsheetId=sid,
        range=f"'{sheet_name}'!A1",
        valueInputOption="USER_ENTERED",
        body={"values": rows},
    ).execute()


def apply_formatting(sheets, sid, sheet_ids):
    """ヘッダー行太字＋背景色、列幅、行高さ調整。"""
    requests = []
    for title, tab_id in sheet_ids.items():
        # ヘッダー行（1行目）の書式
        requests.append({
            "repeatCell": {
                "range": {
                    "sheetId": tab_id,
                    "startRowIndex": 0,
                    "endRowIndex": 1,
                },
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": {"red": 0.85, "green": 0.92, "blue": 0.98},
                        "textFormat": {"bold": True},
                        "verticalAlignment": "MIDDLE",
                        "wrapStrategy": "WRAP",
                    }
                },
                "fields": "userEnteredFormat(backgroundColor,textFormat,verticalAlignment,wrapStrategy)",
            }
        })
        # 全セルに折返し
        requests.append({
            "repeatCell": {
                "range": {"sheetId": tab_id, "startRowIndex": 1},
                "cell": {
                    "userEnteredFormat": {
                        "wrapStrategy": "WRAP",
                        "verticalAlignment": "TOP",
                    }
                },
                "fields": "userEnteredFormat(wrapStrategy,verticalAlignment)",
            }
        })
        # 1行目を固定
        requests.append({
            "updateSheetProperties": {
                "properties": {"sheetId": tab_id, "gridProperties": {"frozenRowCount": 1}},
                "fields": "gridProperties.frozenRowCount",
            }
        })
    # 列幅
    person_id = sheet_ids["人物整理"]
    requests.append({
        "updateDimensionProperties": {
            "range": {"sheetId": person_id, "dimension": "COLUMNS", "startIndex": 0, "endIndex": 1},
            "properties": {"pixelSize": 90},
            "fields": "pixelSize",
        }
    })
    requests.append({
        "updateDimensionProperties": {
            "range": {"sheetId": person_id, "dimension": "COLUMNS", "startIndex": 1, "endIndex": 2},
            "properties": {"pixelSize": 180},
            "fields": "pixelSize",
        }
    })
    requests.append({
        "updateDimensionProperties": {
            "range": {"sheetId": person_id, "dimension": "COLUMNS", "startIndex": 2, "endIndex": 3},
            "properties": {"pixelSize": 240},
            "fields": "pixelSize",
        }
    })
    requests.append({
        "updateDimensionProperties": {
            "range": {"sheetId": person_id, "dimension": "COLUMNS", "startIndex": 3, "endIndex": 4},
            "properties": {"pixelSize": 320},
            "fields": "pixelSize",
        }
    })
    timeline_id = sheet_ids["時系列"]
    requests.append({
        "updateDimensionProperties": {
            "range": {"sheetId": timeline_id, "dimension": "COLUMNS", "startIndex": 0, "endIndex": 1},
            "properties": {"pixelSize": 130},
            "fields": "pixelSize",
        }
    })
    requests.append({
        "updateDimensionProperties": {
            "range": {"sheetId": timeline_id, "dimension": "COLUMNS", "startIndex": 1, "endIndex": 2},
            "properties": {"pixelSize": 560},
            "fields": "pixelSize",
        }
    })
    requests.append({
        "updateDimensionProperties": {
            "range": {"sheetId": timeline_id, "dimension": "COLUMNS", "startIndex": 2, "endIndex": 3},
            "properties": {"pixelSize": 320},
            "fields": "pixelSize",
        }
    })
    req_id = sheet_ids["要求事項"]
    requests.append({
        "updateDimensionProperties": {
            "range": {"sheetId": req_id, "dimension": "COLUMNS", "startIndex": 0, "endIndex": 1},
            "properties": {"pixelSize": 140},
            "fields": "pixelSize",
        }
    })
    requests.append({
        "updateDimensionProperties": {
            "range": {"sheetId": req_id, "dimension": "COLUMNS", "startIndex": 1, "endIndex": 2},
            "properties": {"pixelSize": 520},
            "fields": "pixelSize",
        }
    })
    requests.append({
        "updateDimensionProperties": {
            "range": {"sheetId": req_id, "dimension": "COLUMNS", "startIndex": 2, "endIndex": 3},
            "properties": {"pixelSize": 200},
            "fields": "pixelSize",
        }
    })
    sheets.spreadsheets().batchUpdate(
        spreadsheetId=sid, body={"requests": requests}
    ).execute()


def main():
    creds = get_creds()
    drive = build("drive", "v3", credentials=creds)
    sheets = build("sheets", "v4", credentials=creds)

    ma_folder_id = find_folder_id(drive, MA_FOLDER_PATH_PARTS)
    print(f"馬さんフォルダID: {ma_folder_id}", file=sys.stderr)

    files = list_files_recursive(drive, ma_folder_id)
    print(f"ファイル数: {len(files)}", file=sys.stderr)

    def link(name):
        f = files.get(name)
        if not f:
            return ""
        if f["mimeType"] == "application/vnd.google-apps.folder":
            return hyperlink_formula(folder_link(f["id"]), name)
        return hyperlink_formula(file_link(f["id"]), name)

    # 人物整理シート
    person_rows = [
        ["区分", "氏名", "立場・属性", "備考"],
        ["依頼者側", "馬強（Martin Ma）", "依頼者・保護者／会社役員", "受任者本人。連絡先: martin.ma@letour.co.jp"],
        ["依頼者側", "Angelina Ma", "馬様の次女／KIST G9 在籍", "被害生徒。事案当日 教室を離れ体育館で襟元を掴まれる"],
        ["依頼者側", "George", "馬様の息子／Angelinaの弟", "事案当日 体育館で試合中"],
        ["依頼者側", "George の看護者", "家庭関係者（学校職員ではない）", "4/17以降 校内立入を永久禁止される旨通知。回復が要求事項"],
        ["相手方（学校）", "K. International School Tokyo (KIST)", "中高一貫国際学校", "事案発生校"],
        ["相手方", "Mrs. Komaki", "学校理事会議長（Board Chair）", "Angelinaの襟元を後方から掴み移動させた当事者"],
        ["相手方", "Mrs. Komakiの息子", "副理事長", "事案当日 Angelinaを教室に連れ戻す"],
        ["相手方", "Mr. Kei", "教員", "現場の走路脇に位置。Mrs. Komakiの行為を「心配から出たもの」と認識した旨が陳述に登場"],
        ["相手方", "Mr. Archer", "教員", "4/14面談で校内証言を提示。身体接触の存在を認める"],
        ["相手方", "Ms. Naito", "教員", "現場で走路横断を制止し手振りで合図"],
        ["相手方", "Mark Cowe", "学校管理層", "4/13事案当日の学校→保護者メール送信者"],
        ["相手方", "Karen Donald Godfrey", "スチューデントケアコーディネーター", "学校対応窓口の一人"],
        ["相手方", "スクールカウンセラー（氏名未特定）", "学校カウンセラー", "過去1年継続的に情緒問題を認識・心理相談を提案"],
        ["相手方代理人", "寺井隼人", "学校法務顧問（理事兼顧問弁護士）", "4/20以降 学校側の対応窓口（criminal offense/act等の表現使用）"],
        ["相手方関連", "元警視庁関係者", "学校が関与させた人物", "4/20以降関与"],
    ]
    person_values = [[c for c in row] for row in person_rows]

    # 時系列シート
    timeline_rows = [
        ["年月日", "出来事", "参照ファイル"],
        ["〜2026/4/12", "過去1年間、Angelinaに継続的な情緒的困難。スクールカウンセラーに専門支援を継続要請（学校も認識）",
         link("学校メール_KarenDonaldGodfreyスチューデントケアコーディネーター.pdf")],
        ["2026/4/13 09:18頃", "1時間目算数の授業中、Angelinaが弟Georgeの試合観覧のため約4分間 教室を離脱",
         link("第１事案の概要.docx") or link("Angelina書面陳述（4月13日付）.pdf")],
        ["2026/4/13", "体育館で Mrs. Komakiが後方から襟元を掴み移動させる。Mr. Keiが走路脇、Ms. Naitoが走路横断を制止",
         link("Angelina書面陳述（4月13日付）.pdf")],
        ["2026/4/13", "副理事長（Mrs. Komakiの息子）がAngelinaを教室に連れ戻す",
         link("第１事案の概要.docx")],
        ["2026/4/13", "Angelinaが校長に呼び出され、当日付の書面陳述を提出",
         link("Angelina書面陳述（4月13日付）.pdf")],
        ["2026/4/13", "学校が保護者面談を求めるとともに、協議完了まで登校認めない旨明確化（事実上の登校制限開始）",
         link("学校メール_MarkCowe（4月13日事案当日）.pdf")],
        ["2026/4/14", "学校と保護者の初回面談（保護者側で録音）。Mr. Archerが身体接触を認める証言提示。学校は「誘導」と説明。発言制限・在学資格関連付け・メール停止",
         link("面談録音主要抜粋（4月14日学校面談）.pdf")],
        ["2026/4/14〜15", "学校とのメールやり取り。学校側は「発生は不可能」「支持されない」と否定的結論を反復",
         link("学校メール_KarenDonaldGodfreyスチューデントケアコーディネーター_2.pdf")],
        ["2026/4/15", "American Clinic Tokyoで緊急精神科受診・薬物調整。医療費領収書",
         link("医療費領収書（AmericanClinicTokyo_4月16日精神科）.pdf")],
        ["2026/4/16", "Georgeの同級生の保護者から動画取得（約4分11秒・約4秒の2本）。動画存在は伏せて学校に再度説明要求",
         link("WeChat記録（証拠動画取得経緯）.pdf")],
        ["2026/4/17", "学校が4/20以降のGeorgeの看護者の校内立入を永久禁止する旨通知",
         link("学校メール_KarenDonaldGodfreyスチューデントケアコーディネーター_2.pdf")],
        ["2026/4/19", "当方が動画の存在を初回開示し、学校に見解再考を要求。法律顧問に対応窓口切替",
         link("学校法務顧問メール_寺井隼人（4月20日）.pdf")],
        ["2026/4/20", "学校法務顧問（寺井隼人）の関与、「criminal offense / criminal act」表現使用、元警視庁関係者関与、動画提出要求、登校制限維持",
         link("学校法務顧問メール_寺井隼人（4月20日）.pdf")],
        ["2026/4/20以降", "Angelinaに明確な恐怖反応。早期再受診希望も予約都合で4/30維持",
         link("事案整理書面（依頼者作成）.pdf")],
        ["2026/4/30", "米谷弁護士との打合せ。新年度授業料支払期限が5/15と判明。AmericanClinicTokyo再受診（医療費領収書）",
         link("第１事案の概要.docx")],
        ["2026/5/1", "心理カウンセリング費用請求書",
         link("証拠9：医療費に関する領収書")],
        ["2026/5/7", "法律相談（相談カード）／委任契約書（馬様）締結",
         link("260507_委任契約書（馬様）.pdf")],
        ["2026/5/9", "委任契約書（馬様）第2版作成（次女表記修正・成功報酬構成確定）",
         link("260509_委任契約書（馬様）.pdf")],
    ]

    # 要求事項シート
    req_rows = [
        ["カテゴリ", "項目", "金額・条件"],
        ["A. 達成目標", "1. 学校側の対応および処理過程に不適切な点があったことの確認", ""],
        ["A. 達成目標", "2. Angelinaの名誉および人格的評価の回復", ""],
        ["A. 達成目標", "3. Georgeの看護者による正常な校内立入り権限の回復", ""],
        ["A. 達成目標", "4. 本件によって生じた経済的損害・精神的損害に対する適切な補償", ""],
        ["A. 達成目標", "5. 復学問題に関する検討", ""],
        ["B-1. 実費損害", "CGA G9 Term 2 学費（代替教育費）", "1,260,000円"],
        ["B-1. 実費損害", "CGA G10 オンラインスクール預託金（10% Deposit）", "244,800円"],
        ["B-1. 実費損害", "KIST学費（4/14〜正常復学日まで按分／年間2,897,000円）", "按分相当額"],
        ["B-1. 実費損害", "通学定期券損失（4/14〜復学日まで／6か月定期35,980円）", "按分相当額"],
        ["B-1. 実費損害", "4/16 精神科診療費・薬代", "21,850円"],
        ["B-1. 実費損害", "4/30 精神科診療費・薬代", "57,050円"],
        ["B-1. 実費損害", "5/1 心理カウンセリング費用", "17,000円"],
        ["B-1. 実費損害", "5/8 心理カウンセリング費用", "17,000円"],
        ["B-1. 実費損害", "精神科往復交通費（2〜3週に1回・保護者同行）", "840円/回"],
        ["B-1. 実費損害", "カウンセリング往復交通費（週1回・保護者同行）", "840円/回"],
        ["B-1. 実費損害", "弁護士費用その他の法的対応費用", "別途"],
        ["B-2. 精神的損害", "Angelina本人の精神的損害（暫定請求額 約700万円）", "約7,000,000円（暫定）"],
        ["B-2. 精神的損害", "約1か月の登校不能による不安・心理的負担", "上記に含む"],
        ["B-2. 精神的損害", "学業中断・GPA・Transcript・進学経路への影響", "上記に含む"],
        ["B-2. 精神的損害", "名誉権・人格権・校内評価への侵害", "上記に含む"],
        ["B-2. 精神的損害", "教育機会損失・将来の進学上の不利益リスク", "上記に含む"],
        ["C. 成功報酬構成（5/9確定）", "学校が不適切認容＋Angelinaへの書面謝罪・名誉回復措置", "15万円"],
        ["C. 成功報酬構成（5/9確定）", "Georgeの看護者の校内立入り権限が回復", "15万円"],
        ["C. 成功報酬構成（5/9確定）", "合理的・公平な復学方案で双方合意し正常な就学再開", "15万円"],
        ["C. 成功報酬構成（5/9確定）", "経済的損害賠償部分", "経済的利益の15%"],
        ["C. 成功報酬構成（5/9確定）", "着手金（税込）", "33万円（5/11入金済）"],
        ["D. 重要指示事項", "当方の明確な事前同意なく、いかなる証拠資料（特に動画）も学校・第三者に提示・提供・開示・転送しないこと", "—"],
        ["D. 重要指示事項", "学校宛通知書は、相手方代理人送付前に必ずドラフトを依頼者に事前共有・確認すること", "—"],
        ["E. 早急対応（5/11指示）", "学校宛通知書を作成し、5/14までに学校側代理人・関連者へ到達するよう送付", "期限：5/14到達"],
        ["E. 早急対応（5/11指示）", "通知書趣旨①：学校は十分な法的根拠なくAngelinaに約1か月の事実上の停学を継続している旨", ""],
        ["E. 早急対応（5/11指示）", "通知書趣旨②：5/15期限のG10学費支払が現時点で困難である旨", ""],
        ["E. 早急対応（5/11指示）", "通知書趣旨③：ただし「自主退学」の意思表示ではない旨", ""],
        ["E. 早急対応（5/11指示）", "通知書趣旨④：本件適法解決後、学業継続の権利を留保する旨", ""],
        ["E. 早急対応（5/11指示）", "通知書趣旨⑤：学校側の一方的要求・処分・退学扱いは承諾しない旨", ""],
    ]

    sid, url = create_spreadsheet(drive, sheets, ma_folder_id, SPREADSHEET_TITLE)
    print(f"スプレッドシート作成: {url}", file=sys.stderr)

    sheet_ids = setup_sheets(sheets, sid)
    write_values(sheets, sid, "人物整理", person_values)
    write_values(sheets, sid, "時系列", timeline_rows)
    write_values(sheets, sid, "要求事項", req_rows)
    apply_formatting(sheets, sid, sheet_ids)

    print(url)


if __name__ == "__main__":
    main()
