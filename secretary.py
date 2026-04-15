"""
弁護士法人Re-Start法律事務所 秘書エージェント
06_分類依頼フォルダのファイルを自動分類し、スプレッドシートに記録して佐藤信子へ通知する。
"""

import base64
import logging
import os
import re
import shutil
from datetime import datetime
from email.mime.text import MIMEText

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

# ── 認証設定 ──
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/gmail.send",
    "https://www.googleapis.com/auth/drive",
]
CREDENTIALS_PATH = os.path.expanduser("~/law-secretary/secrets/oauth_credentials.json")
TOKEN_PATH = os.path.expanduser("~/law-secretary/secrets/token.json")

# ── Googleドライブパス ──
BASE_PATH = (
    "/Users/kometaninaoki/Library/CloudStorage/"
    "GoogleDrive-n.kometani@re-startlaw.com/マイドライブ/"
)
WATCH_FOLDER = os.path.join(BASE_PATH, "共有用", "06_分類依頼")
CASE_RECORDS = os.path.join(BASE_PATH, "共有用", "01_事件記録")
ACCOUNTING = os.path.join(BASE_PATH, "共有用", "03_経理")
PRE_TRASH = os.path.join(WATCH_FOLDER, "pre_trash")
SCANSNAP_FOLDER = os.path.join(WATCH_FOLDER, "スキャナーから")

# ── スプレッドシート ──
SPREADSHEET_ID = "1-dZd7iC2-eXLCUOwHGNZCDV8U7y9ECL6wNBk5YC-czc"
SHEET_NAME = "保存ログ"

# ── 通知先 ──
NOTIFY_TO = ["n.sato@re-startlaw.com", "n.kometani@re-startlaw.com"]
NOTIFY_SUBJECT = "ファイル分類完了のご報告"
SENDER = "n.kometani@re-startlaw.com"

# ── バックアップ設定 ──
BACKUP_DEST = os.path.join(
    BASE_PATH, "共有用", "15_AI教育用"
)
BACKUP_TARGETS = [
    os.path.expanduser("~/law-secretary/CLAUDE.md"),
    os.path.expanduser("~/law-secretary/secretary.py"),
]
LOG_PATH = os.path.expanduser("~/law-secretary/secretary.log")

# ── ログ設定 ──
logging.basicConfig(
    filename=LOG_PATH,
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)

EMAIL_SIGNATURE = (
    "\n\n--\n"
    "〒170-6012 東京都豊島区東池袋３丁目１−１ サンシャイン60 12階\n"
    "弁護士法人Re-Start法律事務所\n"
    "弁護士 米谷尚起\n"
    "TEL : 03-6820-3815"
)

# ── 経理キーワード → 勘定科目 ──
ACCOUNTING_KEYWORDS = {
    "地代家賃": ["リージャス", "賃料", "家賃"],
    "通信費": ["Google", "通信", "インターネット"],
    "外注費": ["接見", "弁護士費用", "外注"],
    "会議費": ["スターバックス", "会議"],
    "接待交際費": ["接待", "交際"],
    "旅費交通費": ["交通", "タクシー", "電車", "小田急", "PASMO", "Suica"],
    "消耗品費": ["消耗品", "備品"],
    "広告宣伝費": ["広告", "宣伝"],
    "支払手数料": ["手数料"],
    "研修費": ["研修"],
    "新聞図書費": ["新聞", "図書", "書籍"],
    "諸会費": ["会費"],
}

# ── 刑事案件サブフォルダキーワード ──
CRIMINAL_SUBFOLDER_KEYWORDS = {
    "01_選任関係": ["選任", "委任"],
    "02_身体拘束関係": ["勾留", "身体拘束", "逮捕", "保釈"],
    "03_検察官提出書面": ["検察官提出", "起訴状"],
    "04_弁護人提出書面": ["準抗告", "弁護人", "意見書"],
    "05_検察官証拠": ["検察官証拠", "証拠開示"],
    "06_裁判所手続": ["裁判所", "期日"],
    "07_収集資料": ["収集資料", "資料"],
    "08_論文・文献": ["論文", "文献"],
    "10_メモ": ["メモ"],
    "11_弁護人請求証拠": ["弁護人請求証拠"],
    "12_精算？": ["精算"],
}

# ── 民事案件サブフォルダキーワード ──
CIVIL_SUBFOLDER_KEYWORDS = {
    "00主張": ["準備書面", "主張"],
    "01甲号証": ["甲号証"],
    "02乙号証": ["乙号証"],
    "03連絡文書": ["連絡", "通知"],
    "04事務": ["事務", "送付書"],
    "05期日報告書": ["期日報告"],
    "06資料": ["資料"],
    "委任関係": ["委任", "委任状", "受任"],
}

# ── 案件種別（刑事 / 民事）──
# 実際のフォルダ構造から判定する。01_選任関係があれば刑事、00主張があれば民事。
CRIMINAL_MARKER = "01_選任関係"
CIVIL_MARKER = "00主張"


# ════════════════════════════════════════════════════════════════
# バックアップ
# ════════════════════════════════════════════════════════════════

def backup_files():
    """起動時にCLAUDE.mdとsecretary.pyをGoogleドライブにバックアップする。"""
    today = datetime.now().strftime("%y%m%d")
    os.makedirs(BACKUP_DEST, exist_ok=True)

    for src_path in BACKUP_TARGETS:
        basename = os.path.basename(src_path)
        backup_name = f"{today}_{basename}"
        dest_path = os.path.join(BACKUP_DEST, backup_name)

        try:
            if os.path.exists(dest_path):
                logging.info(f"バックアップスキップ（既存）: {backup_name}")
                print(f"  バックアップスキップ: {backup_name}（既に存在）")
                continue

            if not os.path.exists(src_path):
                logging.warning(f"バックアップ元が見つかりません: {src_path}")
                print(f"  バックアップ失敗: {basename}（ファイルが見つかりません）")
                continue

            shutil.copy2(src_path, dest_path)
            logging.info(f"バックアップ成功: {backup_name} → {BACKUP_DEST}")
            print(f"  バックアップ完了: {backup_name}")
        except Exception as e:
            logging.error(f"バックアップ失敗: {backup_name} - {e}")
            print(f"  バックアップ失敗: {backup_name}（{e}）")


# ════════════════════════════════════════════════════════════════
# 認証
# ════════════════════════════════════════════════════════════════

def get_credentials():
    """OAuth2.0認証情報を取得する。"""
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


# ════════════════════════════════════════════════════════════════
# 依頼者フォルダの一覧を取得
# ════════════════════════════════════════════════════════════════

def load_client_folders():
    """01_事件記録内の依頼者フォルダ一覧を返す。
    Returns:
        list[dict]: [{"folder_name": "た　田村正宜", "client_name": "田村正宜",
                       "path": "/full/path", "case_type": "criminal"|"civil"}]
    """
    clients = []
    if not os.path.isdir(CASE_RECORDS):
        return clients
    for name in os.listdir(CASE_RECORDS):
        full = os.path.join(CASE_RECORDS, name)
        if not os.path.isdir(full):
            continue
        # 「読み仮名＋全角スペース＋氏名」から氏名部分を抽出
        parts = name.split("\u3000", 1)  # 全角スペースで分割
        client_name = parts[1] if len(parts) == 2 else name

        # 案件種別を判定
        if os.path.isdir(os.path.join(full, CRIMINAL_MARKER)):
            case_type = "criminal"
        elif os.path.isdir(os.path.join(full, CIVIL_MARKER)):
            case_type = "civil"
        else:
            case_type = "unknown"

        clients.append({
            "folder_name": name,
            "client_name": client_name,
            "path": full,
            "case_type": case_type,
        })
    return clients


# ════════════════════════════════════════════════════════════════
# 分類判断
# ════════════════════════════════════════════════════════════════

def classify_accounting(filename):
    """経理書類として分類できる場合、移動先パスを返す。"""
    keywords_in_name = ["領収書", "請求書", "精算"]
    is_accounting = any(kw in filename for kw in keywords_in_name)

    matched_category = None
    for category, keywords in ACCOUNTING_KEYWORDS.items():
        if any(kw in filename for kw in keywords):
            matched_category = category
            break

    if is_accounting and matched_category:
        return os.path.join(ACCOUNTING, matched_category)
    if matched_category and not is_accounting:
        # キーワードだけでは経理とは限らない → None
        return None
    if is_accounting and not matched_category:
        return os.path.join(ACCOUNTING, "未分類")
    return None


def classify_case_record(filename, clients):
    """事件記録として分類できる場合、移動先パスを返す。"""
    for client in clients:
        if client["client_name"] not in filename:
            continue
        # 依頼者が見つかった → サブフォルダを判定
        if client["case_type"] == "criminal":
            subfolder = _match_subfolder(filename, CRIMINAL_SUBFOLDER_KEYWORDS)
        elif client["case_type"] == "civil":
            subfolder = _match_subfolder(filename, CIVIL_SUBFOLDER_KEYWORDS)
        else:
            subfolder = None

        if subfolder:
            return os.path.join(client["path"], subfolder)
        # サブフォルダが特定できない → 依頼者フォルダ直下
        return client["path"]
    return None


def _match_subfolder(filename, keyword_map):
    """キーワードマップからサブフォルダ名を返す。"""
    for subfolder, keywords in keyword_map.items():
        if any(kw in filename for kw in keywords):
            return subfolder
    return None


def classify_file(filename, clients):
    """ファイル名から分類先のパスを返す。判断不能ならpre_trashを返す。

    Returns:
        tuple: (dest_path, description)
    """
    # 1. 経理書類チェック
    dest = classify_accounting(filename)
    if dest:
        rel = os.path.relpath(dest, os.path.join(BASE_PATH, "共有用"))
        return dest, rel

    # 2. 事件記録チェック
    dest = classify_case_record(filename, clients)
    if dest:
        rel = os.path.relpath(dest, os.path.join(BASE_PATH, "共有用"))
        return dest, rel

    # 3. 判断不能 → pre_trash
    return PRE_TRASH, "06_分類依頼/pre_trash"


# ════════════════════════════════════════════════════════════════
# ファイル移動
# ════════════════════════════════════════════════════════════════

def rename_file(filename):
    """ファイル命名規則に従いリネームする。
    すでに日付6桁_で始まる場合はそのまま。
    """
    if re.match(r"^\d{6}_", filename):
        return filename
    today = datetime.now().strftime("%y%m%d")
    return f"{today}_{filename}"


def move_file(src_path, dest_dir, new_name):
    """ファイルを移動する。移動先フォルダがなければ作成する。"""
    os.makedirs(dest_dir, exist_ok=True)
    dest_path = os.path.join(dest_dir, new_name)
    # 同名ファイルが存在する場合は連番を付ける
    if os.path.exists(dest_path):
        base, ext = os.path.splitext(new_name)
        i = 1
        while os.path.exists(dest_path):
            dest_path = os.path.join(dest_dir, f"{base}_{i}{ext}")
            i += 1
    shutil.move(src_path, dest_path)
    return dest_path


# ════════════════════════════════════════════════════════════════
# スプレッドシート記録
# ════════════════════════════════════════════════════════════════

def fetch_sheet_data(sheets_service):
    """保存ログシートの全データを取得する。
    Returns:
        list[list[str]]: 全行データ（ヘッダー含む）
    """
    result = sheets_service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{SHEET_NAME}!A:L",
    ).execute()
    return result.get("values", [])


def build_filename_index(sheet_data):
    """シートデータからD列（ファイル名）→ 行番号（1-based）のマップを作る。
    同時にE列（移動先）の値も保持する。
    Returns:
        dict: {filename: {"row": int, "dest": str}}
    """
    index = {}
    for i, row in enumerate(sheet_data):
        if i == 0:
            continue  # ヘッダーをスキップ
        if len(row) >= 4 and row[3]:
            dest = row[4] if len(row) >= 5 else ""
            index[row[3]] = {"row": i + 1, "dest": dest}  # 1-based行番号
    return index


def update_sheet_dest(sheets_service, row_number, dest_description):
    """既存行のE列（移動先）だけを更新する。"""
    sheets_service.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{SHEET_NAME}!E{row_number}",
        valueInputOption="USER_ENTERED",
        body={"values": [[dest_description]]},
    ).execute()


def append_to_sheet(sheets_service, rows):
    """保存ログシートに新しい行を追加する。"""
    body = {"values": rows}
    sheets_service.spreadsheets().values().append(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{SHEET_NAME}!A:L",
        valueInputOption="USER_ENTERED",
        insertDataOption="INSERT_ROWS",
        body=body,
    ).execute()


# ════════════════════════════════════════════════════════════════
# Gmail通知
# ════════════════════════════════════════════════════════════════

def send_notification(gmail_service, file_results):
    """佐藤信子・米谷尚起へ分類完了通知メールを送信する。"""
    body_lines = [
        "各位",
        "",
        "お疲れ様です。",
        "ファイル分類が完了しましたのでご報告いたします。",
        "",
        "【分類結果】",
    ]
    for fname, dest in file_results:
        body_lines.append(f"  ・{fname} → {dest}")
    body_lines.append("")
    body_lines.append("ご確認のほどよろしくお願いいたします。")
    body_lines.append(EMAIL_SIGNATURE)

    message = MIMEText("\n".join(body_lines), "plain", "utf-8")
    message["to"] = ", ".join(NOTIFY_TO)
    message["from"] = SENDER
    message["subject"] = NOTIFY_SUBJECT

    raw = base64.urlsafe_b64encode(message.as_bytes()).decode("utf-8")
    gmail_service.users().messages().send(
        userId="me", body={"raw": raw}
    ).execute()


# ════════════════════════════════════════════════════════════════
# メイン処理
# ════════════════════════════════════════════════════════════════

def scan_watch_folder():
    """監視フォルダ直下のファイル一覧を返す。
    除外対象：サブフォルダ、隠しファイル、ScanSnapフォルダ内ファイル。
    """
    files = []
    if not os.path.isdir(WATCH_FOLDER):
        print(f"監視フォルダが見つかりません: {WATCH_FOLDER}")
        return files
    for name in os.listdir(WATCH_FOLDER):
        if name.startswith("."):
            continue
        full = os.path.join(WATCH_FOLDER, name)
        # サブフォルダは除外（pre_trash, スキャナーから 等）
        if not os.path.isfile(full):
            continue
        # ScanSnapフォルダ内のファイルはGASが処理するので除外
        if os.path.abspath(full).startswith(os.path.abspath(SCANSNAP_FOLDER)):
            continue
        files.append((name, full))
    return files


def main():
    print("=== 秘書エージェント：ファイル分類開始 ===")

    # バックアップ
    print("--- バックアップ ---")
    backup_files()

    print(f"監視フォルダ: {WATCH_FOLDER}")

    # 認証
    creds = get_credentials()
    sheets_service = build("sheets", "v4", credentials=creds)
    gmail_service = build("gmail", "v1", credentials=creds)

    # 依頼者フォルダ読み込み
    clients = load_client_folders()
    print(f"依頼者フォルダ: {len(clients)}件")

    # スプレッドシートの既存データを取得
    sheet_data = fetch_sheet_data(sheets_service)
    filename_index = build_filename_index(sheet_data)
    print(f"スプレッドシート既存レコード: {len(filename_index)}件")

    # 監視フォルダスキャン
    files = scan_watch_folder()
    if not files:
        print("分類対象ファイルはありません。")
        return

    print(f"対象ファイル: {len(files)}件")

    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    new_rows = []
    updated_rows = []
    file_results = []
    skipped = 0

    for original_name, src_path in files:
        # スプレッドシートに既に存在し、E列（移動先）が埋まっている → 処理済みとしてスキップ
        if original_name in filename_index and filename_index[original_name]["dest"]:
            skipped += 1
            continue

        dest_dir, dest_description = classify_file(original_name, clients)
        new_name = rename_file(original_name)

        # ファイル移動
        moved_path = move_file(src_path, dest_dir, new_name)
        actual_name = os.path.basename(moved_path)

        print(f"  {original_name} → {dest_description}/{actual_name}")

        # スプレッドシートにファイル名が既に存在する場合 → E列を更新するだけ
        # リネーム前の名前でもリネーム後の名前でもマッチを確認
        existing = filename_index.get(original_name) or filename_index.get(actual_name)
        if existing:
            update_sheet_dest(sheets_service, existing["row"], dest_description)
            updated_rows.append(actual_name)
        else:
            # 新規行を追加
            # 保存日時 / メールタイトル / 送信者 / ファイル名 / 移動先 / 確認 / free / 弁革 / 経費 / 修正指示 / 済 / 対応事項
            new_rows.append([
                now, "", "", actual_name, dest_description,
                False, False, False, False, "", False, "",
            ])

        file_results.append((actual_name, dest_description))

    if skipped:
        print(f"処理済みスキップ: {skipped}件")

    # スプレッドシートに新規行を追加
    if new_rows:
        append_to_sheet(sheets_service, new_rows)
        print(f"スプレッドシートに{len(new_rows)}行追加しました。")

    if updated_rows:
        print(f"スプレッドシートの{len(updated_rows)}行を更新しました（移動先のみ）。")

    # 通知メール送信
    if file_results:
        send_notification(gmail_service, file_results)
        print(f"通知メールを送信しました: {', '.join(NOTIFY_TO)}")

    print("=== 分類完了 ===")


if __name__ == "__main__":
    main()
