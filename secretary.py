"""
弁護士法人Re-Start法律事務所 秘書エージェント
06_分類依頼フォルダのファイルを自動分類し、スプレッドシートに記録して佐藤信子へ通知する。
"""

import base64
import hashlib
import logging
import os
import re
import shutil
import time
import urllib.parse
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
UNKNOWN_FOLDER = os.path.join(WATCH_FOLDER, "分からなかった")
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

# ── 依頼者別名辞書 ──
# key: 依頼者フォルダ名の氏名部分（全角スペースあり）
# value: ファイル名に現れうる別名・通称のリスト（1つでも filename に含まれればマッチ）
CLIENT_ALIASES = {
    "ファム　ティ　フォン": ["ファム", "Pham", "Phạm", "phuong"],
}

# Drive APIでの folder_id 解決結果キャッシュ
FOLDER_ID_CACHE: dict = {}


# ════════════════════════════════════════════════════════════════
# Drive リンク取得・重複判定ヘルパー
# ════════════════════════════════════════════════════════════════

def _drive_escape(value: str) -> str:
    """Drive クエリ文字列用にバックスラッシュとシングルクオートをエスケープする。"""
    return value.replace("\\", "\\\\").replace("'", "\\'")


def resolve_drive_folder_id(drive_service, local_dir: str):
    """ローカルのGoogleドライブマウント配下ディレクトリを Drive フォルダID に変換する。
    見つからなければ None。
    """
    if not local_dir.startswith(BASE_PATH):
        return None
    relative = local_dir[len(BASE_PATH):].strip("/")
    if not relative:
        return "root"
    if relative in FOLDER_ID_CACHE:
        return FOLDER_ID_CACHE[relative]
    parent = "root"
    for part in relative.split("/"):
        q = (
            f"name = '{_drive_escape(part)}' and '{parent}' in parents "
            f"and mimeType = 'application/vnd.google-apps.folder' and trashed = false"
        )
        try:
            res = drive_service.files().list(
                q=q, fields="files(id)", pageSize=1, spaces="drive"
            ).execute()
        except Exception as e:
            logging.warning(f"Driveフォルダ解決失敗: {relative} ({e})")
            return None
        items = res.get("files", [])
        if not items:
            return None
        parent = items[0]["id"]
    FOLDER_ID_CACHE[relative] = parent
    return parent


def fetch_drive_file_link(drive_service, folder_id: str, filename: str,
                           retries: int = 2, delay: float = 3.0):
    """指定フォルダID内のファイル名から webViewLink を取得する。
    Driveアップロード遅延を考慮して軽くリトライする。見つからなければ None。
    """
    q = (
        f"name = '{_drive_escape(filename)}' and '{folder_id}' in parents "
        f"and trashed = false"
    )
    for attempt in range(retries + 1):
        try:
            res = drive_service.files().list(
                q=q, fields="files(id,webViewLink)", pageSize=1, spaces="drive"
            ).execute()
            items = res.get("files", [])
            if items:
                return items[0].get("webViewLink")
        except Exception as e:
            logging.warning(f"Driveリンク取得失敗: {filename} ({e})")
        if attempt < retries:
            time.sleep(delay)
    return None


def hyperlink_formula(url: str, label: str) -> str:
    """=HYPERLINK 式を組み立てる。ダブルクオートをエスケープ。"""
    safe_url = url.replace('"', '""')
    safe_label = label.replace('"', '""')
    return f'=HYPERLINK("{safe_url}","{safe_label}")'


def drive_search_url(filename: str) -> str:
    """Driveの検索URL（Drive API無効時のフォールバック用）。
    クリックすると指定ファイル名でドライブ内検索が開く。"""
    return "https://drive.google.com/drive/search?q=" + urllib.parse.quote(filename)


def files_identical(path_a: str, path_b: str) -> bool:
    """サイズとMD5ハッシュで2ファイルの内容一致を判定する。"""
    try:
        if os.path.getsize(path_a) != os.path.getsize(path_b):
            return False

        def md5(p):
            h = hashlib.md5()
            with open(p, "rb") as f:
                for chunk in iter(lambda: f.read(65536), b""):
                    h.update(chunk)
            return h.hexdigest()

        return md5(path_a) == md5(path_b)
    except OSError:
        return False


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
    """事件記録として分類できる場合、移動先パスを返す。
    依頼者名は client_name と CLIENT_ALIASES に列挙した別名のどちらでもマッチする。
    """
    for client in clients:
        names = [client["client_name"]] + CLIENT_ALIASES.get(client["client_name"], [])
        if not any(name in filename for name in names):
            continue
        # 依頼者が見つかった → サブフォルダを判定
        if client["case_type"] == "criminal":
            subfolder = _match_subfolder(filename, CRIMINAL_SUBFOLDER_KEYWORDS)
        elif client["case_type"] == "civil":
            subfolder = _match_subfolder(filename, CIVIL_SUBFOLDER_KEYWORDS)
        else:
            # 刑事/民事マーカーが無い案件（在留特別許可等）は双方のキーワードで試み、
            # 実在するサブフォルダのみ採用する。
            combined = {**CIVIL_SUBFOLDER_KEYWORDS, **CRIMINAL_SUBFOLDER_KEYWORDS}
            subfolder = _match_subfolder(filename, combined)
            if subfolder and not os.path.isdir(os.path.join(client["path"], subfolder)):
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

    # 3. 判断不能 → 分からなかったフォルダ（pre_trashは重複退避専用）
    return UNKNOWN_FOLDER, "06_分類依頼/分からなかった"


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


def update_sheet_row(sheets_service, row_number, d_value, dest_description):
    """既存行のD列（ファイル名＋リンク）とE列（移動先）をまとめて更新する。"""
    sheets_service.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{SHEET_NAME}!D{row_number}:E{row_number}",
        valueInputOption="USER_ENTERED",
        body={"values": [[d_value, dest_description]]},
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


def set_done_checkbox(sheets_service, row_number, value=True):
    """K列（済）チェックボックスを更新する。"""
    sheets_service.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{SHEET_NAME}!K{row_number}",
        valueInputOption="USER_ENTERED",
        body={"values": [[value]]},
    ).execute()


# ════════════════════════════════════════════════════════════════
# J列修正指示の反映
# ════════════════════════════════════════════════════════════════

def _fuzzy_match_subfolder(parent_dir, hint):
    """parent_dir 直下のサブフォルダから hint に最もマッチするものを返す。
    完全一致 > 番号プレフィックス除去後一致 > 部分一致 > キーワードマップ経由 の順。
    """
    if not os.path.isdir(parent_dir):
        return None
    hint_clean = hint.strip()
    hint_core = re.sub(r"^\d{2}_?", "", hint_clean)
    candidates = [
        d for d in os.listdir(parent_dir)
        if os.path.isdir(os.path.join(parent_dir, d)) and not d.startswith(".")
    ]
    if not candidates:
        return None
    for c in candidates:
        if c == hint_clean:
            return c
    for c in candidates:
        if re.sub(r"^\d{2}_?", "", c) == hint_core:
            return c
    if hint_core:
        for c in candidates:
            c_core = re.sub(r"^\d{2}_?", "", c)
            if hint_core and (hint_core in c_core or c_core in hint_core):
                return c
    combined = {**CRIMINAL_SUBFOLDER_KEYWORDS, **CIVIL_SUBFOLDER_KEYWORDS}
    for subfolder, keywords in combined.items():
        if any(kw in hint_clean for kw in keywords):
            for c in candidates:
                if subfolder == c or subfolder in c or c in subfolder:
                    return c
    return None


def _infer_correction_target(correction, current_dest, clients):
    """J列修正指示とE列現在位置から、(target_dir, relative_desc) を推定する。"""
    base_share = os.path.join(BASE_PATH, "共有用")
    norm = correction.strip().strip("/")

    # 1. フルパスっぽい指示（スラッシュを含む）
    if "/" in norm:
        candidate = os.path.join(base_share, norm)
        if os.path.isdir(candidate):
            return candidate, norm.rstrip("/") + "/"

    # 2. 現在と同じ依頼者内の別サブフォルダ
    cur_parts = [p for p in current_dest.strip("/").split("/") if p]
    if len(cur_parts) >= 2 and cur_parts[0] == "01_事件記録":
        client_folder = cur_parts[1]
        client_path = os.path.join(base_share, "01_事件記録", client_folder)
        sub = _fuzzy_match_subfolder(client_path, norm)
        if sub:
            return (
                os.path.join(client_path, sub),
                f"01_事件記録/{client_folder}/{sub}/",
            )

    # 3. 別の依頼者を指示文中に明示
    for client in clients:
        names = (
            [client["client_name"], client["folder_name"]]
            + CLIENT_ALIASES.get(client["client_name"], [])
        )
        if not any(name and name in norm for name in names):
            continue
        sub = _fuzzy_match_subfolder(client["path"], norm)
        if sub:
            return (
                os.path.join(client["path"], sub),
                f"01_事件記録/{client['folder_name']}/{sub}/",
            )

    # 4. 経理科目
    for category in ACCOUNTING_KEYWORDS.keys():
        if category in norm:
            return (
                os.path.join(ACCOUNTING, category),
                f"03_経理/{category}/",
            )

    return None, None


def apply_corrections(sheets_service, drive_service, clients):
    """保存ログのJ列（修正指示）を反映する。

    対象: F列（確認）未チェック かつ J列に指示あり かつ K列（済）未チェックの行。
    処理: 現在位置（E列＋ファイル名）からJ列の推定先へファイル移動、
         D列リンク・E列移動先を更新、K列をチェック。
    """
    sheet_data = fetch_sheet_data(sheets_service)
    base_share = os.path.join(BASE_PATH, "共有用")
    corrected = 0

    def _cell(row, idx):
        if len(row) <= idx or row[idx] is None:
            return ""
        return str(row[idx]).strip()

    for i, row in enumerate(sheet_data):
        if i == 0:
            continue  # ヘッダー
        if _cell(row, 5).upper() == "TRUE":  # F列 確認済
            continue
        correction = _cell(row, 9)
        if not correction:
            continue
        if _cell(row, 10).upper() == "TRUE":  # K列 既に済
            continue

        filename = _cell(row, 3)
        current_dest = _cell(row, 4)
        if not filename or not current_dest:
            continue

        current_dir = os.path.join(base_share, current_dest.strip("/"))
        current_path = os.path.join(current_dir, filename)
        if not os.path.isfile(current_path):
            logging.info(f"修正スキップ: ファイル未検出 {current_path}")
            print(f"  修正スキップ: ファイル未検出 {filename} ({current_dest})")
            continue

        target_dir, target_desc = _infer_correction_target(
            correction, current_dest, clients
        )
        if not target_dir:
            print(f"  修正指示の解釈失敗: {filename} (指示: '{correction}')")
            continue
        if os.path.normpath(target_dir) == os.path.normpath(current_dir):
            print(f"  修正先が現在位置と同一: {filename} → {target_desc}")
            continue

        os.makedirs(target_dir, exist_ok=True)
        new_path = os.path.join(target_dir, filename)
        if os.path.exists(new_path):
            if files_identical(current_path, new_path):
                os.remove(current_path)
                print(f"  修正反映（同内容既存のため元削除）: {filename} → {target_desc}")
            else:
                base, ext = os.path.splitext(filename)
                n = 1
                while os.path.exists(new_path):
                    new_path = os.path.join(target_dir, f"{base}_v{n}{ext}")
                    n += 1
                shutil.move(current_path, new_path)
                print(f"  修正反映（衝突回避）: {filename} → {target_desc}/{os.path.basename(new_path)}")
        else:
            shutil.move(current_path, new_path)
            print(f"  修正反映: {filename} → {target_desc}")

        actual_name = os.path.basename(new_path)
        folder_id = resolve_drive_folder_id(drive_service, target_dir)
        link = (
            fetch_drive_file_link(drive_service, folder_id, actual_name)
            if folder_id else None
        )
        url = link or drive_search_url(actual_name)
        d_cell = hyperlink_formula(url, actual_name)
        update_sheet_row(sheets_service, i + 1, d_cell, target_desc)
        set_done_checkbox(sheets_service, i + 1, True)
        corrected += 1

    if corrected:
        print(f"修正指示反映: {corrected}件")
    return corrected


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
    drive_service = build("drive", "v3", credentials=creds)

    # 依頼者フォルダ読み込み
    clients = load_client_folders()
    print(f"依頼者フォルダ: {len(clients)}件")

    # J列修正指示の反映（F未チェック & J指示あり & K未チェック の行）
    print("--- 修正指示の反映 ---")
    apply_corrections(sheets_service, drive_service, clients)

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
    pending_new = []   # (original_name, actual_name, d_cell, dest_description)
    updated_rows = []
    file_results = []
    skipped = 0
    duplicate_count = 0

    for original_name, src_path in files:
        # スプレッドシートに既に存在し、E列（移動先）が埋まっている → 処理済みとしてスキップ
        if original_name in filename_index and filename_index[original_name]["dest"]:
            skipped += 1
            continue

        dest_dir, dest_description = classify_file(original_name, clients)
        new_name = rename_file(original_name)

        # 重複検出: 移動先に同名かつ同一内容のファイルが既にある場合は pre_trash へ退避
        intended_path = os.path.join(dest_dir, new_name)
        if os.path.exists(intended_path) and files_identical(src_path, intended_path):
            os.makedirs(PRE_TRASH, exist_ok=True)
            dup_name = f"重複_{new_name}"
            dup_path = os.path.join(PRE_TRASH, dup_name)
            base, ext = os.path.splitext(dup_name)
            i = 1
            while os.path.exists(dup_path):
                dup_path = os.path.join(PRE_TRASH, f"{base}_{i}{ext}")
                i += 1
            shutil.move(src_path, dup_path)
            duplicate_count += 1
            logging.info(
                f"重複検出: {original_name} → pre_trash 退避（既存: {intended_path}）"
            )
            print(f"  重複検出: {original_name} → pre_trash/{os.path.basename(dup_path)}")

            # GASが追加した未処理行がある場合は「重複」として更新
            existing = filename_index.get(original_name) or filename_index.get(new_name)
            if existing:
                update_sheet_dest(
                    sheets_service,
                    existing["row"],
                    f"{dest_description}（重複・退避済）",
                )
                updated_rows.append(new_name)
            continue

        # ファイル移動
        moved_path = move_file(src_path, dest_dir, new_name)
        actual_name = os.path.basename(moved_path)

        # Drive上の webViewLink を取得して D列を HYPERLINK 式にする
        # Drive API無効時やまだ同期されていない場合は、検索URLへフォールバック
        folder_id = resolve_drive_folder_id(drive_service, dest_dir)
        drive_link = (
            fetch_drive_file_link(drive_service, folder_id, actual_name)
            if folder_id else None
        )
        link_url = drive_link or drive_search_url(actual_name)
        d_cell = hyperlink_formula(link_url, actual_name)

        print(f"  {original_name} → {dest_description}/{actual_name}"
              f"{' [直リンク]' if drive_link else ' [検索リンク]'}")

        # スプレッドシートにファイル名が既に存在する場合 → D列（リンク付き）とE列を更新
        existing = filename_index.get(original_name) or filename_index.get(actual_name)
        if existing:
            update_sheet_row(sheets_service, existing["row"], d_cell, dest_description)
            updated_rows.append(actual_name)
        else:
            pending_new.append((original_name, actual_name, d_cell, dest_description))

        file_results.append((actual_name, dest_description))

    if skipped:
        print(f"処理済みスキップ: {skipped}件")
    if duplicate_count:
        print(f"重複検出: {duplicate_count}件（pre_trashへ退避）")

    # append 直前にシートを再フェッチして GAS とのレース重複を防ぐ
    if pending_new:
        fresh_data = fetch_sheet_data(sheets_service)
        fresh_index = build_filename_index(fresh_data)
        truly_new = []
        for original_name, actual_name, d_cell, dest_description in pending_new:
            race_hit = (
                fresh_index.get(original_name) or fresh_index.get(actual_name)
            )
            if race_hit and race_hit["row"] not in {
                filename_index.get(original_name, {}).get("row"),
                filename_index.get(actual_name, {}).get("row"),
            }:
                # 当初の index には無く、再フェッチで見つかった = GAS等が追記した行
                update_sheet_row(
                    sheets_service, race_hit["row"], d_cell, dest_description
                )
                updated_rows.append(actual_name)
            else:
                truly_new.append([
                    now, "", "", d_cell, dest_description,
                    False, False, False, False, "", False, "",
                ])

        if truly_new:
            append_to_sheet(sheets_service, truly_new)
            print(f"スプレッドシートに{len(truly_new)}行追加しました。")

    if updated_rows:
        print(f"スプレッドシートの{len(updated_rows)}行を更新しました。")

    # 通知メール送信
    if file_results:
        send_notification(gmail_service, file_results)
        print(f"通知メールを送信しました: {', '.join(NOTIFY_TO)}")

    print("=== 分類完了 ===")


if __name__ == "__main__":
    main()
