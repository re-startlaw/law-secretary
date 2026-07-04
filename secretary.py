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
import sys
import time
import unicodedata
import urllib.parse
from datetime import datetime
from email.mime.text import MIMEText

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

try:
    import freee_filebox
except Exception:
    freee_filebox = None

try:
    from dotenv import load_dotenv

    load_dotenv(os.path.expanduser("~/law-secretary/secrets/.env"))
except ImportError:
    pass

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
SHEET_GID = 2139854819

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
LOCK_PATH = os.path.expanduser("~/law-secretary/secretary.lock")

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
    "専属AIエージェント\n"
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
    "03_検察官提出書面": [
        "検察官提出", "起訴状", "追起訴状",
        "令和8年検第※9063", "令和8年検第※9064", "令和8年検第※9067",
        "14160", "14161", "14164",
    ],
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
    "ふ_ファム・ティ・フォン": ["ファム", "Pham", "Phạm", "phuong", "ダオズイフン"],
    "ま_馬強": ["馬様", "馬強", "Angelina", "アンジェリーナ"],
    "た_田村正宣": [
        "田村", "田中成子", "有限会社タナカ", "有限会社タナ力",
        "令和8年検第※9063", "令和8年検第※9064", "令和8年検第※9067",
        "14160", "14161", "14164",
    ],
    "す_鈴木七海": ["鈴木七海", "服部咲"],
}

# 送信者メール（小文字）→ 依頼者の client_name（CLIENT_ALIASES のキーと同じ表記）。
# 「分からなかった」になった定番送信元をここに追補していく。
SENDER_EMAIL_TO_CLIENT: dict[str, str] = {
    "nao@teruiglobal.jp": "ふ_ファム・ティ・フォン",
    "martin.ma@letour.co.jp": "ま_馬強",
}

# 手動で登録した固定エントリの凍結コピー。main() で保存ログ学習を適用したあと、
# この手動エントリを後勝ちで再適用し、学習が固定値を上書きしないようにする。
_MANUAL_SENDER_EMAIL_TO_CLIENT: dict[str, str] = dict(SENDER_EMAIL_TO_CLIENT)

# 同一ドメインが複数依頼者に紐づく場合は登録しないこと。
SENDER_DOMAIN_TO_CLIENT: dict[str, str] = {}

# Drive APIでの folder_id 解決結果キャッシュ
FOLDER_ID_CACHE: dict = {}

LLM_LOW_CONFIDENCE_LOG: list[dict] = []

# (送信者キー, 件名キー) → 共有用配下の相対パス
# main() 起動時に保存ログから build_sender_subject_routing() で構築・更新する。
LEARNED_ROUTING: dict[tuple[str, str], str] = {}


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


def fetch_drive_file_meta(drive_service, folder_id: str, filename: str,
                           retries: int = 3, delay: float = 2.0):
    """指定フォルダID内のファイル名から id と webViewLink を取得する。
    Driveアップロード遅延を考慮して軽くリトライする。見つからなければ None。
    Returns: dict({"id": str, "webViewLink": str}) | None
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
                return {
                    "id": items[0].get("id"),
                    "webViewLink": items[0].get("webViewLink"),
                }
        except Exception as e:
            logging.warning(f"Driveメタ取得失敗: {filename} ({e})")
        if attempt < retries:
            time.sleep(delay)
    return None


def hyperlink_formula(url: str, label: str) -> str:
    """=HYPERLINK 式を組み立てる。ダブルクオートをエスケープ。"""
    safe_url = url.replace('"', '""')
    safe_label = label.replace('"', '""')
    return f'=HYPERLINK("{safe_url}","{safe_label}")'


def build_drive_link_from_id(file_id: str) -> str:
    """file ID から汎用Drive URL を組み立てる。webViewLink が取れない時の保険。"""
    return f"https://drive.google.com/open?id={file_id}"


DRIVE_ID_RE = re.compile(r"/d/([A-Za-z0-9_-]{20,})|[?&]id=([A-Za-z0-9_-]{20,})")


def extract_drive_id_from_formula(cell: str):
    """既存D列の HYPERLINK 式やDrive URL から file ID を抽出する。
    `https://drive.google.com/file/d/<ID>/view` や `?id=<ID>` 形式に対応。
    Returns: str | None
    """
    if not cell:
        return None
    m = DRIVE_ID_RE.search(cell)
    if not m:
        return None
    return m.group(1) or m.group(2)


def drive_search_url(filename: str) -> str:
    """Drive内検索URL。最終フォールバック専用（Drive API取得失敗・オフライン時）。
    クリックすると指定ファイル名でドライブ内検索が開く。"""
    return "https://drive.google.com/drive/search?q=" + urllib.parse.quote(filename)


def _fetch_d_column_id(sheets_service, row_number: int):
    """指定行のD列セル（HYPERLINK式）から file ID を抽出する。Driveメタ取得失敗時の最終フォールバック。"""
    try:
        res = sheets_service.spreadsheets().values().get(
            spreadsheetId=SPREADSHEET_ID,
            range=f"{SHEET_NAME}!D{row_number}",
            valueRenderOption="FORMULA",
        ).execute()
        values = res.get("values", [])
        if values and values[0]:
            return extract_drive_id_from_formula(str(values[0][0]))
    except Exception as e:
        logging.warning(f"D列式取得失敗 row={row_number}: {e}")
    return None


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

def _extract_emails_from_sender(sender: str) -> list[str]:
    """送信者文字列からメールアドレスを列挙（小文字化）。"""
    if not sender:
        return []
    return [
        m.lower()
        for m in re.findall(
            r"[\w.+-]+@[\w.-]+\.[A-Za-z]{2,}",
            sender,
            flags=re.IGNORECASE,
        )
    ]


def normalize_sender_for_context(sender: str) -> str:
    """C列の送信者をマッチ用に展開（表示名・メール・ローカル部・ドメイン）。"""
    if not sender or not sender.strip():
        return ""
    s = sender.strip()
    tokens: list[str] = [s]
    for em in _extract_emails_from_sender(s):
        tokens.append(em)
        local, _, domain = em.partition("@")
        if local:
            tokens.append(local)
        if domain:
            tokens.append(domain)
    m = re.match(
        r'^(.+?)\s*<[\w.+-]+@[\w.-]+\.[A-Za-z]{2,}>\s*$',
        s,
        flags=re.DOTALL,
    )
    if m:
        disp = m.group(1).strip().strip('"').strip("'")
        if disp:
            tokens.append(disp)
    seen: set[str] = set()
    out: list[str] = []
    for t in tokens:
        t = t.strip()
        if not t:
            continue
        key = t.lower()
        if key in seen:
            continue
        seen.add(key)
        out.append(t)
    return " ".join(out)


def build_sender_email_index(
    sheet_data: list[list[str]], clients: list[dict]
) -> tuple[dict[str, str], dict[str, int]]:
    """保存ログのC列(送信者)→E列(移動先)から「メアド→依頼者folder_name」マップを構築。

    曖昧排除:
      - 同一メアドが複数依頼者にマップされたら除外
      - 経理 / pre_trash / 分からなかった 行は依頼者特定にならないので無視

    Returns:
        (index, stats) — index は {email: folder_name}、stats は件数情報
    """
    folder_names = {c["folder_name"] for c in clients}
    candidates: dict[str, set[str]] = {}
    rows_total = 0
    rows_skipped = 0
    for i, row in enumerate(sheet_data):
        if i == 0:
            continue
        if len(row) < 5:
            continue
        sender = row[2] if len(row) >= 3 else ""
        dest = row[4] if len(row) >= 5 else ""
        if not sender or not dest:
            continue
        rows_total += 1
        # 移動先の最初のセグメントが 01_事件記録 かどうか
        segs = dest.replace("\\", "/").strip("/").split("/")
        if len(segs) < 2 or segs[0] != "01_事件記録":
            rows_skipped += 1
            continue
        folder = segs[1]
        if folder not in folder_names:
            rows_skipped += 1
            continue
        for em in _extract_emails_from_sender(sender):
            if _is_generic_sender_key(em):
                continue
            candidates.setdefault(em, set()).add(folder)

    index: dict[str, str] = {}
    ambiguous = 0
    for em, folders in candidates.items():
        if len(folders) == 1:
            index[em] = next(iter(folders))
        else:
            ambiguous += 1
    stats = {
        "rows_scanned": rows_total,
        "rows_skipped": rows_skipped,
        "emails_indexed": len(index),
        "emails_ambiguous": ambiguous,
    }
    return index, stats


def _sender_keys(sender: str) -> list[str]:
    """送信者文字列からマッチ用キーを列挙する（メアド・ローカル部・表示名）。"""
    if not sender:
        return []
    keys: list[str] = []
    seen: set[str] = set()

    def _add(k: str) -> None:
        k = k.strip().lower()
        if k and k not in seen:
            seen.add(k)
            keys.append(k)

    for em in _extract_emails_from_sender(sender):
        _add(em)

    m = re.match(
        r'^(.+?)\s*<[\w.+-]+@[\w.-]+\.[A-Za-z]{2,}>\s*$',
        sender,
        flags=re.DOTALL,
    )
    if m:
        disp = m.group(1).strip().strip('"').strip("'")
        if disp:
            _add(disp)
    elif "@" not in sender:
        _add(sender)
    return keys


def normalize_subject_for_key(subject: str) -> str:
    """件名を学習・照合用キーに正規化する。

    - 先頭の `Re:` `Fwd:` `RE:` `FWD:` の連続を除去
    - 先頭の ML 接頭辞 `[name:01510]` を `[name]` に丸める（番号だけ削除）
    - 連続空白（半角・全角）を1個に圧縮
    - 小文字化・前後空白除去
    """
    if not subject:
        return ""
    s = subject
    s = re.sub(r"^(\s*(re|fw|fwd)\s*[:\uff1a]+\s*)+", "", s, flags=re.IGNORECASE)
    s = re.sub(r"\[([^\[\]:\uff1a]+)[:\uff1a]\s*\d+\]", r"[\1]", s)
    s = re.sub(r"[\s\u3000]+", " ", s).strip()
    return s.lower()


# スキャナー・eFax等の汎用送信元と空件名は (送信者, 件名) 学習・照合の対象外。
# 全スキャンが同一キー（例: 送信者=米谷スキャン/件名=なし）に潰れ、過去に件数の
# 多かった保存先へ無関係書類が誤ルーティングされる事故（260701: 税金通知が
# た_田村正宣/02_身体拘束関係 行き）を防ぐ。
_GENERIC_SENDER_KEYS = {"米谷スキャン", "スキャン", "scansnap", "scan", "efax"}
_GENERIC_SENDER_EMAIL_RE = re.compile(
    r"^(no-?reply|donotreply|scan\w*|fax\w*|efax\w*)@|efax", re.IGNORECASE
)
_GENERIC_SUBJECT_KEYS = {
    "", "なし", "無題", "no subject", "(no subject)", "untitled", "スキャン",
}


def _is_generic_sender_key(key: str) -> bool:
    """スキャナー・eFax・noreply等、案件特定に使えない送信者キーか。"""
    k = key.strip().lower()
    if k in _GENERIC_SENDER_KEYS:
        return True
    if "@" in k and _GENERIC_SENDER_EMAIL_RE.search(k):
        return True
    return False


def _is_generic_subject_key(subj_key: str) -> bool:
    """「なし」「無題」等、書類種別の情報を持たない件名キーか。"""
    return subj_key.strip().lower() in _GENERIC_SUBJECT_KEYS


def build_sender_subject_routing(
    sheet_data: list[list[str]],
) -> tuple[dict[tuple[str, str], str], dict[str, int]]:
    """保存ログ B/C/E 列から (送信者キー, 件名キー) → 共有用配下の相対パス を学習する。

    学習対象:
      - dest が空でない、かつ「分からなかった」「pre_trash」「06_分類依頼」を含まない行
      - dest 末尾の `（重複・退避済）` 等の括弧注記は除去してから集計
    曖昧排除:
      - 同じ (送信者, 件名) で複数の dest が矛盾する場合、最多が次点の2倍以上なら採用、
        そうでなければ除外。
    """
    counts: dict[tuple[str, str], dict[str, int]] = {}
    skipped = 0
    for i, row in enumerate(sheet_data):
        if i == 0:
            continue
        if len(row) < 5:
            continue
        subject = row[1] if len(row) >= 2 else ""
        sender = row[2] if len(row) >= 3 else ""
        dest = row[4] if len(row) >= 5 else ""
        if not sender or not dest:
            continue
        dest_clean = re.sub(r"[（(][^（()）]*[）)]\s*$", "", dest).strip()
        if not dest_clean:
            skipped += 1
            continue
        if (
            "分からなかった" in dest_clean
            or "06_分類依頼" in dest_clean
            or "pre_trash" in dest_clean
        ):
            skipped += 1
            continue
        subj_key = normalize_subject_for_key(subject)
        if not subj_key or _is_generic_subject_key(subj_key):
            continue
        for sk in _sender_keys(sender):
            if _is_generic_sender_key(sk):
                continue
            cell = counts.setdefault((sk, subj_key), {})
            cell[dest_clean] = cell.get(dest_clean, 0) + 1

    routing: dict[tuple[str, str], str] = {}
    ambiguous = 0
    for key, dests in counts.items():
        if len(dests) == 1:
            routing[key] = next(iter(dests.keys()))
            continue
        ranked = sorted(dests.items(), key=lambda kv: -kv[1])
        if ranked[0][1] >= ranked[1][1] * 2:
            routing[key] = ranked[0][0]
        else:
            ambiguous += 1

    stats = {
        "learned": len(routing),
        "ambiguous": ambiguous,
        "skipped_rows": skipped,
    }
    return routing, stats


def lookup_learned_route(sender: str, subject: str) -> str | None:
    """(送信者, 件名) で学習済みルーティングをルックアップする。"""
    if not LEARNED_ROUTING:
        return None
    subj_key = normalize_subject_for_key(subject or "")
    if not subj_key or _is_generic_subject_key(subj_key):
        return None
    for sk in _sender_keys(sender or ""):
        if _is_generic_sender_key(sk):
            continue
        rel = LEARNED_ROUTING.get((sk, subj_key))
        if rel:
            return rel
    return None


def sender_routing_hints(sender: str) -> str:
    """登録済みメール／ドメインから依頼者名を推定し haystack 用の断片を返す。"""
    if not sender:
        return ""
    hints: list[str] = []
    seen_name: set[str] = set()
    for em in _extract_emails_from_sender(sender):
        cn = SENDER_EMAIL_TO_CLIENT.get(em)
        if cn and cn not in seen_name:
            hints.append(cn)
            seen_name.add(cn)
        if "@" in em:
            dom = em.split("@", 1)[1]
            cn_d = SENDER_DOMAIN_TO_CLIENT.get(dom)
            if cn_d and cn_d not in seen_name:
                hints.append(cn_d)
                seen_name.add(cn_d)
    return " ".join(hints)


def _norm_text(value: str) -> str:
    """濁点分解や全半角揺れを分類前に寄せる。"""
    return unicodedata.normalize("NFKC", value or "")


def classify_special_document(filename, context=""):
    """事務所運用上、経理より優先する固定分類ルール。"""
    haystack = _norm_text(f"{filename} {context}")
    if "プリア常盤台" in haystack or "プリアときわ台" in haystack:
        return os.path.join(BASE_PATH, "共有用", "04_事務", "各種書類", "社宅関係")
    return None


def classify_accounting(filename, context=""):
    """経理書類として分類できる場合、移動先パスを返す。
    context は送信者・メールタイトルなど、ファイル名以外のヒント文字列。
    """
    haystack = _norm_text(f"{filename} {context}")
    keywords_in_name = ["領収書", "請求書", "精算"]
    is_accounting = any(kw in haystack for kw in keywords_in_name)

    matched_category = None
    for category, keywords in ACCOUNTING_KEYWORDS.items():
        if any(kw in haystack for kw in keywords):
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


def classify_case_record(filename, clients, context=""):
    """事件記録として分類できる場合、移動先パスを返す。
    依頼者名は client_name と CLIENT_ALIASES に列挙した別名のどちらでもマッチする。
    context（送信者・件名）も依頼者名・サブフォルダ判定の検索対象にする。
    """
    haystack = _norm_text(f"{filename} {context}")
    for client in clients:
        names = (
            [client["folder_name"], client["client_name"]]
            + CLIENT_ALIASES.get(client["client_name"], [])
        )
        names = [_norm_text(n) for n in names if n]
        if not any(name in haystack for name in names):
            continue
        # 依頼者が見つかった → サブフォルダを判定して移動先を返す
        return _resolve_client_dest(client, haystack)
    return None


def _llm_case_dest_allowed(rel: str, clients, guard_text: str) -> bool:
    """LLMが返した保存先 rel が 01_事件記録/<依頼者> の場合、その依頼者の
    氏名・別名が guard_text（元名＋送信者＋件名＋本文）に実在するか検証する。

    - rel が事件記録でない、または依頼者フォルダを特定できない → True（対象外）
    - 依頼者を特定でき、その名称が guard_text に一切現れない → False（誤割当として拒否）

    保存ログのキーワードヒント過学習で無関係文書が活発案件へ誤爆するのを防ぐ。
    外形分類 classify_case_record と同じ「依頼者名が実在する」基準を LLM にも課す。
    """
    if not rel.startswith("01_事件記録/"):
        return True
    parts = rel.split("/")
    folder_name = parts[1] if len(parts) >= 2 else ""
    client = next(
        (c for c in clients if c.get("folder_name") == folder_name), None
    )
    if not client:
        return True  # 依頼者を特定できないときは判定せず通す（保守的）
    names = (
        [client.get("folder_name", ""), client.get("client_name", "")]
        + CLIENT_ALIASES.get(client.get("client_name", ""), [])
    )
    names = [_norm_text(n) for n in names if n]
    hay = _norm_text(guard_text)
    return any(n in hay for n in names)


def _select_case_subfolder(client, haystack):
    """依頼者フォルダ内の適切なサブフォルダ名を返す（無ければ None）。

    案件種別（criminal/civil/unknown）に応じてキーワードマップを切り替える。
    刑事/民事マーカーが無い案件（unknown・在留特別許可等）は双方のキーワードで
    試み、**実在するサブフォルダのみ**採用する（存在しないサブフォルダへの誤割当を防ぐ）。
    """
    if client["case_type"] == "criminal":
        return _match_subfolder(haystack, CRIMINAL_SUBFOLDER_KEYWORDS)
    if client["case_type"] == "civil":
        return _match_subfolder(haystack, CIVIL_SUBFOLDER_KEYWORDS)
    combined = {**CIVIL_SUBFOLDER_KEYWORDS, **CRIMINAL_SUBFOLDER_KEYWORDS}
    subfolder = _match_subfolder(haystack, combined)
    if subfolder and not os.path.isdir(os.path.join(client["path"], subfolder)):
        subfolder = None
    return subfolder


def _resolve_client_dest(client, haystack):
    """依頼者が確定した後、サブフォルダ込みの移動先パスを返す。"""
    subfolder = _select_case_subfolder(client, haystack)
    if subfolder:
        return os.path.join(client["path"], subfolder)
    # サブフォルダが特定できない → 依頼者フォルダ直下
    return client["path"]


def _match_subfolder(text, keyword_map):
    """キーワードマップからサブフォルダ名を返す。"""
    norm_text = _norm_text(text)
    for subfolder, keywords in keyword_map.items():
        if any(_norm_text(kw) in norm_text for kw in keywords):
            return subfolder
    return None


def classify_by_sender_email(filename, clients, sender="", subject=""):
    """送信者メールアドレスから依頼者が一意に確定できる場合、移動先を返す。

    SENDER_EMAIL_TO_CLIENT（手動固定＋保存ログ学習）でメアド→依頼者を解決し、
    文字列再マッチ（os.listdir 先勝ちの脆さ）に頼らず該当依頼者へ確定的に振り分ける。
    ファイル名・中身に依頼者名が無くても送信者だけで案件が決まる。

    Returns:
        tuple(dest_path, rel) または None（送信者で確定できないとき）。
    """
    if not sender:
        return None
    target = None
    for em in _extract_emails_from_sender(sender):
        cn = SENDER_EMAIL_TO_CLIENT.get(em)
        if cn:
            target = cn
            break
    if not target:
        return None
    # 学習値は folder_name 表記、手動値は client_name 表記。全角スペース依頼者では
    # 両者が食い違うため、client_name と folder_name の両方で照合する。
    client = next(
        (
            c for c in clients
            if target in (c.get("client_name"), c.get("folder_name"))
        ),
        None,
    )
    if not client:
        logging.warning(
            f"送信者→依頼者 解決名がフォルダ一覧に無い: {target!r}（sender={sender!r}）"
        )
        return None
    haystack = _norm_text(f"{filename} {subject}")
    dest = _resolve_client_dest(client, haystack)
    rel = os.path.relpath(dest, os.path.join(BASE_PATH, "共有用"))
    logging.info(
        f"送信者メアド確定: {filename} → {rel} (sender={sender!r})"
    )
    return dest, rel


def classify_file(filename, clients, sender="", subject=""):
    """ファイル名＋送信者＋メールタイトルから分類先のパスを返す。
    判断不能なら分からなかったフォルダを返す。

    Args:
        filename: ファイル名（命名規則どおりの論理名。`rename_file` 後の名前を渡すこと）
        clients: 依頼者フォルダ一覧
        sender: 保存ログC列（送信者）— スプレッドシートから渡す
        subject: 保存ログB列（メールタイトル）— スプレッドシートから渡す

    Returns:
        tuple: (dest_path, description)
    """
    sctx = normalize_sender_for_context(sender or "")
    rh = sender_routing_hints(sender or "")
    context = " ".join(
        p for p in (sctx, rh, subject or "") if p
    ).strip()

    # 0. 学習済み (送信者, 件名) ルーティング — 同じ送信者から同じタイトルで届いたら過去の保存先へ
    learned_rel = lookup_learned_route(sender or "", subject or "")
    if learned_rel:
        abs_dest = os.path.join(BASE_PATH, "共有用", learned_rel)
        if os.path.isdir(abs_dest):
            logging.info(
                f"学習ルーティング適用: {filename} → {learned_rel} "
                f"(sender={sender!r}, subject={subject!r})"
            )
            return abs_dest, learned_rel
        # 学習済みフォルダが消えている場合は無視して通常分類に進む
        logging.info(f"学習ルーティング先が消失: {learned_rel}（無視して通常分類）")

    # 1. 固定ルール（社宅など、請求書でも経理へ入れないもの）
    dest = classify_special_document(filename, context)
    if dest:
        rel = os.path.relpath(dest, os.path.join(BASE_PATH, "共有用"))
        return dest, rel

    # 2. 送信者メアド確定（既知送信者なら依頼者を一意に確定。ファイル名・中身に
    #    依頼者名が無くても案件フォルダへ。経理より前＝依頼者本人発の請求書等が
    #    案件フォルダでなく経理へ吸われる回帰を避ける）
    sender_dest = classify_by_sender_email(filename, clients, sender, subject)
    if sender_dest:
        return sender_dest

    # 3. 事件記録チェック（ファイル名・件名から依頼者名でマッチ）
    dest = classify_case_record(filename, clients, context)
    if dest:
        rel = os.path.relpath(dest, os.path.join(BASE_PATH, "共有用"))
        return dest, rel

    # 4. 経理書類チェック（ベンダー送信者や件名から経理勘定を推定）
    dest = classify_accounting(filename, context)
    if dest:
        rel = os.path.relpath(dest, os.path.join(BASE_PATH, "共有用"))
        return dest, rel

    # 5. 判断不能 → 分からなかったフォルダ（pre_trashは重複退避専用）
    return UNKNOWN_FOLDER, "06_分類依頼/分からなかった"


# ════════════════════════════════════════════════════════════════
# LLM フォールバック分類（外形で決まらないファイル用）
# ════════════════════════════════════════════════════════════════

LLM_MODEL = "claude-sonnet-4-6"
LLM_MAX_TOKENS = 800

LLM_SYSTEM_RULES = """あなたは弁護士法人Re-Start法律事務所のファイル分類アシスタントです。
渡されたファイルの中身と送信者・件名から、保存先と最終ファイル名を決定してください。

【ファイル命名規則】
- 日付6桁_書類名.拡張子 （例: 260413_準抗告申立書.pdf）
- 既に正規名なら維持。スキャナー出力等の暫定名は中身から書類名を決める。
- 日付が中身から読み取れない場合は元ファイル名の日付があればそれを使い、なければ送信日付を使う。

【保存先（dest_relative_path）の形式】 — 共有用配下の相対パスを返す
- 事件記録: `01_事件記録/<依頼者フォルダ名>/<サブフォルダ>`
- 経理: `03_経理/<勘定科目>`
- 営業: `11_営業/...`（既存サブ: ココナラGBP / 事務所写真 / 印刷物用ラージサイズ など）
- 委員会: `12_委員会/...`（既存サブ: 刑事弁護委員会/<年度>/、広報部会/、法廷弁護技術研修(...)、外国人権利、性犯罪関係... など）
- 研修: `13_研修/...`（既存サブ: TATAゼミ、倫理研修、単発の公演）
- 事務一般: `04_事務/各種書類/<取引先>` のように、契約・規約・お知らせ等は取引先別サブへ
  （既存サブ: Regus / Legalscape / D1-Law / 日弁連_東弁会 / 日本司法支援センター / 法テラス霞ヶ関 / 秘書 など）
- 不明な場合は信頼度を low にして dest_relative_path は空文字でよい

【刑事案件サブフォルダ】
01_選任関係（選任・委任）/ 02_身体拘束関係（勾留・逮捕・保釈）/ 03_検察官提出書面（起訴状）/
04_弁護人提出書面（準抗告・意見書）/ 05_検察官証拠（証拠開示）/ 06_裁判所手続（期日）/
07_収集資料 / 08_論文・文献 / 10_メモ / 11_弁護人請求証拠 / 12_精算？

【民事案件サブフォルダ】
00主張（準備書面）/ 01甲号証 / 02乙号証 / 03連絡文書（通知）/ 04事務（送付書）/
05期日報告書 / 06資料 / 07_ワードファイル / 委任関係（委任状・受任）

【経理勘定科目】
地代家賃 / 通信費 / 外注費（接見・弁護士費用）/ 会議費 / 接待交際費 /
旅費交通費（タクシー・電車・PASMO/Suica）/ 消耗品費 / 広告宣伝費 /
支払手数料 / 研修費 / 新聞図書費 / 諸会費

【判断のヒント】
- 委員会メーリスの件名（[saiban-yousei:nnnn] / [keijibengo:nnnn] / 東弁・三会など）は 12_委員会 配下の該当サブが第一候補。
- ココナラ・GBP・事例紹介・広告系は 11_営業 配下が第一候補。
- リージャス・Legalscape・法テラス・日弁連など事務所運営の取引先からのお知らせ・契約は 04_事務/各種書類/<取引先>。
- 経理（請求書・領収書・支払案内）は 03_経理 を優先。

【信頼度】
- high: 中身から依頼者・書類種別が明確
- medium: 依頼者は推定だが書類種別は明確、もしくは逆
- low: 確証なし。dest_relative_path は空でよい

【出力】
必ず submit_classification ツールを使って構造化データで返してください。"""


_KEYWORD_RE = re.compile(
    r"[一-鿿]{2,}"          # 漢字2字以上
    r"|[ァ-ー]{3,}"         # カタカナ3字以上（長音含む）
    r"|[A-Za-z][A-Za-z0-9_\-]{3,}"  # 英数字混じり4字以上
)

# 候補語抽出から除外する一般語・分類フォルダ名の語幹
_KEYWORD_STOP = {
    "ファイル", "中身", "先頭", "抜粋", "送信者", "件名", "依頼者", "フォルダ",
    "一覧", "事件記録", "経理", "事務", "委員会", "研修", "営業", "分類依頼",
    "分からなかった", "領収書", "請求書", "送付書", "準備書面", "保存ログ",
    "Re-Start", "re-startlaw", "kometani", "gmail", "@gmail", "com", "co",
}


def build_keyword_dest_index(sheet_data) -> dict[str, list[str]]:
    """保存ログ全行をキーワード→dest頻度マップに変換する。

    各行の B(件名) / C(送信者) / D(ファイル名) / E(移動先) を連結したテキストを保持。
    `lookup_keyword_dests(idx, word)` で部分一致 → dest を頻度集計する。
    """
    rows: list[tuple[str, str]] = []
    for row in sheet_data:
        subject = str(row[1]) if len(row) > 1 else ""
        sender = str(row[2]) if len(row) > 2 else ""
        filename = str(row[3]) if len(row) > 3 else ""
        dest = str(row[4]) if len(row) > 4 else ""
        if not dest:
            continue
        haystack = f"{subject} | {sender} | {filename}"
        rows.append((haystack, dest))
    return {"rows": rows}  # type: ignore[return-value]


def _extract_keyword_candidates(*texts: str, limit: int = 12) -> list[str]:
    """テキスト群から固有名詞らしき語の候補を抽出する。

    漢字 4 字以上の語は、人名・固有名詞の部分一致を可能にするため
    先頭 2/3 字のプレフィックスも候補に含める（例「大宮瑞希様」→「大宮」「大宮瑞」も）。
    """
    seen: dict[str, int] = {}
    for text in texts:
        if not text:
            continue
        for m in _KEYWORD_RE.finditer(text):
            word = m.group(0)
            if word in _KEYWORD_STOP:
                continue
            if len(word) > 30:
                continue
            seen[word] = seen.get(word, 0) + 1
            # 漢字長語の prefix も追加（人名・社名の連結を救う）
            if re.fullmatch(r"[一-鿿]+", word) and len(word) >= 4:
                for prefix_len in (2, 3):
                    prefix = word[:prefix_len]
                    if prefix in _KEYWORD_STOP:
                        continue
                    seen[prefix] = max(seen.get(prefix, 0), 1)
    # 出現回数の多い順、同点なら長い語を優先
    return sorted(seen, key=lambda w: (-seen[w], -len(w)))[:limit]


def _lookup_keyword_dests(index, words: list[str], max_per_word: int = 3) -> list[str]:
    """候補語ごとに保存ログ全行を substring match → (word, dest, count) の集計行を返す。

    06_分類依頼/* 配下（pre_trash・分からなかった）への過去ヒットはノイズになるため除外。
    """
    if not index or not isinstance(index, dict):
        return []
    rows = index.get("rows", [])
    if not rows:
        return []
    out: list[str] = []
    for word in words:
        dest_counts: dict[str, int] = {}
        for haystack, dest in rows:
            if word in haystack:
                dest_counts[dest] = dest_counts.get(dest, 0) + 1
        if not dest_counts:
            continue
        ranked = [
            (d, c) for d, c in sorted(dest_counts.items(), key=lambda x: -x[1])
            if not d.startswith("06_分類依頼/")
        ]
        for dest, count in ranked[:max_per_word]:
            out.append(f'  - "{word}" → {dest}（過去 {count} 件）')
    return out


def _build_llm_user_prompt(
    original_name: str,
    sender: str,
    subject: str,
    extracted_text: str,
    clients: list[dict],
    sheet_keyword_index=None,
) -> str:
    client_lines = "\n".join(
        f"- {c['folder_name']} ({c['case_type']})" for c in clients
    )

    sheet_hint_section = ""
    if sheet_keyword_index is not None:
        candidates = _extract_keyword_candidates(
            original_name, sender, subject, extracted_text[:3000], limit=20
        )
        hits = _lookup_keyword_dests(sheet_keyword_index, candidates, max_per_word=2)
        if hits:
            sheet_hint_section = (
                "\n【関連する過去エントリ（保存ログ全文ヒット — 固有名詞ベース）】\n"
                + "\n".join(hits[:20])
                + "\n"
            )

    return f"""以下のファイルを分類してください。

【元ファイル名】 {original_name}
【メール送信者】 {sender or '(不明)'}
【メール件名】 {subject or '(不明)'}

【依頼者フォルダ一覧】
{client_lines}
{sheet_hint_section}
【ファイル中身（先頭抜粋）】
{extracted_text}
"""


_ACCOUNTING_DOC_TYPES = {"領収書", "請求書", "取引明細", "クレカ明細", "明細"}


def _build_accounting_filename(
    original_ext: str,
    vendor: str,
    amount,
    doc_type: str,
    accounting_category: str,
    fallback_date: str | None = None,
) -> str | None:
    """経理ファイルを `YYYYMMDD_業者名_金額_書類種別_勘定科目.<ext>` に整形する。

    すべてのフィールドが揃った場合のみ整形名を返し、揃わなければ None。
    日付は fallback_date（YYYYMMDD or YYMMDD）を採用し、無ければ None。
    """
    if not (vendor and doc_type and accounting_category):
        return None
    if doc_type not in _ACCOUNTING_DOC_TYPES:
        return None
    try:
        amount_int = int(str(amount).replace(",", "").replace("円", "").strip())
    except (ValueError, TypeError):
        return None
    if amount_int <= 0:
        return None
    if not fallback_date:
        return None
    if re.fullmatch(r"\d{6}", fallback_date):
        # YYMMDD → YYYYMMDD（2000年代と仮定）
        fallback_date = "20" + fallback_date
    if not re.fullmatch(r"\d{8}", fallback_date):
        return None
    # ファイル名に使えない文字を除去
    vendor_clean = re.sub(r"[\s/\\:\*\?\"<>\|]", "", vendor).strip()
    if not vendor_clean:
        return None
    return f"{fallback_date}_{vendor_clean}_{amount_int}_{doc_type}_{accounting_category}{original_ext}"


def _extract_date_token(*texts: str) -> str | None:
    """テキスト群から YYYYMMDD or YYMMDD 日付トークンを最初の1件返す。"""
    for text in texts:
        if not text:
            continue
        m = re.search(r"(20\d{6}|\d{6})_", text)
        if m:
            return m.group(1)
    return None


def classify_by_content(
    src_path: str,
    original_name: str,
    sender: str,
    subject: str,
    clients: list[dict],
    sheet_keyword_index=None,
) -> tuple[str, str, str, str, dict] | None:
    """中身を1回抽出してLLMに投げ、(canonical_name, dest_dir, dest_relative, reasoning, extras) を返す。

    LLM呼び出しに失敗・低信頼・不正パスの場合は None を返す（UNKNOWN_FOLDERへ）。
    画像ファイル（HEIC/JPG/PNG）の場合は vision モードで処理する。
    extras dict には requires_bengokakumei, accounting_category, vendor, amount, document_type を含む。
    """
    if not os.environ.get("ANTHROPIC_API_KEY"):
        logging.warning("ANTHROPIC_API_KEY 未設定: LLMフォールバックをスキップ")
        return None

    # 内容抽出
    sys.path.insert(0, os.path.join(os.path.dirname(__file__), "scripts"))
    try:
        from extract_text import extract_text
    except ImportError as e:
        logging.error(f"extract_text のimport失敗: {e}")
        return None

    extracted = extract_text(src_path, max_chars=3000)
    image_path = extracted.get("image_path")
    has_text = bool(extracted["text"].strip())

    if not has_text and not image_path:
        reason = extracted["error"] or ("OCR必要" if extracted["needs_ocr"] else "本文なし")
        logging.info(f"LLM抽出スキップ: {original_name} ({reason})")
        return None

    # Anthropic API 呼び出し
    try:
        from anthropic import Anthropic
    except ImportError:
        logging.error("anthropic SDK 未インストール")
        return None

    client = Anthropic()
    tool_schema = {
        "name": "submit_classification",
        "description": "ファイル分類結果を提出する",
        "input_schema": {
            "type": "object",
            "properties": {
                "canonical_filename": {
                    "type": "string",
                    "description": "YYMMDD_書類名.拡張子 の形式",
                },
                "dest_relative_path": {
                    "type": "string",
                    "description": "01_事件記録/.../... または 03_経理/... 形式。不明なら空文字",
                },
                "accounting_category": {
                    "type": ["string", "null"],
                    "description": "経理書類のときの勘定科目。それ以外は null",
                },
                "vendor": {
                    "type": ["string", "null"],
                    "description": "経理書類のみ。業者名・取引先名。それ以外は null",
                },
                "amount": {
                    "type": ["integer", "null"],
                    "description": "経理書類のみ。税込合計金額（整数円）。それ以外は null",
                },
                "document_type": {
                    "type": ["string", "null"],
                    "enum": [None, "領収書", "請求書", "取引明細", "クレカ明細", "明細"],
                    "description": "経理書類のみ。書類種別。それ以外は null",
                },
                "requires_bengokakumei": {
                    "type": ["boolean", "null"],
                    "description": (
                        "裁判所への提出書面・裁判所からの通知・期日報告・相手方代理人との"
                        "やりとり書面など、弁護革命システムへの登録が望ましい裁判資料なら true。"
                        "委任契約書・領収書・打合せメモ・案件資料スキャンなどは false。"
                        "01_事件記録 以外に分類するファイルは null。"
                    ),
                },
                "reasoning": {
                    "type": "string",
                    "description": "判断根拠を一文で",
                },
                "confidence": {
                    "type": "string",
                    "enum": ["high", "medium", "low"],
                },
            },
            "required": [
                "canonical_filename",
                "dest_relative_path",
                "reasoning",
                "confidence",
            ],
        },
    }
    user_prompt_text = _build_llm_user_prompt(
        original_name, sender, subject, extracted["text"], clients,
        sheet_keyword_index=sheet_keyword_index,
    )

    if image_path:
        try:
            with open(image_path, "rb") as f:
                image_b64 = base64.standard_b64encode(f.read()).decode("ascii")
        except OSError as e:
            logging.error(f"画像読込失敗: {original_name}: {e}")
            return None
        message_content = [
            {
                "type": "image",
                "source": {
                    "type": "base64",
                    "media_type": "image/jpeg",
                    "data": image_b64,
                },
            },
            {"type": "text", "text": user_prompt_text},
        ]
    else:
        message_content = user_prompt_text

    try:
        response = client.messages.create(
            model=LLM_MODEL,
            max_tokens=LLM_MAX_TOKENS,
            system=[
                {
                    "type": "text",
                    "text": LLM_SYSTEM_RULES,
                    "cache_control": {"type": "ephemeral"},
                }
            ],
            tools=[tool_schema],
            tool_choice={"type": "tool", "name": "submit_classification"},
            messages=[{"role": "user", "content": message_content}],
        )
    except Exception as e:
        logging.error(f"LLM呼出失敗: {original_name}: {e}")
        return None

    # tool_use ブロックを取り出す
    tool_input: dict | None = None
    for block in response.content:
        if getattr(block, "type", None) == "tool_use":
            tool_input = block.input
            break
    if not tool_input:
        logging.warning(f"LLM tool_use 欠落: {original_name}")
        return None

    confidence = tool_input.get("confidence", "low")
    if confidence == "low":
        logging.info(f"LLM信頼度low: {original_name} reason={tool_input.get('reasoning')}")
        LLM_LOW_CONFIDENCE_LOG.append({
            "filename": original_name,
            "sender": sender,
            "subject": subject,
            "reasoning": tool_input.get("reasoning", ""),
        })
        return None

    # LLM返却名は検証してから採用（"/"・拡張子すり替え等はrename_fileへフォールバック）
    canonical_name = (
        _validate_llm_filename(tool_input.get("canonical_filename") or "", original_name)
        or rename_file(original_name)
    )
    rel = (tool_input.get("dest_relative_path") or "").strip().lstrip("/")
    reasoning = tool_input.get("reasoning") or ""
    if not rel:
        return None

    abs_dest = os.path.join(BASE_PATH, "共有用", rel)
    if not os.path.isdir(abs_dest):
        # 親まで遡って存在チェック（依頼者フォルダはあるがサブフォルダがない場合のフォールバック）
        parent = os.path.dirname(abs_dest)
        if os.path.isdir(parent):
            abs_dest = parent
            rel = os.path.relpath(abs_dest, os.path.join(BASE_PATH, "共有用"))
        else:
            logging.warning(f"LLM返却パス不在: {original_name} → {rel}")
            return None

    # 事件記録への誤割当ガード:
    # LLMが 01_事件記録/<依頼者> に振り分けたのに、その依頼者の氏名・別名が
    # 元ファイル名・送信者・件名・本文のどこにも現れない場合は、保存ログの
    # キーワードヒント過学習（活発な案件への引っ張られ）による誤爆とみなして拒否する。
    # 外形分類 classify_case_record と同じ「依頼者名が実在する」基準を LLM にも課す。
    if not _llm_case_dest_allowed(
        rel, clients, f"{original_name} {sender} {subject} {extracted['text']}"
    ):
        logging.warning(
            f"LLM事件記録誤割当ガード: {original_name} → {rel} "
            f"（依頼者名が本文・件名・送信者に不在のため拒否）"
        )
        return None

    # 経理配下なら 4 フィールド抽出結果でリネームを試みる
    if rel.startswith("03_経理/") or rel == "03_経理":
        accounting_category = tool_input.get("accounting_category")
        vendor = tool_input.get("vendor")
        amount = tool_input.get("amount")
        document_type = tool_input.get("document_type")
        if accounting_category and vendor and amount and document_type:
            original_ext = os.path.splitext(original_name)[1].lower()
            date_token = _extract_date_token(canonical_name, original_name)
            renamed = _build_accounting_filename(
                original_ext, vendor, amount, document_type,
                accounting_category, fallback_date=date_token,
            )
            if renamed:
                logging.info(f"経理リネーム適用: {original_name} → {renamed}")
                canonical_name = renamed

    extras = {
        "requires_bengokakumei": bool(tool_input.get("requires_bengokakumei")),
        "accounting_category": tool_input.get("accounting_category"),
        "vendor": tool_input.get("vendor"),
        "amount": tool_input.get("amount"),
        "document_type": tool_input.get("document_type"),
    }
    return canonical_name, abs_dest, rel, reasoning, extras


# ════════════════════════════════════════════════════════════════
# ファイル移動
# ════════════════════════════════════════════════════════════════

def rename_file(filename):
    """ファイル命名規則に従いリネームする。
    8桁日付(YYYYMMDD)で始まる場合は6桁(YYMMDD)に正規化、
    既に6桁日付_で始まる場合はそのまま、それ以外は今日の日付を頭に付ける。
    """
    m = re.match(r"^(\d{8})_(.+)", filename)
    if m:
        return f"{m.group(1)[2:]}_{m.group(2)}"
    if re.match(r"^\d{6}_", filename):
        return filename
    today = datetime.now().strftime("%y%m%d")
    return f"{today}_{filename}"


_AMBIGUOUS_STEM_RE = re.compile(
    r"""(?xi)
    ^(?:
        receipt[\-_.\d]*           # receipt, receipt_123
      | invoice[\-_.\d]*           # invoice, invoice_2026
      | statement[\-_.\d]*         # statement
      | bill[\-_.\d]*              # bill, bill_001
      | scan[\-_\d]*               # scan001, scan_20260612
      | scansnap[\-_.\d]*          # ScanSnap出力
      | img[\-_\d][\w\-_]*         # IMG_1234
      | image[\-_\d]*              # image001
      | dcim[\-_\d]*               # DCIM_1234
      | document[\-_\d]*           # document, document1
      | file[\-_\d]*               # file001
      | download[\-_\d]*           # download, download1
      | untitled[\-_\d]*           # untitled
      | noname[\-_\d]*             # noname
      | [\da-f]{8}-[\da-f]{4}-[\da-f]{4}-[\da-f]{4}-[\da-f]{12}  # UUID
      | [0-9a-f]{32,}              # MD5/SHA ハッシュ
      | \d+                        # 純粋な数字列
    )$
    """,
    re.IGNORECASE,
)


def _is_ambiguous_name(filename: str) -> bool:
    """ファイル名が内容を示さない汎用名かどうかを判定する。
    日付プレフィックス（YYMMDD_ or YYYYMMDD_）と拡張子を除いたステムで判定。
    """
    stem = os.path.splitext(filename)[0]
    stem = re.sub(r"^\d{6,8}[_\-]", "", stem)
    return bool(_AMBIGUOUS_STEM_RE.match(stem))


def _validate_llm_filename(llm_name: str, original_name: str) -> str | None:
    """LLMが返したファイル名の最小バリデーション。問題があれば None を返す。"""
    if not llm_name:
        return None
    if "/" in llm_name or ".." in llm_name:
        return None
    if re.search(r"[\x00:\\]", llm_name):
        return None
    if len(llm_name.encode("utf-8")) > 255:
        return None
    orig_ext = os.path.splitext(original_name)[1].lower()
    if orig_ext and not llm_name.lower().endswith(orig_ext):
        return None
    return llm_name


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

def fetch_sheet_data(sheets_service, retries=3, delay=5.0):
    """保存ログシートの全データを取得する（リトライ付き）。
    Returns:
        list[list[str]]: 全行データ（ヘッダー含む）
    """
    for attempt in range(retries):
        try:
            result = sheets_service.spreadsheets().values().get(
                spreadsheetId=SPREADSHEET_ID,
                range=f"{SHEET_NAME}!A:L",
            ).execute()
            return result.get("values", [])
        except Exception as e:
            if attempt < retries - 1:
                logging.warning(f"Sheets API一時障害（{attempt+1}/{retries}）: {e}")
                time.sleep(delay)
            else:
                raise


def build_filename_index(sheet_data):
    """シートデータからD列（ファイル名）→ 行番号（1-based）のマップを作る。
    同時にB列（メールタイトル）・C列（送信者）・E列（移動先）の値も保持する。
    Returns:
        dict: {filename: {"row": int, "subject": str, "sender": str, "dest": str}}
    """
    index = {}
    for i, row in enumerate(sheet_data):
        if i == 0:
            continue  # ヘッダーをスキップ
        if len(row) >= 4 and row[3]:
            subject = row[1] if len(row) >= 2 else ""
            sender = row[2] if len(row) >= 3 else ""
            dest = row[4] if len(row) >= 5 else ""
            index[row[3]] = {
                "row": i + 1,  # 1-based行番号
                "subject": subject,
                "sender": sender,
                "dest": dest,
            }
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


def highlight_freee_cell(sheets_service, row_number):
    """G列（freee）セルの背景を黄色にする（チェックは入れない）。"""
    sheets_service.spreadsheets().batchUpdate(
        spreadsheetId=SPREADSHEET_ID,
        body={"requests": [{
            "repeatCell": {
                "range": {
                    "sheetId": SHEET_GID,
                    "startRowIndex": row_number - 1,
                    "endRowIndex": row_number,
                    "startColumnIndex": 6,
                    "endColumnIndex": 7,
                },
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": {
                            "red": 1.0, "green": 1.0, "blue": 0.0, "alpha": 1.0,
                        }
                    }
                },
                "fields": "userEnteredFormat.backgroundColor",
            }
        }]},
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

        # 移動前に元位置でDriveメタを取得しておく。IDはmove後も不変。
        cur_folder_id = resolve_drive_folder_id(drive_service, current_dir)
        pre_meta = (
            fetch_drive_file_meta(drive_service, cur_folder_id, filename)
            if cur_folder_id else None
        )

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
        if pre_meta and pre_meta.get("webViewLink"):
            url = pre_meta["webViewLink"]
        elif pre_meta and pre_meta.get("id"):
            url = build_drive_link_from_id(pre_meta["id"])
        else:
            # フォールバック: 既存D列の式から ID を抽出してみる
            existing_id = _fetch_d_column_id(sheets_service, i + 1)
            if existing_id:
                url = build_drive_link_from_id(existing_id)
            else:
                logging.warning(f"修正後link取得失敗: {filename}")
                url = drive_search_url(actual_name)
        d_cell = hyperlink_formula(url, actual_name)
        try:
            update_sheet_row(sheets_service, i + 1, d_cell, target_desc)
            set_done_checkbox(sheets_service, i + 1, True)
        except Exception as e:
            logging.error(f"修正後シート更新失敗: {filename} row={i+1}: {e}")
            print(f"  修正反映済みだがシート更新失敗: {filename} ({e})")
        corrected += 1

    if corrected:
        print(f"修正指示反映: {corrected}件")
    return corrected


def extract_learning_proposals(sheet_data, clients) -> list[dict]:
    """保存ログのJ列修正履歴から、定数追加の提案を抽出する。

    修正済み行（K=TRUE かつ J列に値あり）を走査し、
    同一送信者メアドが3回以上同じ依頼者フォルダに修正されていたら提案。
    """
    from collections import defaultdict
    sender_to_client: dict[str, dict[str, list[int]]] = defaultdict(
        lambda: defaultdict(list)
    )

    for i, row in enumerate(sheet_data):
        if i == 0:
            continue
        def _c(idx):
            return str(row[idx]).strip() if len(row) > idx and row[idx] else ""
        if _c(10).upper() != "TRUE":
            continue
        if not _c(9):
            continue
        sender_raw = _c(2)
        dest = _c(4)
        if not sender_raw or not dest:
            continue
        if not dest.startswith("01_事件記録"):
            continue
        emails = re.findall(r'[\w.+-]+@[\w.-]+', sender_raw.lower())
        client_folder = dest.split("/")[1] if "/" in dest else dest
        for em in emails:
            sender_to_client[em][client_folder].append(i + 1)

    proposals = []
    existing_emails = {k.lower() for k in SENDER_EMAIL_TO_CLIENT}
    for email, folders in sender_to_client.items():
        if email in existing_emails:
            continue
        for folder, rows in folders.items():
            if len(rows) >= 3:
                proposals.append({
                    "type": "sender_email",
                    "email": email,
                    "client_folder": folder,
                    "count": len(rows),
                    "evidence_rows": rows[:5],
                })

    return proposals


def compute_classification_stats(sheet_data) -> dict:
    """直近30日の保存ログから分類精度の統計を算出する。"""
    from datetime import datetime as _dt, timedelta
    cutoff = _dt.now() - timedelta(days=30)
    confirmed = 0
    corrected = 0
    pending_review = 0

    for i, row in enumerate(sheet_data):
        if i == 0:
            continue
        def _c(idx):
            return str(row[idx]).strip() if len(row) > idx and row[idx] else ""
        saved_at = _c(0)
        dt = None
        for fmt in ("%Y/%m/%d %H:%M:%S", "%Y-%m-%d %H:%M:%S",
                     "%Y/%m/%d %H:%M", "%Y-%m-%d %H:%M"):
            try:
                dt = _dt.strptime(saved_at, fmt)
                break
            except ValueError:
                continue
        if not dt or dt < cutoff:
            continue
        dest = _c(4)
        if not dest or "分からなかった" in dest or "pre_trash" in dest:
            continue
        has_correction = bool(_c(9))
        is_confirmed = _c(5).upper() == "TRUE"
        if has_correction:
            corrected += 1
        elif is_confirmed:
            confirmed += 1
        else:
            pending_review += 1

    total_judged = confirmed + corrected
    rate = (confirmed / total_judged * 100) if total_judged > 0 else 0
    return {
        "confirmed": confirmed,
        "corrected": corrected,
        "pending_review": pending_review,
        "accuracy_rate": rate,
    }


# ════════════════════════════════════════════════════════════════
# Gmail通知
# ════════════════════════════════════════════════════════════════

_KOMETANI_NAME_PATTERNS = (
    "米谷尚起", "米谷 尚起", "米谷　尚起", "コメタニ ナオキ", "コメタニナオキ",
    "Naoki Kometani", "kometani", "N.Kometani", "n.kometani",
)

_RESTART_FIRM_PATTERNS = (
    "弁護士法人Re-Start法律事務所", "弁護士法人Re-Start", "弁護士法人リスタート",
    "Re-Start法律事務所", "Re-Start", "リスタート法律事務所", "re-startlaw",
)

# 03_経理 配下でも freee 登録判定の対象外とするサブツリー
_FREEE_EXCLUDE_SEGMENTS = ("預り金", "事業主貸", "預け金", "04_会計勉強")

# ファイル名先頭の日付プレフィックス（YYYYMMDD または YYMMDD）→ ISO 日付
_FILENAME_DATE_RE = re.compile(r"^(20\d{2}|2[6-9])(\d{2})(\d{2})[_\.\-]")


def _guess_issue_date(filename: str, fallback_path: str = "") -> str:
    """ファイル名先頭の日付プレフィックスを ISO YYYY-MM-DD に変換する。
    取れない場合はファイル更新日 → 当日 にフォールバック。
    """
    m = _FILENAME_DATE_RE.match(filename)
    if m:
        y, mo, d = m.group(1), m.group(2), m.group(3)
        if len(y) == 2:
            y = "20" + y
        return f"{y}-{mo}-{d}"
    if fallback_path and os.path.isfile(fallback_path):
        ts = datetime.fromtimestamp(os.path.getmtime(fallback_path))
        return ts.strftime("%Y-%m-%d")
    return datetime.now().strftime("%Y-%m-%d")


_ACCOUNTING_DOC_KEYWORDS = ("請求書", "領収書", "領収証", "納品書", "見積書", "レシート")


def _detect_freee_target(filename: str, dest: str, src_path: str) -> str:
    """経理処理が必要な書類（freee登録対象）かを判定する。

    判定（OR条件、いずれか該当で対象）:
      1. 03_経理/ 配下に分類 + 本文に米谷尚起 または Re-Start法律事務所名義
      2. 03_経理/ 配下に分類 + ファイル名に請求書・領収書等の経理書類キーワード
    除外: 預り金・事業主貸・預け金・会計勉強 サブツリー。
    Returns: 該当時は理由メモ、非該当時は空文字。
    """
    if not dest or not dest.startswith("03_経理"):
        return ""
    if any(seg in dest for seg in _FREEE_EXCLUDE_SEGMENTS):
        return ""

    for kw in _ACCOUNTING_DOC_KEYWORDS:
        if kw in filename:
            return f"経理書類（ファイル名に「{kw}」を含む・03_経理/配下）"

    try:
        sys.path.insert(0, os.path.join(os.path.dirname(__file__), "scripts"))
        from extract_text import extract_text as _et  # type: ignore
        text = _et(src_path, max_chars=3000).get("text", "")
    except Exception as e:
        logging.warning(f"freee判定の本文抽出失敗: {filename}: {e}")
        return ""
    if not text:
        return ""
    for kw in _KOMETANI_NAME_PATTERNS:
        if kw in text:
            return f"米谷尚起 名義が本文に登場（宛先または発行者・「{kw}」検出）"
    for kw in _RESTART_FIRM_PATTERNS:
        if kw in text:
            return f"Re-Start法律事務所 名義が本文に登場（「{kw}」検出）"
    return ""


# 裁判系サブフォルダ（弁護革命登録対象の判定に使う）
_COURT_SUBFOLDERS = {
    # 民事案件サブフォルダ
    "00主張", "01甲号証", "02乙号証", "03連絡文書", "04事務", "05期日報告書",
    # 刑事案件サブフォルダ
    "02_身体拘束関係", "03_検察官提出書面", "04_弁護人提出書面",
    "05_検察官証拠", "06_裁判所手続", "11_弁護人請求証拠",
}


def _detect_bengokakumei_target(
    filename: str,
    dest: str,
    llm_flag: bool = False,
) -> bool:
    """裁判資料（弁護革命登録対象）かを判定する。

    条件: 以下のいずれか
      1. 01_事件記録/ 配下に分類され、パス内に裁判系サブフォルダ名が含まれる
      2. LLM が requires_bengokakumei=true を返した（フォルダ判定の補助）
    """
    if llm_flag:
        return True
    if not dest or not dest.startswith("01_事件記録"):
        return False
    segments = dest.split("/")
    return any(seg in _COURT_SUBFOLDERS for seg in segments)


def _collect_pending_files(sheet_data) -> list[dict]:
    """保存ログから「分からなかった」滞留行を集める。

    実フォルダ存在チェック付き: シートのE列が「分からなかった」でも
    実際にファイルがそのフォルダに存在しない場合はゴースト行として除外する。
    （同名ファイルの重複行が build_filename_index で上書きされ、
      古い行のE列が更新されないまま残るのを救済する）

    Returns: [{"row","filename","saved_at","days","subject","sender","link"}, ...]
    """
    from datetime import datetime as _dt

    unknown_dir = os.path.join(BASE_PATH, "共有用", "06_分類依頼", "分からなかった")
    try:
        actual_files = set(
            n for n in os.listdir(unknown_dir) if not n.startswith(".")
        ) if os.path.isdir(unknown_dir) else set()
    except OSError:
        actual_files = set()

    out: list[dict] = []
    today = _dt.now()
    for i, row in enumerate(sheet_data):
        if i == 0:
            continue
        def _c(idx):
            if len(row) <= idx or row[idx] is None:
                return ""
            return str(row[idx]).strip()
        dest = _c(4)
        if not dest.startswith("06_分類依頼/分からなかった"):
            continue
        if _c(5).upper() == "TRUE":  # F列 確認済
            continue
        if _c(10).upper() == "TRUE":  # K列 既に済
            continue

        d_cell = _c(3)
        link = ""
        m = re.search(r'HYPERLINK\("([^"]+)"', d_cell)
        if m:
            link = m.group(1)
        filename = d_cell
        m2 = re.search(r'HYPERLINK\("[^"]+","([^"]+)"\)', d_cell)
        if m2:
            filename = m2.group(1)
        # FORMATTED_VALUE 取得ではD列がラベル文字列になりHYPERLINK式を持たないため、
        # リンクが取れないときはDrive内検索URLで代替する（報告メールに必ずリンクを載せる）
        if not link and filename:
            link = drive_search_url(filename)

        # ゴースト行除外: 実フォルダに無いならスキップ
        if filename not in actual_files:
            logging.info(
                f"ゴースト行スキップ: row{i+1} filename={filename!r} "
                f"（E列は分からなかっただが実フォルダに不在）"
            )
            continue

        saved_at = _c(0)
        days = ""
        for fmt in ("%Y/%m/%d %H:%M:%S", "%Y-%m-%d %H:%M:%S"):
            try:
                dt = _dt.strptime(saved_at, fmt)
                days = str((today - dt).days)
                break
            except ValueError:
                continue
        out.append({
            "row": i + 1,  # 1-based 行番号
            "filename": filename,
            "saved_at": saved_at,
            "days": days,
            "subject": _c(1),
            "sender": _c(2),
            "link": link,
        })
    return out


def send_notification(
    gmail_service,
    file_results,
    pending_files=None,
    freee_action_items=None,
    bengokakumei_action_items=None,
    move_failures=None,
    maybe_accounting_in_unknown=None,
    learning_proposals=None,
    classification_stats=None,
    llm_low_confidence=None,
):
    """佐藤信子・米谷尚起へ分類完了通知メールを送信する。"""
    body_lines = [
        "佐藤様　米谷様",
        "",
        "お疲れ様です。",
        "ファイル分類が完了しましたのでご報告いたします。",
        "",
        "【分類結果】",
    ]
    if file_results:
        for fname, dest, drive_link in file_results:
            body_lines.append(f"  ・{fname} → {dest}")
            if drive_link:
                body_lines.append(f"    Drive: {drive_link}")
    else:
        body_lines.append("  （本日の分類対象ファイルはありませんでした）")

    # 移動失敗セクション
    if move_failures:
        body_lines.append("")
        body_lines.append(f"【!!! 移動失敗 {len(move_failures)}件】")
        for fname, dest, err in move_failures:
            body_lines.append(f"  ・{fname} → {dest}")
            body_lines.append(f"    エラー: {err}")

    # 佐藤さんへのアクション依頼セクション
    body_lines.append("")
    body_lines.append("───────────────────────────")
    body_lines.append("【佐藤様 本日のお願い】")
    body_lines.append("───────────────────────────")

    body_lines.append("")
    if freee_action_items:
        uploaded = [it for it in freee_action_items if it[3]]
        failed = [it for it in freee_action_items if not it[3]]
        body_lines.append(
            f"■ freeeファイルボックス登録（米谷尚起／Re-Start法律事務所名義の経理書類 "
            f"{len(freee_action_items)}件・成功{len(uploaded)}件／失敗{len(failed)}件）"
        )
        if uploaded:
            body_lines.append(
                "以下のファイルはfreeeファイルボックスにアップロード済みです。"
                "freee上での仕訳登録（勘定科目・取引先・金額入力）をお願いいたします。"
            )
        for fname, dest, note, receipt_id, receipt_number, ui_url, err in freee_action_items:
            body_lines.append(f"  ・{fname}")
            body_lines.append(f"    保存先: {dest}")
            if note:
                body_lines.append(f"    判定: {note}")
            if receipt_id:
                num_display = str(receipt_number) if receipt_number else str(receipt_id)
                if ui_url:
                    body_lines.append(f"    freeeファイルボックス No.{num_display}: {ui_url}")
                else:
                    body_lines.append(f"    freeeファイルボックス No.{num_display}")
            elif err:
                body_lines.append(f"    アップロード失敗: {err}（手動アップロードをお願いします）")
    else:
        body_lines.append(
            "■ freeeファイルボックス登録: 本日は米谷尚起／Re-Start法律事務所名義の経理書類はありませんでした。"
        )

    body_lines.append("")
    if bengokakumei_action_items:
        body_lines.append(f"■ 弁護革命登録（裁判資料 {len(bengokakumei_action_items)}件）")
        body_lines.append("以下の裁判資料は弁護革命システムへのご登録をお願いいたします。")
        for fname, dest in bengokakumei_action_items:
            body_lines.append(f"  ・{fname}")
            body_lines.append(f"    保存先: {dest}")
    else:
        body_lines.append("■ 弁護革命登録: 本日は弁護革命登録対象の裁判資料はありませんでした。")

    # 分からなかったの経理書類候補（Phase 2C）
    if maybe_accounting_in_unknown:
        body_lines.append("")
        body_lines.append(f"■ 注意: 以下のファイルは分類先不明ですが、経理書類の可能性があります。")
        body_lines.append("  米谷弁護士の確認後、J列に修正指示をお願いいたします。")
        for fname in maybe_accounting_in_unknown:
            body_lines.append(f"  ・{fname}")

    # 滞留ファイル（Phase 2D: 7日超エスカレーション付き）
    has_aged_pending = False
    if pending_files:
        body_lines.append("")
        body_lines.append("───────────────────────────")
        body_lines.append(f"【米谷確認待ち（分からなかった滞留 {len(pending_files)}件）】")
        body_lines.append("───────────────────────────")
        body_lines.append("対応方法: 保存ログの該当行のJ列に「分類」または具体的な移動先を記入すると、")
        body_lines.append("         翌朝の自動分類で apply_corrections が反映します。")
        body_lines.append("")
        for p in pending_files:
            row_str = f"行{p['row']}" if p.get('row') else ""
            days_int = int(p['days']) if p['days'] else 0
            if days_int > 7:
                has_aged_pending = True
                days_str = f"!!!{days_int}日超滞留!!!"
            else:
                days_str = f"{p['days']}日経過" if p['days'] else ""
            body_lines.append(f"  ・{p['filename']}")
            meta_bits = [b for b in (
                row_str,
                days_str,
                f"送信者: {p['sender']}" if p['sender'] else "",
            ) if b]
            if meta_bits:
                body_lines.append(f"     {' / '.join(meta_bits)}")
            if p['subject']:
                body_lines.append(f"     件名: {p['subject'][:80]}")
            if p['link']:
                body_lines.append(f"     {p['link']}")

    # Phase 3: LLM低信頼度
    if llm_low_confidence:
        body_lines.append("")
        body_lines.append(f"【LLMでも分類困難だったファイル ({len(llm_low_confidence)}件)】")
        for item in llm_low_confidence:
            body_lines.append(f"  ・{item['filename']}")
            if item.get('reasoning'):
                body_lines.append(f"    理由: {item['reasoning'][:80]}")

    # Phase 3: 学習提案
    if learning_proposals:
        body_lines.append("")
        body_lines.append("───────────────────────────")
        body_lines.append(f"【学習提案 ({len(learning_proposals)}件)】")
        body_lines.append("───────────────────────────")
        body_lines.append("以下のパターンが修正履歴から検出されました。")
        body_lines.append("コードに反映すると今後の自動分類精度が向上します。")
        for prop in learning_proposals:
            body_lines.append(
                f"  ・送信者 {prop['email']} → {prop['client_folder']} "
                f"（修正{prop['count']}回・行{prop['evidence_rows'][:3]}）"
            )

    # Phase 3: 分類精度統計
    if classification_stats and (classification_stats.get("confirmed", 0) + classification_stats.get("corrected", 0)) > 0:
        body_lines.append("")
        body_lines.append(
            f"分類精度（直近30日）: 正解率{classification_stats['accuracy_rate']:.0f}% "
            f"(確認済{classification_stats['confirmed']}件 / "
            f"修正{classification_stats['corrected']}件 / "
            f"未確認{classification_stats['pending_review']}件)"
        )

    body_lines.append("")
    body_lines.append("ご確認のほどよろしくお願いいたします。")
    body_lines.append(EMAIL_SIGNATURE)

    subject = NOTIFY_SUBJECT
    if has_aged_pending:
        subject = f"[要確認] {NOTIFY_SUBJECT}"

    message = MIMEText("\n".join(body_lines), "plain", "utf-8")
    message["to"] = ", ".join(NOTIFY_TO)
    message["from"] = SENDER
    message["subject"] = subject

    raw = base64.urlsafe_b64encode(message.as_bytes()).decode("utf-8")
    gmail_service.users().messages().send(
        userId="me", body={"raw": raw}
    ).execute()


# ════════════════════════════════════════════════════════════════
# メイン処理
# ════════════════════════════════════════════════════════════════

def scan_watch_folder():
    """監視フォルダ直下と「スキャナーから」配下のファイル一覧を返す。
    スキャナーからは再帰走査（日付サブフォルダ対応）。
    除外対象：隠しファイル、pre_trash・分からなかった配下。
    """
    files = []
    if not os.path.isdir(WATCH_FOLDER):
        print(f"監視フォルダが見つかりません: {WATCH_FOLDER}")
        return files

    exclude_dirs = {
        os.path.basename(PRE_TRASH),
        os.path.basename(UNKNOWN_FOLDER),
        os.path.basename(SCANSNAP_FOLDER),
    }
    for name in os.listdir(WATCH_FOLDER):
        if name.startswith(".") or name in exclude_dirs:
            continue
        full = os.path.join(WATCH_FOLDER, name)
        if os.path.isfile(full):
            files.append((name, full))

    if os.path.isdir(SCANSNAP_FOLDER):
        for dirpath, dirnames, filenames in os.walk(SCANSNAP_FOLDER):
            dirnames[:] = [d for d in dirnames if not d.startswith(".")]
            for name in filenames:
                if name.startswith("."):
                    continue
                files.append((name, os.path.join(dirpath, name)))

    return files


def _acquire_single_instance_lock():
    """二重起動防止のプロセス排他ロックを取得する。

    260701 に secretary.py が2プロセス並走し、同一ファイルへのLLM二重課金・
    片方が移動済みのファイルでもう片方が「移動失敗」を量産する事故が発生した。
    取得できなければ None（既に別プロセスが実行中）。
    """
    import fcntl
    f = open(LOCK_PATH, "w")
    try:
        fcntl.flock(f, fcntl.LOCK_EX | fcntl.LOCK_NB)
    except OSError:
        f.close()
        return None
    f.write(str(os.getpid()))
    f.flush()
    return f


def main():
    lock = _acquire_single_instance_lock()
    if lock is None:
        msg = "別の secretary.py が実行中のため終了します（二重起動防止）。"
        logging.warning(msg)
        print(msg)
        return

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

    # 保存ログからメアド→依頼者マップを自動構築（既存固定辞書を補強）
    learned_index, learn_stats = build_sender_email_index(sheet_data, clients)
    SENDER_EMAIL_TO_CLIENT.update(learned_index)
    # 手動固定エントリを後勝ちで再適用（学習が固定値を上書きしないように）
    SENDER_EMAIL_TO_CLIENT.update(_MANUAL_SENDER_EMAIL_TO_CLIENT)
    print(
        f"送信者メアド学習: {learn_stats['emails_indexed']}件追加 "
        f"(走査{learn_stats['rows_scanned']}行 / "
        f"曖昧除外{learn_stats['emails_ambiguous']}件)"
    )

    # 保存ログから (送信者キー, 件名キー) → 共有用配下の保存先 を学習する。
    # メアド単独で覚えるのではなく、件名と組み合わせて学習することで
    # 同じ送信者から異なる種類の書類が来ても正しく振り分けられる。
    routing, route_stats = build_sender_subject_routing(sheet_data)
    LEARNED_ROUTING.clear()
    LEARNED_ROUTING.update(routing)
    print(
        f"送信者+件名ルーティング学習: {route_stats['learned']}件 "
        f"(矛盾除外{route_stats['ambiguous']}件 / 対象外{route_stats['skipped_rows']}行)"
    )

    # 保存ログ全行を「固有名詞 → 過去 dest 集計」用にインデックス化（LLM 文脈強化）
    sheet_keyword_index = build_keyword_dest_index(sheet_data)

    # 監視フォルダスキャン
    files = scan_watch_folder()
    if not files:
        print("分類対象ファイルはありません。")
        # 滞留ファイルがあれば通知は送る
        pending_files = []
        try:
            pending_files = _collect_pending_files(sheet_data)
            if pending_files:
                print(f"米谷確認待ち: {len(pending_files)}件（分からなかった滞留）")
        except Exception as e:
            logging.warning(f"確認待ち集計に失敗: {e}")
        if pending_files:
            try:
                send_notification(gmail_service, [], pending_files=pending_files)
                print(f"滞留通知メールを送信しました: {', '.join(NOTIFY_TO)}")
            except Exception as e:
                logging.critical(f"滞留通知メール送信失敗: {e}")
        return

    print(f"対象ファイル: {len(files)}件")

    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    pending_new = []   # (original_name, actual_name, d_cell, dest_description, l_note)
    updated_rows = []
    file_results = []
    freee_action_items = []  # (filename, dest, note)
    freee_uploaded_names = set()
    bengokakumei_action_items = []  # (filename, dest)
    move_failures = []  # (filename, dest, error_message)
    maybe_accounting_in_unknown = []  # (filename,) 分からなかったに入った経理書類候補
    skipped = 0
    vanished = 0
    duplicate_count = 0
    llm_used = 0

    for original_name, src_path in files:
        # スプレッドシートに既に存在し、E列（移動先）が埋まっている → 処理済みとしてスキップ
        if original_name in filename_index and filename_index[original_name]["dest"]:
            skipped += 1
            continue

        # スキャン後にDrive同期・手動操作で消えたファイルは静かにスキップ
        # （LLM課金・Driveメタ取得の前に判定し、移動失敗として警報も出さない）
        if not os.path.exists(src_path):
            logging.info(f"消失スキップ: {original_name}（スキャン後にフォルダから消えた）")
            print(f"  消失スキップ: {original_name}")
            vanished += 1
            continue

        # スプレッドシートから送信者・件名を取り出して分類ヒントに使う
        meta = filename_index.get(original_name, {})
        sender = meta.get("sender", "")
        subject = meta.get("subject", "")

        # 移動「前」にDriveメタを取得しておく。move後はDrive側のインデックス
        # 同期遅延でファイルが見つからないことが多いため。IDはmove後も不変。
        # スキャン直後のファイルはDrive未同期が常態のため、リトライは1回に抑える
        # （retries=3だと未同期ファイル1件あたり約6秒浪費。取れなければ検索リンクで代替）
        src_dir = os.path.dirname(src_path)
        src_folder_id = resolve_drive_folder_id(drive_service, src_dir)
        pre_meta = (
            fetch_drive_file_meta(drive_service, src_folder_id, original_name, retries=1)
            if src_folder_id else None
        )
        if not pre_meta:
            logging.warning(f"Drive未同期: {original_name}（移動前メタ取得失敗）")

        canonical_name = rename_file(original_name)
        dest_dir, dest_description = classify_file(
            canonical_name, clients, sender=sender, subject=subject
        )

        # 外形分類で UNKNOWN に落ちたら LLM フォールバック（分類もリネームも LLM が決定）
        # 汎用ファイル名（receipt, scan 等）は分類成功でも LLM でリネームのみ実施
        l_note = ""
        llm_extras: dict = {}
        if dest_dir == UNKNOWN_FOLDER:
            llm_result = classify_by_content(
                src_path, original_name, sender, subject, clients,
                sheet_keyword_index=sheet_keyword_index,
            )
            if llm_result:
                canonical_name, dest_dir, dest_description, reasoning, llm_extras = llm_result
                l_note = f"LLM推定: {reasoning}"
                llm_used += 1
                print(f"  [LLM] {original_name} → {dest_description} ({reasoning})")
        elif _is_ambiguous_name(original_name):
            llm_result = classify_by_content(
                src_path, original_name, sender, subject, clients,
                sheet_keyword_index=sheet_keyword_index,
            )
            if llm_result:
                llm_name, _llm_dir, _llm_desc, reasoning, new_extras = llm_result
                validated = _validate_llm_filename(llm_name, original_name)
                if validated:
                    canonical_name = validated
                    l_note = f"LLMリネーム: {reasoning}"
                llm_extras = new_extras
                llm_used += 1
                print(f"  [LLMリネーム] {original_name} → {canonical_name}")
        new_name = canonical_name

        # 佐藤さんへのアクション判定（移動前にsrc_pathで本文をチェック）
        needs_freee_note = _detect_freee_target(canonical_name, dest_description, src_path)
        needs_bengokakumei = _detect_bengokakumei_target(
            canonical_name,
            dest_description,
            llm_flag=bool(llm_extras.get("requires_bengokakumei")),
        )

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

        # Phase 2C: 分からなかったに入った経理書類候補を記録
        if dest_dir == UNKNOWN_FOLDER:
            for kw in _ACCOUNTING_DOC_KEYWORDS:
                if kw in canonical_name:
                    maybe_accounting_in_unknown.append(canonical_name)
                    break

        # ファイル移動〜シート更新〜freeeアップロード
        try:
            moved_path = move_file(src_path, dest_dir, new_name)
        except Exception as e:
            logging.error(f"ファイル移動失敗: {original_name} → {dest_description}: {e}")
            print(f"  !!! 移動失敗: {original_name} → {dest_description}: {e}")
            move_failures.append((original_name, dest_description, str(e)))
            continue
        actual_name = os.path.basename(moved_path)

        if pre_meta and pre_meta.get("webViewLink"):
            link_url = pre_meta["webViewLink"]
            link_kind = "[直リンク]"
        elif pre_meta and pre_meta.get("id"):
            link_url = build_drive_link_from_id(pre_meta["id"])
            link_kind = "[直リンク:ID]"
        else:
            link_url = drive_search_url(actual_name)
            link_kind = "[検索リンク・要確認]"
        d_cell = hyperlink_formula(link_url, actual_name)

        print(f"  {original_name} → {dest_description}/{actual_name} {link_kind}")

        # スプレッドシートにファイル名が既に存在する場合 → D列（リンク付き）とE列を更新
        existing = filename_index.get(original_name) or filename_index.get(actual_name)
        if existing:
            update_sheet_row(sheets_service, existing["row"], d_cell, dest_description)
            updated_rows.append(actual_name)
        else:
            pending_new.append(
                (original_name, actual_name, d_cell, dest_description, l_note)
            )

        file_results.append((actual_name, dest_description, link_url))
        if needs_freee_note:
            receipt_id = None
            receipt_number = None
            ui_url = None
            upload_error = None
            if freee_filebox is not None:
                try:
                    issue_date = _guess_issue_date(actual_name, moved_path)
                    r = freee_filebox.upload_receipt(
                        moved_path,
                        issue_date=issue_date,
                        description=actual_name,
                        business_type="corporate",
                    )
                    receipt_id = r.get("id")
                    receipt_number = r.get("receipt_number")
                    ui_url = r.get("ui_url")
                    logging.info(
                        f"freeeファイルボックス登録: {actual_name} → "
                        f"receipt_id={receipt_id}"
                    )
                    print(f"    freeeアップロード成功 receipt_id={receipt_id}")
                    freee_uploaded_names.add(actual_name)
                except Exception as e:
                    upload_error = str(e)
                    logging.warning(
                        f"freeeファイルボックス登録失敗: {actual_name}: {e}"
                    )
                    print(f"    freeeアップロード失敗: {e}")
            else:
                upload_error = "freeeモジュール未読込（import失敗）"
                logging.warning(f"freee_filebox未読込: {actual_name}（手動アップロード必要）")
            freee_action_items.append(
                (actual_name, dest_description, needs_freee_note,
                 receipt_id, receipt_number, ui_url, upload_error)
            )
        if needs_bengokakumei:
            bengokakumei_action_items.append((actual_name, dest_description))

    if skipped:
        print(f"処理済みスキップ: {skipped}件")
    if vanished:
        print(f"消失スキップ: {vanished}件")
    if duplicate_count:
        print(f"重複検出: {duplicate_count}件（pre_trashへ退避）")
    if move_failures:
        print(f"移動失敗: {len(move_failures)}件")
    if llm_used:
        print(f"LLMフォールバック: {llm_used}件")

    # 1F: 処理件数照合
    processed_total = (
        len(file_results) + skipped + vanished + duplicate_count + len(move_failures)
    )
    if processed_total != len(files):
        logging.warning(
            f"処理件数不一致: 対象{len(files)}件 vs "
            f"分類{len(file_results)}+スキップ{skipped}+消失{vanished}"
            f"+重複{duplicate_count}+失敗{len(move_failures)}={processed_total}件"
        )
        print(f"  ⚠ 処理件数不一致: 対象{len(files)} vs 処理済{processed_total}")

    # append 直前にシートを再フェッチして GAS とのレース重複を防ぐ
    if pending_new:
        fresh_data = fetch_sheet_data(sheets_service)
        fresh_index = build_filename_index(fresh_data)
        truly_new = []
        for original_name, actual_name, d_cell, dest_description, l_note in pending_new:
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
                    False, False, False, False, "", False, l_note,
                ])

        if truly_new:
            append_to_sheet(sheets_service, truly_new)
            print(f"スプレッドシートに{len(truly_new)}行追加しました。")

    if updated_rows:
        print(f"スプレッドシートの{len(updated_rows)}行を更新しました。")

    # 「分からなかった」滞留ファイル一覧をシートから再取得（移動・追記後の最新状態で）
    pending_files = []
    latest_sheet = None
    try:
        latest_sheet = fetch_sheet_data(sheets_service)
        pending_files = _collect_pending_files(latest_sheet)
        if pending_files:
            print(f"米谷確認待ち: {len(pending_files)}件（分からなかった滞留）")
    except Exception as e:
        logging.warning(f"確認待ち集計に失敗: {e}")

    # freeeアップロード成功行のG列を黄色ハイライト
    if freee_uploaded_names and latest_sheet:
        latest_index = build_filename_index(latest_sheet)
        for fname in freee_uploaded_names:
            row_info = latest_index.get(fname)
            if row_info:
                try:
                    highlight_freee_cell(sheets_service, row_info["row"])
                except Exception as e:
                    logging.warning(f"G列ハイライト失敗: {fname} row={row_info['row']}: {e}")

    # Phase 3: 学習提案・精度統計
    learning_proposals = []
    classification_stats = {}
    try:
        data_for_stats = latest_sheet if latest_sheet else sheet_data
        learning_proposals = extract_learning_proposals(data_for_stats, clients)
        classification_stats = compute_classification_stats(data_for_stats)
        if learning_proposals:
            print(f"学習提案: {len(learning_proposals)}件")
        if classification_stats:
            print(
                f"分類精度（直近30日）: 正解率{classification_stats['accuracy_rate']:.0f}% "
                f"(確認済{classification_stats['confirmed']}件 / "
                f"修正{classification_stats['corrected']}件)"
            )
    except Exception as e:
        logging.warning(f"学習分析に失敗: {e}")

    # 通知メール送信
    if file_results or pending_files or move_failures:
        try:
            send_notification(
                gmail_service,
                file_results,
                pending_files=pending_files,
                freee_action_items=freee_action_items,
                bengokakumei_action_items=bengokakumei_action_items,
                move_failures=move_failures,
                maybe_accounting_in_unknown=maybe_accounting_in_unknown,
                learning_proposals=learning_proposals,
                classification_stats=classification_stats,
                llm_low_confidence=LLM_LOW_CONFIDENCE_LOG,
            )
            print(f"通知メールを送信しました: {', '.join(NOTIFY_TO)}")
            if freee_action_items:
                print(f"  freee登録依頼: {len(freee_action_items)}件")
            if bengokakumei_action_items:
                print(f"  弁護革命登録依頼: {len(bengokakumei_action_items)}件")
        except Exception as e:
            logging.critical(f"通知メール送信失敗: {e}")
            print(f"  !!! 通知メール送信失敗: {e}")
            print(f"  !!! 分類処理は完了済。手動でメール報告が必要です。")

    print("=== 分類完了 ===")


if __name__ == "__main__":
    main()
