"""
freee会計API連携モジュール
- 機能1: PDFレシート・請求書の自動取引登録
- 機能2: ファイルボックス未連携ファイルの取引紐付け
- 処理結果をメールで報告（修正指示しやすい番号付き一覧）
"""

import base64
import logging
import os
import re
import shutil
import time
from datetime import datetime
from email.mime.text import MIMEText

import pdfplumber
import requests
from google.auth.transport.requests import Request as GoogleRequest
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

import freee_auth

# ── パス設定 ──
BASE_PATH = (
    "/Users/kometaninaoki/Library/CloudStorage/"
    "GoogleDrive-n.kometani@re-startlaw.com/マイドライブ/"
)
FREEE_FOLDER = os.path.join(BASE_PATH, "共有用", "06_分類依頼", "freee")
DONE_FOLDER = os.path.join(FREEE_FOLDER, "処理済み")
REVIEW_FOLDER = os.path.join(FREEE_FOLDER, "要確認")

# ── Google認証 ──
SCOPES = [
    "https://www.googleapis.com/auth/gmail.send",
]
TOKEN_PATH = os.path.expanduser("~/law-secretary/secrets/token.json")

# ── メール設定 ──
NOTIFY_TO = ["n.kometani@re-startlaw.com"]
SENDER = "n.kometani@re-startlaw.com"
EMAIL_SIGNATURE = (
    "\n\n--\n"
    "〒170-6012 東京都豊島区東池袋３丁目１−１ サンシャイン60 12階\n"
    "弁護士法人Re-Start法律事務所\n"
    "弁護士 米谷尚起\n"
    "TEL : 03-6820-3815"
)

# ── freee API ──
FREEE_API_BASE = "https://api.freee.co.jp"
RATE_LIMIT_WAIT = 0.5  # 500req/時 に余裕を持たせる

# ── ログ設定 ──
LOG_DIR = os.path.expanduser("~/law-secretary/logs")
os.makedirs(LOG_DIR, exist_ok=True)

log_file = os.path.join(LOG_DIR, f"freee_sync_{datetime.now():%Y%m%d}.log")
logging.basicConfig(
    filename=log_file,
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
logger = logging.getLogger(__name__)

# コンソールにも出力
console = logging.StreamHandler()
console.setLevel(logging.INFO)
console.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))
logging.getLogger().addHandler(console)

# ── 勘定科目キーワード（secretary.py と同じ基準 + freee用ID対応）──
ACCOUNT_KEYWORDS = {
    "地代家賃":     ["リージャス", "賃料", "家賃"],
    "通信費":       ["Google", "通信", "インターネット", "電話"],
    "外注費":       ["接見", "弁護士費用", "外注"],
    "会議費":       ["スターバックス", "会議"],
    "接待交際費":   ["接待", "交際"],
    "旅費交通費":   ["交通", "タクシー", "電車", "小田急", "PASMO", "Suica"],
    "消耗品費":     ["消耗品", "備品"],
    "広告宣伝費":   ["広告", "宣伝"],
    "支払手数料":   ["手数料", "Claude", "サービス利用"],
    "研修費":       ["研修"],
    "新聞図書費":   ["新聞", "図書", "書籍"],
    "諸会費":       ["会費"],
    "租税公課":     ["弁護士会", "印紙"],
    "雑費":         ["差し入れ"],
}


# ════════════════════════════════════════════════════════════════
# PDF解析
# ════════════════════════════════════════════════════════════════

def extract_pdf_info(pdf_path):
    """PDFからレシート・請求書の情報を抽出する。
    Returns:
        dict | None: {"date", "amount", "vendor", "description", "tax_amount"}
    """
    try:
        text = ""
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                page_text = page.extract_text()
                if page_text:
                    text += page_text + "\n"

        if not text.strip():
            logger.warning("PDF テキスト抽出不可（スキャン画像の可能性）: %s", pdf_path)
            return None

        info = {
            "date": _extract_date(text),
            "amount": _extract_amount(text),
            "vendor": _extract_vendor(text),
            "description": _extract_description(text, os.path.basename(pdf_path)),
            "tax_amount": _extract_tax(text),
        }

        if not info["date"] or not info["amount"]:
            logger.warning("必須項目（日付 or 金額）抽出不可: %s", pdf_path)
            return None

        return info

    except Exception as e:
        logger.error("PDF解析エラー: %s - %s", pdf_path, e)
        return None


def _extract_date(text):
    """テキストから日付を抽出する。"""
    patterns = [
        r"(\d{4})[年/\-.](\d{1,2})[月/\-.](\d{1,2})",  # 2026年4月14日, 2026/4/14
        r"(令和\s*\d+)[年](\d{1,2})[月](\d{1,2})",       # 令和8年4月14日
        r"(R\s*\d+)[./](\d{1,2})[./](\d{1,2})",           # R8.4.14
    ]
    for pat in patterns:
        m = re.search(pat, text)
        if m:
            g1 = m.group(1)
            month = int(m.group(2))
            day = int(m.group(3))
            if "令和" in g1 or g1.startswith("R"):
                num = int(re.search(r"\d+", g1).group())
                year = 2018 + num
            else:
                year = int(g1)
            if 2020 <= year <= 2030 and 1 <= month <= 12 and 1 <= day <= 31:
                return f"{year:04d}-{month:02d}-{day:02d}"
    return None


def _extract_amount(text):
    """テキストから合計金額（税込）を抽出する。"""
    patterns = [
        r"合計[金額]?\s*[¥￥]?\s*([\d,]+)",
        r"請求金額\s*[¥￥]?\s*([\d,]+)",
        r"お支払[い]?\s*[¥￥]?\s*([\d,]+)",
        r"税込[合計]?\s*[¥￥]?\s*([\d,]+)",
        r"[¥￥]\s*([\d,]+)",
        r"([\d,]{4,})\s*円",
    ]
    amounts = []
    for pat in patterns:
        for m in re.finditer(pat, text):
            val = int(m.group(1).replace(",", ""))
            if val > 0:
                amounts.append(val)
    return max(amounts) if amounts else None


def _extract_tax(text):
    """消費税額を抽出する。"""
    patterns = [
        r"消費税[額]?\s*[¥￥]?\s*([\d,]+)",
        r"税[額]?\s*[¥￥]?\s*([\d,]+)",
    ]
    for pat in patterns:
        m = re.search(pat, text)
        if m:
            return int(m.group(1).replace(",", ""))
    return None


def _extract_vendor(text):
    """取引先名を抽出する（先頭数行から推定）。"""
    lines = [l.strip() for l in text.split("\n") if l.strip()]
    for line in lines[:5]:
        if any(kw in line for kw in ["株式会社", "合同会社", "有限会社", "法人"]):
            return line[:40]
    if lines:
        return lines[0][:40]
    return "不明"


def _extract_description(text, filename):
    """摘要を生成する。"""
    keywords = ["領収書", "請求書", "明細", "レシート", "納品書"]
    for kw in keywords:
        if kw in text or kw in filename:
            return kw
    return os.path.splitext(filename)[0][:30]


def guess_account_item(text, filename):
    """テキスト+ファイル名から勘定科目を推定する。"""
    combined = text + " " + filename
    for account, keywords in ACCOUNT_KEYWORDS.items():
        if any(kw in combined for kw in keywords):
            return account
    return None


# ════════════════════════════════════════════════════════════════
# freee API 呼び出し
# ════════════════════════════════════════════════════════════════

def _api_get(path, params=None):
    """freee API GET リクエスト（レート制限付き）。"""
    time.sleep(RATE_LIMIT_WAIT)
    url = f"{FREEE_API_BASE}{path}"
    resp = requests.get(url, headers=freee_auth.get_headers(), params=params, timeout=30)
    resp.raise_for_status()
    return resp.json()


def _api_post(path, json_data):
    """freee API POST リクエスト（レート制限付き）。"""
    time.sleep(RATE_LIMIT_WAIT)
    url = f"{FREEE_API_BASE}{path}"
    resp = requests.post(url, headers=freee_auth.get_headers(), json=json_data, timeout=30)
    resp.raise_for_status()
    return resp.json()


def _api_put(path, json_data):
    """freee API PUT リクエスト（レート制限付き）。"""
    time.sleep(RATE_LIMIT_WAIT)
    url = f"{FREEE_API_BASE}{path}"
    resp = requests.put(url, headers=freee_auth.get_headers(), json=json_data, timeout=30)
    resp.raise_for_status()
    return resp.json()


def create_deal(info, account_item_name):
    """freee に取引（支出）を登録する。
    Returns:
        dict: 登録された取引データ
    """
    company_id = freee_auth.get_company_id()

    # 税額が分かれば税抜・税額を分離、不明なら税込金額をそのまま使う
    amount = info["amount"]
    tax = info["tax_amount"]

    deal_data = {
        "company_id": company_id,
        "issue_date": info["date"],
        "type": "expense",
        "details": [
            {
                "account_item_id": 0,  # 名前では指定できないので後述
                "tax_code": 2,         # 課対仕入10%
                "amount": amount,
                "description": f"{info['vendor']} {info['description']}",
            }
        ],
    }

    # 勘定科目IDを名前から解決
    account_id = resolve_account_item_id(company_id, account_item_name)
    if account_id:
        deal_data["details"][0]["account_item_id"] = account_id
    else:
        logger.warning("勘定科目ID解決失敗: %s → デフォルト（雑費）で登録", account_item_name)

    return _api_post("/api/1/deals", deal_data)


_account_item_cache = {}


def resolve_account_item_id(company_id, name):
    """勘定科目名からIDを解決する（キャッシュ付き）。"""
    if not _account_item_cache:
        _load_account_items(company_id)
    return _account_item_cache.get(name)


def _load_account_items(company_id):
    """勘定科目一覧をAPIから取得してキャッシュする。"""
    try:
        data = _api_get("/api/1/account_items", {"company_id": company_id})
        for item in data.get("account_items", []):
            _account_item_cache[item["name"]] = item["id"]
        logger.info("勘定科目マスタ取得完了: %d件", len(_account_item_cache))
    except Exception as e:
        logger.error("勘定科目マスタ取得失敗: %s", e)


# ════════════════════════════════════════════════════════════════
# 機能1: PDFレシート・請求書の自動取引登録
# ════════════════════════════════════════════════════════════════

def sync_pdf_deals():
    """freee/フォルダ内のPDFを読み取り、freeeに取引登録する。
    Returns:
        list[dict]: 処理結果リスト
    """
    results = []

    if not os.path.isdir(FREEE_FOLDER):
        logger.info("freeeフォルダが存在しません: %s", FREEE_FOLDER)
        return results

    os.makedirs(DONE_FOLDER, exist_ok=True)
    os.makedirs(REVIEW_FOLDER, exist_ok=True)

    pdf_files = [f for f in os.listdir(FREEE_FOLDER)
                 if f.lower().endswith(".pdf") and os.path.isfile(os.path.join(FREEE_FOLDER, f))]

    if not pdf_files:
        logger.info("処理対象のPDFファイルなし")
        return results

    logger.info("PDF取引登録開始: %d件", len(pdf_files))

    for filename in pdf_files:
        src_path = os.path.join(FREEE_FOLDER, filename)
        result = {"filename": filename, "status": "", "detail": ""}

        # PDF解析
        info = extract_pdf_info(src_path)
        if not info:
            result["status"] = "要確認"
            result["detail"] = "PDF読み取り不可（スキャン画像 or 必須項目なし）"
            shutil.move(src_path, os.path.join(REVIEW_FOLDER, filename))
            logger.warning("要確認へ移動: %s", filename)
            results.append(result)
            continue

        # 勘定科目推定
        with pdfplumber.open(src_path) as pdf:
            full_text = "\n".join(p.extract_text() or "" for p in pdf.pages)
        account_name = guess_account_item(full_text, filename)
        if not account_name:
            account_name = "雑費"

        # freee API で取引登録
        try:
            deal = create_deal(info, account_name)
            deal_id = deal.get("deal", {}).get("id", "?")
            result["status"] = "登録済み"
            result["detail"] = (
                f"取引ID: {deal_id} / {info['date']} / "
                f"¥{info['amount']:,} / {info['vendor']} / "
                f"勘定科目: {account_name}"
            )
            shutil.move(src_path, os.path.join(DONE_FOLDER, filename))
            logger.info("取引登録成功: %s → deal_id=%s", filename, deal_id)

        except requests.HTTPError as e:
            result["status"] = "登録失敗"
            result["detail"] = f"APIエラー: {e.response.status_code} {e.response.text[:200]}"
            shutil.move(src_path, os.path.join(REVIEW_FOLDER, filename))
            logger.error("取引登録失敗: %s - %s", filename, e)

        except Exception as e:
            result["status"] = "登録失敗"
            result["detail"] = f"エラー: {str(e)[:200]}"
            shutil.move(src_path, os.path.join(REVIEW_FOLDER, filename))
            logger.error("取引登録失敗: %s - %s", filename, e)

        results.append(result)

    return results


# ════════════════════════════════════════════════════════════════
# 機能2: ファイルボックス未連携ファイルの取引紐付け
# ════════════════════════════════════════════════════════════════

def link_unmatched_receipts():
    """ファイルボックスの未連携ファイルを取引に紐付ける。
    Returns:
        list[dict]: 処理結果リスト
    """
    results = []
    company_id = freee_auth.get_company_id()

    try:
        data = _api_get("/api/1/receipts", {
            "company_id": company_id,
            "start_date": "2026-01-01",
            "limit": 100,
        })
    except Exception as e:
        logger.error("ファイルボックス取得失敗: %s", e)
        return results

    receipts = data.get("receipts", [])
    unlinked = [r for r in receipts if r.get("deal_id") is None]

    if not unlinked:
        logger.info("未連携ファイルなし")
        return results

    logger.info("未連携ファイル: %d件", len(unlinked))

    for receipt in unlinked:
        receipt_id = receipt["id"]
        issue_date = receipt.get("issue_date")
        amount = receipt.get("receipt_metadatum", {}).get("amount")
        result = {
            "receipt_id": receipt_id,
            "filename": receipt.get("file_name", "不明"),
            "status": "",
            "detail": "",
        }

        if not issue_date or not amount:
            result["status"] = "スキップ"
            result["detail"] = "日付 or 金額が不明"
            results.append(result)
            continue

        # 取引を検索（日付 ± 3日、金額一致）
        try:
            deals_data = _api_get("/api/1/deals", {
                "company_id": company_id,
                "start_issue_date": issue_date,
                "end_issue_date": issue_date,
                "limit": 50,
            })
        except Exception as e:
            result["status"] = "検索失敗"
            result["detail"] = str(e)[:200]
            results.append(result)
            continue

        matched_deal = None
        for deal in deals_data.get("deals", []):
            deal_amount = sum(
                abs(d.get("amount", 0)) for d in deal.get("details", [])
            )
            if deal_amount == amount:
                matched_deal = deal
                break

        if matched_deal:
            try:
                _api_put(f"/api/1/receipts/{receipt_id}", {
                    "company_id": company_id,
                    "deal_id": matched_deal["id"],
                })
                result["status"] = "紐付け済み"
                result["detail"] = f"取引ID: {matched_deal['id']}"
                logger.info("紐付け成功: receipt=%d → deal=%d",
                            receipt_id, matched_deal["id"])
            except Exception as e:
                result["status"] = "紐付け失敗"
                result["detail"] = str(e)[:200]
                logger.error("紐付け失敗: receipt=%d - %s", receipt_id, e)
        else:
            result["status"] = "該当取引なし"
            result["detail"] = f"日付={issue_date}, 金額={amount}"
            logger.info("該当取引なし: receipt=%d", receipt_id)

        results.append(result)

    return results


# ════════════════════════════════════════════════════════════════
# メール通知（番号付き一覧で修正指示しやすい形式）
# ════════════════════════════════════════════════════════════════

def get_gmail_service():
    """Gmail API サービスを取得する。"""
    creds = Credentials.from_authorized_user_file(TOKEN_PATH, SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(GoogleRequest())
    return build("gmail", "v1", credentials=creds)


def send_report(deal_results, receipt_results):
    """処理結果をメールで報告する。

    番号付き一覧にして「#3 の勘定科目が違います」と返信しやすい形式にする。
    """
    lines = [
        "米谷先生",
        "",
        "お疲れ様です。freee連携処理が完了しましたのでご報告いたします。",
        "",
    ]

    # ── 機能1: 取引登録結果 ──
    if deal_results:
        lines.append("━━━ PDF取引登録 ━━━")
        lines.append("")
        for i, r in enumerate(deal_results, 1):
            status_mark = "○" if r["status"] == "登録済み" else "△" if r["status"] == "要確認" else "×"
            lines.append(f"  #{i}  {status_mark} {r['filename']}")
            lines.append(f"       {r['detail']}")
            lines.append("")
        lines.append(
            "※ 修正が必要な場合は番号を指定してご返信ください。"
        )
        lines.append(
            '  例: 「#3 の勘定科目を旅費交通費に変更」「#5 は登録不要」'
        )
        lines.append("")
    else:
        lines.append("PDF取引登録: 対象ファイルなし")
        lines.append("")

    # ── 機能2: ファイルボックス紐付け結果 ──
    if receipt_results:
        lines.append("━━━ ファイルボックス紐付け ━━━")
        lines.append("")
        linked = [r for r in receipt_results if r["status"] == "紐付け済み"]
        unlinked = [r for r in receipt_results if r["status"] != "紐付け済み"]
        if linked:
            lines.append(f"  紐付け成功: {len(linked)}件")
            for r in linked:
                lines.append(f"    - {r['filename']} → {r['detail']}")
        if unlinked:
            lines.append(f"  未紐付け: {len(unlinked)}件")
            for r in unlinked:
                lines.append(f"    - {r['filename']}（{r['status']}: {r['detail']}）")
        lines.append("")
    else:
        lines.append("ファイルボックス紐付け: 未連携ファイルなし")
        lines.append("")

    lines.append("ご確認のほどよろしくお願いいたします。")
    lines.append(EMAIL_SIGNATURE)

    body_text = "\n".join(lines)

    # メール送信
    gmail = get_gmail_service()
    message = MIMEText(body_text, "plain", "utf-8")
    message["to"] = ", ".join(NOTIFY_TO)
    message["from"] = SENDER
    message["subject"] = f"freee連携処理結果のご報告（{datetime.now():%m/%d}）"

    raw = base64.urlsafe_b64encode(message.as_bytes()).decode("utf-8")
    gmail.users().messages().send(userId="me", body={"raw": raw}).execute()
    logger.info("報告メール送信完了: %s", ", ".join(NOTIFY_TO))


# ════════════════════════════════════════════════════════════════
# メイン処理
# ════════════════════════════════════════════════════════════════

def main():
    logger.info("=== freee連携処理開始 ===")
    print(f"ログ: {log_file}")

    # 機能1: PDF取引登録
    deal_results = sync_pdf_deals()

    # 機能2: ファイルボックス紐付け
    receipt_results = link_unmatched_receipts()

    # メール報告
    if deal_results or receipt_results:
        send_report(deal_results, receipt_results)
        print("報告メールを送信しました。")
    else:
        logger.info("処理対象なし — メール送信スキップ")
        print("処理対象なし。")

    logger.info("=== freee連携処理完了 ===")


if __name__ == "__main__":
    main()
