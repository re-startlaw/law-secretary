"""freeeファイルボックスへのレシートアップロード。

POST /api/1/receipts (multipart/form-data) で領収書/請求書等のPDFを
freeeのファイルボックスへ登録する。アップロード結果（receipt id・取得URL）を返す。

CLI使用例:
    venv/bin/python freee_filebox.py upload \
        --file "/path/to/領収書.pdf" \
        --issue-date 2026-05-20 \
        --description "ピアリビング 防音パネル代" \
        --business corporate
"""

import argparse
import logging
import mimetypes
import os
import sys
import time

import requests

import freee_auth

FREEE_API_BASE = "https://api.freee.co.jp"
RATE_LIMIT_WAIT = 0.5

logger = logging.getLogger(__name__)


def upload_receipt(
    file_path: str,
    issue_date: str,
    description: str = "",
    business_type: str = "corporate",
) -> dict:
    """ファイルボックスにファイルをアップロードする。

    Args:
        file_path: アップロード対象のローカルパス
        issue_date: 発行日 (YYYY-MM-DD)
        description: 摘要（任意）
        business_type: "corporate"=法人, "personal"=個人

    Returns:
        dict: {"id": int, "file_name": str, "issue_date": str, ...}
              + 追加で "ui_url" (ファイルボックスUIの直リンク推定)
    """
    if not os.path.isfile(file_path):
        raise FileNotFoundError(file_path)

    company_id = freee_auth.get_company_id(business_type)
    token = freee_auth.get_access_token()

    mime = mimetypes.guess_type(file_path)[0] or "application/octet-stream"
    filename = os.path.basename(file_path)

    url = f"{FREEE_API_BASE}/api/1/receipts"
    headers = {"Authorization": f"Bearer {token}"}
    data = {
        "company_id": str(company_id),
        "issue_date": issue_date,
    }
    if description:
        data["description"] = description

    time.sleep(RATE_LIMIT_WAIT)
    with open(file_path, "rb") as fh:
        files = {"receipt": (filename, fh, mime)}
        resp = requests.post(url, headers=headers, data=data, files=files, timeout=60)

    if resp.status_code >= 400:
        raise RuntimeError(f"freee upload failed: HTTP {resp.status_code} {resp.text}")

    payload = resp.json()
    receipt = payload.get("receipt") or payload
    receipt["ui_url"] = (
        f"https://secure.freee.co.jp/receipts?company_id={company_id}"
        f"&receipt_id={receipt.get('id')}"
    )
    receipt["company_id"] = company_id
    return receipt


def _cli():
    p = argparse.ArgumentParser()
    sub = p.add_subparsers(dest="cmd", required=True)
    up = sub.add_parser("upload")
    up.add_argument("--file", required=True)
    up.add_argument("--issue-date", required=True, help="YYYY-MM-DD")
    up.add_argument("--description", default="")
    up.add_argument("--business", choices=["corporate", "personal"], default="corporate")
    args = p.parse_args()

    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
    if args.cmd == "upload":
        r = upload_receipt(args.file, args.issue_date, args.description, args.business)
        print(f"OK receipt_id={r.get('id')}  ui_url={r.get('ui_url')}")
        print(r)
        return 0
    return 1


if __name__ == "__main__":
    sys.exit(_cli())
