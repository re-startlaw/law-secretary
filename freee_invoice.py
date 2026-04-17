"""
freee 請求書作成 + Gmail送付スクリプト。

使い方:
venv/bin/python freee_invoice.py create-and-send \
  --business corporate \
  --partner-name "山田太郎" \
  --partner-email "client@example.com" \
  --title "法律顧問料" \
  --description "2026年4月分 顧問料" \
  --amount 55000 \
  --issue-date 2026-04-17 \
  --due-date 2026-04-30
"""

import argparse
import base64
import datetime as dt
from email.mime.text import MIMEText

import requests
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

import freee_auth

FREEE_API_BASE = "https://api.freee.co.jp/api/1"
SCOPES = ["https://www.googleapis.com/auth/gmail.send"]
TOKEN_PATH = "/Users/kometaninaoki/law-secretary/secrets/token.json"
SENDER = "n.kometani@re-startlaw.com"


def _api_get(path, params=None):
    url = f"{FREEE_API_BASE}{path}"
    res = requests.get(url, headers=freee_auth.get_headers(), params=params, timeout=30)
    res.raise_for_status()
    return res.json()


def _api_post(path, payload):
    url = f"{FREEE_API_BASE}{path}"
    res = requests.post(url, headers=freee_auth.get_headers(), json=payload, timeout=30)
    res.raise_for_status()
    return res.json()


def _api_put(path, payload):
    url = f"{FREEE_API_BASE}{path}"
    res = requests.put(url, headers=freee_auth.get_headers(), json=payload, timeout=30)
    res.raise_for_status()
    return res.json()


def get_gmail_service():
    creds = Credentials.from_authorized_user_file(TOKEN_PATH, SCOPES)
    return build("gmail", "v1", credentials=creds)


def resolve_partner(company_id, partner_name):
    data = _api_get("/partners", {"company_id": company_id, "limit": 100})
    partners = data.get("partners", [])
    exact = [p for p in partners if p.get("name") == partner_name]
    if exact:
        return exact[0]
    partial = [p for p in partners if partner_name in (p.get("name") or "")]
    if partial:
        return partial[0]
    return None


def create_invoice(company_id, partner, args):
    if args.items:
        import json
        items_list = json.loads(args.items)
        invoice_contents = [
            {
                "order": i,
                "type": "normal",
                "qty": item.get("qty", 1),
                "unit": item.get("unit", "式"),
                "unit_price": item["amount"],
                "tax_code": item.get("tax_code", 1),
                "account_item_id": item.get("account_item_id", args.account_item_id),
                "description": item["description"],
            }
            for i, item in enumerate(items_list)
        ]
    else:
        invoice_contents = [
            {
                "order": 0,
                "type": "normal",
                "qty": 1,
                "unit": "式",
                "unit_price": args.amount,
                "tax_code": 1,
                "account_item_id": args.account_item_id,
                "description": args.description,
            }
        ]
    payload = {
        "company_id": company_id,
        "partner_id": partner["id"],
        "partner_display_name": partner.get("name") or args.partner_name,
        "partner_title": "御中",
        "title": args.title,
        "description": args.description,
        "issue_date": args.issue_date,
        "due_date": args.due_date,
        "invoice_status": "unsubmitted",
        "invoice_contents": invoice_contents,
    }
    if args.message:
        payload["message"] = args.message
    return _api_post("/invoices", payload)


def mark_invoice_submitted(company_id, invoice, args):
    payload = {
        "company_id": company_id,
        "partner_id": invoice.get("partner_id"),
        "partner_display_name": invoice.get("partner_display_name"),
        "partner_title": invoice.get("partner_title") or "御中",
        "title": invoice.get("title") or args.title,
        "description": invoice.get("description") or args.description,
        "issue_date": invoice.get("issue_date") or args.issue_date,
        "due_date": invoice.get("due_date") or args.due_date,
        "invoice_status": "submitted",
        "invoice_contents": invoice.get("invoice_contents", []),
    }
    return _api_put(f"/invoices/{invoice['id']}", payload)


def send_invoice_mail(partner_email, partner_name, invoice, business):
    body = (
        f"{partner_name} 様\n\n"
        "いつもお世話になっております。\n"
        "請求書を送付いたしますので、ご確認をお願いいたします。\n\n"
        f"【事業区分】{business}\n"
        f"【請求書ID】{invoice.get('id')}\n"
        f"【請求日】{invoice.get('issue_date')}\n"
        f"【支払期日】{invoice.get('due_date')}\n"
        f"【件名】{invoice.get('title')}\n"
        f"【概要】{invoice.get('description')}\n\n"
        "よろしくお願いいたします。"
    )
    msg = MIMEText(body, "plain", "utf-8")
    msg["to"] = partner_email
    msg["from"] = SENDER
    msg["subject"] = f"請求書送付のご連絡（{invoice.get('title', '請求書')}）"
    raw = base64.urlsafe_b64encode(msg.as_bytes()).decode("utf-8")
    gmail = get_gmail_service()
    gmail.users().messages().send(userId="me", body={"raw": raw}).execute()


def cmd_create_and_send(args):
    company_id = freee_auth.get_company_id(args.business)
    partner = resolve_partner(company_id, args.partner_name)
    if not partner:
        raise RuntimeError(
            f"取引先が見つかりません: {args.partner_name} "
            f"(company_id={company_id})"
        )

    created = create_invoice(company_id, partner, args)
    invoice = created.get("invoice", created)
    mark_invoice_submitted(company_id, invoice, args)
    send_invoice_mail(args.partner_email, args.partner_name, invoice, args.business)
    print(f"請求書を作成して送信しました。invoice_id={invoice.get('id')}")


def build_parser():
    today = dt.date.today()
    end_of_month = (today.replace(day=28) + dt.timedelta(days=4)).replace(day=1) - dt.timedelta(days=1)

    parser = argparse.ArgumentParser(description="freee請求書作成・送付")
    sub = parser.add_subparsers(dest="command", required=True)

    p = sub.add_parser("create-and-send", help="請求書を作成してメール送付")
    p.add_argument("--business", choices=["corporate", "personal"], required=True)
    p.add_argument("--partner-name", required=True)
    p.add_argument("--partner-email", required=True)
    p.add_argument("--title", default="請求書")
    p.add_argument("--description", required=True)
    p.add_argument("--amount", type=float, required=True)
    p.add_argument("--issue-date", default=today.isoformat())
    p.add_argument("--due-date", default=end_of_month.isoformat())
    p.add_argument("--account-item-id", type=int, default=27)
    p.add_argument("--items", default=None, help="JSON array of items: [{\"description\":...,\"amount\":...},...]")
    p.add_argument("--message", default=None, help="備考欄に表示するメッセージ")
    p.set_defaults(func=cmd_create_and_send)

    return parser


def main():
    parser = build_parser()
    args = parser.parse_args()
    args.func(args)


if __name__ == "__main__":
    main()
