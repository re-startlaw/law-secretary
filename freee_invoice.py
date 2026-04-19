"""
freee 請求書作成スクリプト（請求書freee API版）。

- 請求書freee API (/iv/invoices) で請求書を作成する
- 取引先は会計freee API (/api/1/partners) から解決（両プロダクトでID共有）
- メール送信・URL共有は API で不可能。作成後に freee UI の URL を返す
  → 米谷弁護士が UI で「URL共有で送付」を1クリックするだけで送付完了

使い方:
venv/bin/python freee_invoice.py create \
  --business corporate \
  --partner-name "取引先名" \
  --title "件名" \
  --amount 55000 \
  --issue-date 2026-04-17 \
  --due-date 2026-04-30

複数品目の場合:
  --items '[{"description":"項目A","amount":400000},{"description":"項目B","amount":100000}]'
"""

import argparse
import datetime as dt
import json

import requests

import freee_auth

FREEE_API_BASE = "https://api.freee.co.jp"
INVOICE_UI_BASE = "https://invoice.secure.freee.co.jp"


def _api_get(path, params=None):
    url = f"{FREEE_API_BASE}{path}"
    res = requests.get(url, headers=freee_auth.get_headers(), params=params, timeout=30)
    res.raise_for_status()
    return res.json()


def _api_post(path, payload, params=None):
    url = f"{FREEE_API_BASE}{path}"
    res = requests.post(url, headers=freee_auth.get_headers(),
                        json=payload, params=params, timeout=30)
    if not res.ok:
        raise RuntimeError(f"POST {path} failed: {res.status_code} {res.text}")
    return res.json()


def resolve_partner(company_id, partner_name):
    data = _api_get("/api/1/partners", {"company_id": company_id, "limit": 300})
    partners = data.get("partners", [])
    exact = [p for p in partners if p.get("name") == partner_name]
    if exact:
        return exact[0]
    partial = [p for p in partners if partner_name in (p.get("name") or "")]
    if partial:
        return partial[0]
    return None


def resolve_template_id(company_id):
    data = _api_get("/iv/invoices/templates", {"company_id": company_id})
    templates = data.get("templates", [])
    if not templates:
        raise RuntimeError(f"請求書テンプレートが見つかりません (company_id={company_id})")
    return templates[0]["id"]


def build_lines(args):
    if args.items:
        items_list = json.loads(args.items)
    else:
        items_list = [{"description": args.description, "amount": args.amount}]
    return [
        {
            "type": "item",
            "description": item["description"],
            "quantity": item.get("quantity", 1),
            "unit_price": str(item["amount"]),
            "tax_rate": item.get("tax_rate", 10),
            "reduced_tax_rate": item.get("reduced_tax_rate", False),
            "withholding": item.get("withholding", False),
        }
        for item in items_list
    ]


def create_invoice(args):
    company_id = freee_auth.get_company_id(args.business)

    partner = resolve_partner(company_id, args.partner_name)
    if not partner:
        raise RuntimeError(
            f"取引先が見つかりません: {args.partner_name} (company_id={company_id})"
        )

    template_id = args.template_id or resolve_template_id(company_id)
    partner_title = args.partner_title or partner.get("default_title") or "様"

    payload = {
        "company_id": company_id,
        "template_id": template_id,
        "subject": args.title,
        "billing_date": args.issue_date,
        "payment_date": args.due_date,
        "payment_type": "transfer",
        "tax_entry_method": "out",
        "tax_fraction": "omit",
        "line_amount_fraction": "omit",
        "withholding_tax_entry_method": "out",
        "partner_id": partner["id"],
        "partner_title": partner_title,
        "memo": args.description or "",
        "invoice_note": args.note or "",
        "lines": build_lines(args),
    }

    created = _api_post("/iv/invoices", payload)
    invoice = created.get("invoice", created)
    invoice_id = invoice.get("id")

    print("請求書を作成しました。")
    print(f"  invoice_id     : {invoice_id}")
    print(f"  invoice_number : {invoice.get('invoice_number')}")
    print(f"  partner        : {invoice.get('partner_display_name')}")
    print(f"  billing_date   : {invoice.get('billing_date')}")
    print(f"  payment_date   : {invoice.get('payment_date')}")
    print(f"  amount(tax_in) : {invoice.get('amount_including_tax')}")
    print(f"  UI URL         : {INVOICE_UI_BASE}/reports/invoices/{invoice_id}")
    print()
    print("※ 送付は freee UI の上記URLを開き「URL共有で送付」をクリックしてください。")
    return invoice


def build_parser():
    today = dt.date.today()
    end_of_month = (today.replace(day=28) + dt.timedelta(days=4)).replace(day=1) - dt.timedelta(days=1)

    parser = argparse.ArgumentParser(description="請求書freee 請求書作成")
    sub = parser.add_subparsers(dest="command", required=True)

    p = sub.add_parser("create", help="請求書を作成")
    p.add_argument("--business", choices=["corporate", "personal"], required=True)
    p.add_argument("--partner-name", required=True)
    p.add_argument("--partner-title", default=None, help="様/御中 (省略時は取引先のdefault_title、なければ『様』)")
    p.add_argument("--title", default="請求書", help="件名（subject）")
    p.add_argument("--description", default="", help="備考メモ（memo）")
    p.add_argument("--note", default="", help="請求書備考欄（invoice_note）")
    p.add_argument("--amount", type=float, default=0, help="単一項目の税別金額（--items未指定時）")
    p.add_argument("--issue-date", default=today.isoformat(), help="billing_date")
    p.add_argument("--due-date", default=end_of_month.isoformat(), help="payment_date")
    p.add_argument("--template-id", type=int, default=None)
    p.add_argument("--items", default=None,
                   help='JSON array: [{"description":...,"amount":...,"quantity":1,"tax_rate":10},...]')
    p.set_defaults(func=lambda a: create_invoice(a))

    return parser


def main():
    parser = build_parser()
    args = parser.parse_args()
    args.func(args)


if __name__ == "__main__":
    main()
