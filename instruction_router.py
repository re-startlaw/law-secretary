"""
自然文の業務指示を受け取り、freee請求書作成まで実行するルーター。

例:
  venv/bin/python instruction_router.py "請求書を作成して"
"""

from __future__ import annotations

import argparse
import re
from types import SimpleNamespace

import freee_invoice


def _prompt(label: str, default: str | None = None) -> str:
    suffix = f" [{default}]" if default else ""
    value = input(f"{label}{suffix}: ").strip()
    if value:
        return value
    if default is not None:
        return default
    return _prompt(label, default)


def _build_invoice_args_from_prompt() -> SimpleNamespace:
    print("請求書の作成に必要な情報を入力してください。")
    business = _prompt("事業区分 corporate/personal", "corporate")
    partner_name = _prompt("取引先名")
    title = _prompt("請求書タイトル", "請求書")
    description = _prompt("備考メモ", "")
    amount = float(_prompt("金額（税別、円）", "0"))
    issue_date = _prompt("請求日 yyyy-mm-dd")
    due_date = _prompt("支払期日 yyyy-mm-dd")

    return SimpleNamespace(
        business=business,
        partner_name=partner_name,
        partner_title=None,
        title=title,
        description=description,
        note="",
        amount=amount,
        issue_date=issue_date,
        due_date=due_date,
        template_id=None,
        items=None,
    )


def handle_request(request_text: str) -> None:
    normalized = request_text.strip()
    if re.search(r"請求書.*(作成|作って|発行)", normalized):
        args = _build_invoice_args_from_prompt()
        freee_invoice.create_invoice(args)
        return
    raise ValueError(
        "未対応の指示です。現在は『請求書を作成して』系の指示に対応しています。"
    )


def main() -> int:
    parser = argparse.ArgumentParser(description="自然文業務指示ルーター")
    parser.add_argument("request_text", help="例: 請求書を作成して")
    args = parser.parse_args()
    handle_request(args.request_text)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
