"""
freee API 初回接続セットアップ用スクリプト。

使い方:
1) 認可URL生成
   venv/bin/python freee_connect.py auth-url --client-id xxx --redirect-uri http://localhost:8080/callback

2) ブラウザで認可後、code を取得してトークン交換
   venv/bin/python freee_connect.py exchange --client-id xxx --client-secret yyy --redirect-uri http://localhost:8080/callback --code zzz
"""

from __future__ import annotations

import argparse
import secrets
import time
from pathlib import Path
from urllib.parse import urlencode

import requests
from dotenv import dotenv_values, set_key

AUTH_URL = "https://accounts.secure.freee.co.jp/public_api/authorize"
TOKEN_URL = "https://accounts.secure.freee.co.jp/public_api/token"
API_BASE = "https://api.freee.co.jp/api/1"
ENV_PATH = Path(__file__).resolve().parent / ".env"


def ensure_env_file() -> None:
    if not ENV_PATH.exists():
        ENV_PATH.write_text("", encoding="utf-8")


def build_auth_url(client_id: str, redirect_uri: str, state: str) -> str:
    params = {
        "client_id": client_id,
        "redirect_uri": redirect_uri,
        "response_type": "code",
        "state": state,
        "prompt": "consent",
    }
    return f"{AUTH_URL}?{urlencode(params)}"


def exchange_code(client_id: str, client_secret: str, redirect_uri: str, code: str) -> dict:
    response = requests.post(
        TOKEN_URL,
        data={
            "grant_type": "authorization_code",
            "client_id": client_id,
            "client_secret": client_secret,
            "code": code,
            "redirect_uri": redirect_uri,
        },
        timeout=30,
    )
    response.raise_for_status()
    return response.json()


def fetch_company_id(access_token: str) -> int | None:
    headers = {"Authorization": f"Bearer {access_token}"}
    response = requests.get(f"{API_BASE}/companies", headers=headers, timeout=30)
    response.raise_for_status()
    companies = response.json().get("companies", [])
    if not companies:
        return None
    return int(companies[0]["id"])


def save_tokens(client_id: str, client_secret: str, token_data: dict, company_id: int | None) -> None:
    ensure_env_file()
    expires_at = int(time.time()) + int(token_data.get("expires_in", 21600))
    set_key(str(ENV_PATH), "FREEE_CLIENT_ID", client_id)
    set_key(str(ENV_PATH), "FREEE_CLIENT_SECRET", client_secret)
    set_key(str(ENV_PATH), "FREEE_ACCESS_TOKEN", token_data["access_token"])
    set_key(str(ENV_PATH), "FREEE_REFRESH_TOKEN", token_data["refresh_token"])
    set_key(str(ENV_PATH), "FREEE_TOKEN_EXPIRES_AT", str(expires_at))
    if company_id is not None:
        set_key(str(ENV_PATH), "FREEE_COMPANY_ID", str(company_id))


def cmd_auth_url(args: argparse.Namespace) -> int:
    state = args.state or secrets.token_urlsafe(16)
    url = build_auth_url(args.client_id, args.redirect_uri, state)
    print("以下のURLをブラウザで開いて認可してください。")
    print(url)
    print("")
    print(f"state: {state}")
    return 0


def cmd_exchange(args: argparse.Namespace) -> int:
    token_data = exchange_code(
        client_id=args.client_id,
        client_secret=args.client_secret,
        redirect_uri=args.redirect_uri,
        code=args.code,
    )
    company_id = fetch_company_id(token_data["access_token"])
    save_tokens(args.client_id, args.client_secret, token_data, company_id)
    print(f".env にfreeeトークンを保存しました: {ENV_PATH}")
    if company_id is not None:
        print(f"FREEE_COMPANY_ID={company_id} を自動設定しました。")
    else:
        print("事業所IDを取得できませんでした。手動で FREEE_COMPANY_ID を設定してください。")
    return 0


def cmd_check(_: argparse.Namespace) -> int:
    env = dotenv_values(str(ENV_PATH))
    access_token = env.get("FREEE_ACCESS_TOKEN")
    if not access_token:
        print(".env に FREEE_ACCESS_TOKEN がありません。まず exchange を実行してください。")
        return 1
    company_id = env.get("FREEE_COMPANY_ID")
    headers = {"Authorization": f"Bearer {access_token}"}
    params = {"company_id": company_id} if company_id else None
    response = requests.get(f"{API_BASE}/account_items", headers=headers, params=params, timeout=30)
    response.raise_for_status()
    print("freee API 接続確認OK")
    return 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="freee API 初回接続セットアップ")
    sub = parser.add_subparsers(dest="command", required=True)

    p_auth = sub.add_parser("auth-url", help="認可URLを生成する")
    p_auth.add_argument("--client-id", required=True)
    p_auth.add_argument("--redirect-uri", required=True)
    p_auth.add_argument("--state")
    p_auth.set_defaults(func=cmd_auth_url)

    p_exchange = sub.add_parser("exchange", help="認可コードをトークン交換して .env に保存する")
    p_exchange.add_argument("--client-id", required=True)
    p_exchange.add_argument("--client-secret", required=True)
    p_exchange.add_argument("--redirect-uri", required=True)
    p_exchange.add_argument("--code", required=True)
    p_exchange.set_defaults(func=cmd_exchange)

    p_check = sub.add_parser("check", help="保存済みトークンで接続確認する")
    p_check.set_defaults(func=cmd_check)

    return parser


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()
    return args.func(args)


if __name__ == "__main__":
    raise SystemExit(main())
