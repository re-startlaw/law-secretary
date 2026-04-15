"""
freee API OAuth2 トークン管理モジュール
- .env からトークンを読み込み
- 期限切れ前に自動リフレッシュ
- 新トークンを .env に上書き保存
"""

import logging
import os
import time

import requests
from dotenv import dotenv_values, set_key

ENV_PATH = os.path.expanduser("~/law-secretary/.env")
TOKEN_URL = "https://accounts.secure.freee.co.jp/public_api/token"

logger = logging.getLogger(__name__)


def _load_env():
    """現在の .env の値を読み込む。"""
    return dotenv_values(ENV_PATH)


def _refresh_token(env):
    """リフレッシュトークンで新しいアクセストークンを取得し .env を更新する。"""
    resp = requests.post(TOKEN_URL, data={
        "grant_type": "refresh_token",
        "client_id": env["FREEE_CLIENT_ID"],
        "client_secret": env["FREEE_CLIENT_SECRET"],
        "refresh_token": env["FREEE_REFRESH_TOKEN"],
    }, timeout=30)
    resp.raise_for_status()
    data = resp.json()

    new_access = data["access_token"]
    new_refresh = data["refresh_token"]
    expires_at = int(time.time()) + data.get("expires_in", 21600)

    set_key(ENV_PATH, "FREEE_ACCESS_TOKEN", new_access)
    set_key(ENV_PATH, "FREEE_REFRESH_TOKEN", new_refresh)
    set_key(ENV_PATH, "FREEE_TOKEN_EXPIRES_AT", str(expires_at))

    logger.info("freee トークンをリフレッシュしました（有効期限: %s）",
                time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(expires_at)))
    return new_access


def get_access_token():
    """有効なアクセストークンを返す。期限切れ5分前なら自動リフレッシュ。"""
    env = _load_env()
    expires_at = int(env.get("FREEE_TOKEN_EXPIRES_AT") or "0")

    if time.time() > expires_at - 300:
        return _refresh_token(env)
    return env["FREEE_ACCESS_TOKEN"]


def get_headers():
    """freee API 用の認証ヘッダーを返す。"""
    token = get_access_token()
    return {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    }


def get_company_id():
    """事業所IDを返す。"""
    env = _load_env()
    return int(env["FREEE_COMPANY_ID"])
