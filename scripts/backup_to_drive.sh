#!/bin/bash
# 事務所Google Driveのマイドライブ直下（本人専用領域）へ
# ~/law-secretary/ をミラー同期する。launchd から毎日23:00に起動される。
# 目的: PC紛失時の引き継ぎ。Gitで管理していない依頼者情報入りスクリプトや
# バックアップフォルダ、.env・secrets等の機密もまとめて保全する。
# 同期先は共有用配下ではないため、事務員など他ユーザーは参照できない。

set -euo pipefail

SRC="/Users/kometaninaoki/law-secretary/"
DEST="/Users/kometaninaoki/Library/CloudStorage/GoogleDrive-n.kometani@re-startlaw.com/マイドライブ/秘書エージェント_バックアップ/law-secretary/"
LOG_DIR="$HOME/law-secretary/.sync"
LOG_FILE="$LOG_DIR/backup.log"

mkdir -p "$LOG_DIR"
mkdir -p "$DEST"

{
  echo "=== $(date '+%Y-%m-%d %H:%M:%S') START ==="
  rsync -a --delete \
    --exclude='.git/' \
    --exclude='venv/' \
    --exclude='__pycache__/' \
    --exclude='*.pyc' \
    --exclude='.DS_Store' \
    --exclude='.sync/' \
    --exclude='logs/' \
    --exclude='secretary.log' \
    "$SRC" "$DEST"
  echo "=== $(date '+%Y-%m-%d %H:%M:%S') END (exit=$?) ==="
} >> "$LOG_FILE" 2>&1
