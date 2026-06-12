#!/usr/bin/env bash
# 記録ビューア起動スクリプト。
#   bash tools/kiroku-viewer/start.sh
# で localhost:8788 にサーバを立てる。ポート占有時は lsof で検知して案内する。
set -euo pipefail

HERE="$(cd "$(dirname "$0")" && pwd)"
REPO_ROOT="$(cd "$HERE/../.." && pwd)"
PORT=8788
PYTHON="$REPO_ROOT/venv/bin/python"

if [ ! -x "$PYTHON" ]; then
  echo "venv が見つかりません: $PYTHON" >&2
  echo "リポジトリ直下で 'python3 -m venv venv' 後に依存をインストールしてください。" >&2
  exit 1
fi

if [ ! -f "$HERE/cases.json" ]; then
  echo "警告: cases.json がありません。ブラウザの「＋ 追加」から事件を登録してください。" >&2
fi

if lsof -nP -iTCP:"$PORT" -sTCP:LISTEN >/dev/null 2>&1; then
  echo "ポート $PORT は既に使用中です。占有プロセス:" >&2
  lsof -nP -iTCP:"$PORT" -sTCP:LISTEN >&2 || true
  echo "既存サーバが起動済みなら http://127.0.0.1:$PORT を開いてください。" >&2
  exit 1
fi

echo "記録ビューアを起動します: http://127.0.0.1:$PORT"
exec "$PYTHON" "$HERE/server.py"
