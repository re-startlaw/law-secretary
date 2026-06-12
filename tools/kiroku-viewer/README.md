# 記録ビューア（弁護革命代替）

事件記録PDFをローカルで検索・閲覧・注釈するための自作ビューア。
弁護革命（月8,000円）のコスト削減を目的とする。設計プラン: `docs/kiroku_viewer_plan.md`。

## 起動

```
bash tools/kiroku-viewer/start.sh
```

`http://127.0.0.1:8788` をブラウザで開く。ポート8788が使用中なら占有プロセスを案内する。

## 初回セットアップ

1. 依存インストール（リポジトリ直下の venv を使用）:
   ```
   venv/bin/pip install -r tools/kiroku-viewer/requirements.txt
   ```
2. 起動してブラウザで開く:
   ```
   bash tools/kiroku-viewer/start.sh
   ```
   初回は `cases.json` が無い旨の警告が表示されるが起動は続行する。
3. ブラウザの「**＋ 追加**」ボタンから事件フォルダを登録する。
   - 「フォルダを選択…」で弁護革命の `__Document__` フォルダを選択（または手入力）。
   - 「追加後すぐ索引を作成」はONにしておくと索引が自動生成される。
   - 「このMacで索引の作成・更新を許可する」は**米谷のMacのみON**。
     **佐藤さんのMacではOFF**にして登録する（2台が並行build するのを防ぐ）。

`cases.json` はマシンローカル（gitignore）なので、2台で別々に登録する。

## オフライン利用（外出先・必須設定）

- 事件フォルダ（Drive）を **オフラインアクセス有効** に設定する（Google Drive のファイル右クリック →「オフラインで使用可能にする」）。
- 一度 **機内モード** にして、フォルダとビューアが開けることを確認してから持ち出す。
- ビューアは localhost 常駐＋ローカルキャッシュ索引で動くため、ネット接続は不要。

## データ配置（事件フォルダ内＝Drive同期で共有）

- 索引: `{事件フォルダ}/_index/export/evidence_index.sqlite`（`evidence_index.py` が生成）
- 注釈: `{事件フォルダ}/_index/annotations/{user}/{sha256}.json`
- メモ・符号上書き・回転状態: `{事件フォルダ}/_index/viewer_meta/{user}.json`
- ユーザー名は `getpass.getuser()`。表示名は `cases.json` の `display_names` で変換。

## 運用上の割り切り（D13）

- **矩形注釈はマスキングではない。** 第三者提出用の黒塗りは別途墨消しツールで行う（注釈込み書き出し時に警告表示）。
- **画面からの文書削除機能は作らない。** Finder で削除し、再索引する。
- 索引対象フォルダは既存方針に従う（弁護人メモ・接見メモは検察官開示証拠と索引を分ける）。

## セキュリティ

- サーバは `127.0.0.1` のみで待受。`Host` ヘッダ検証＋更新系は `X-Kiroku-Viewer: 1` 必須（CSRF遮断）。
- 索引DBは **読み取り専用 immutable 接続**。Drive側DBへは一切書き込まない（D3）。
- PDF.js はローカル同梱（CDN参照なし・完全オフライン）。

## アーキテクチャ

```
tools/kiroku-viewer/
├── server.py             # FastAPI（API＋静的配信＋PDFストリーム）
├── requirements.txt
├── cases.example.json    # 雛形（実体 cases.json はローカル・gitignore）
├── start.sh
├── static/
│   ├── index.html / app.js / viewer.js / style.css
│   └── vendor/pdfjs/     # pdfjs-dist@4.10.38（build/cmaps/standard_fonts）
└── tests/                # pytest
```
