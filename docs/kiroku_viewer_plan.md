# 記録ビューア（弁護革命代替）開発プラン v2（レビュー反映済み）

## Context（背景）

米谷弁護士は事件記録管理SaaS「弁護革命」（月額8,000円）のコスト削減のため自作代替を決定。要件はgrill-meで確定、実画面収録（2026-06-12・高山岩男案件）を解析済み。本プランはエンジニア視点・弁護革命開発者視点の2つのサブエージェントレビューを反映した最終版。

- **利用者**: 米谷＋佐藤さん（両者Mac）。データは共有Google Drive（CloudStorageマウント）。オフライン必須（MacBook持ち歩き）
- **必須**: OCR＋全文検索／検索とビューア一体のUI／タブ表示／付箋・マーカー・書き込み
- **対象外**: AI Agent・AI検索（後回し）・注釈の旧データ移行・見開き表示（並列2画面で代替）・画面からの削除機能（Finderで削除→再索引）・マスキング機能・表記揺れ同義語検索
- **確認時間**: 弁護士の確認は30〜45分×2回（計2時間以内）。動作検証はAIが自己完結
- **土台**: `scripts/evidence_index.py`（SQLite FTS5索引・Vision/Tesseract OCR、942行）
- 経緯記録: `obsidian-vault/60_ツール開発/記録ビューア（弁護革命代替）/進捗ログ.md`

## 再現する画面仕様（収録から特定）

### 文書DBタブ
- 符号タブ: `甲号証 / 乙号証 / 弁号証 / 訴訟書類 / 資料 / 符号無し`（＋全件）
- 表の列: `ID / 符号 / タイトル / 日付 / 作成者 / メモ / page`。全列ソート可・和暦切替・キーワード簡易絞り込み・`注釈あり`フィルタ・件数表示
- **行クリックで一覧内にインライン展開のPDFビューアが開く**（画面遷移しない）

### インラインビューア
- 注釈ツールバー: 付箋(★)・矩形・ペン・直線・テキスト・コメント・ハンド・undo/redo・ページジャンプ・ズーム・文書内検索・**ページ回転90°**
- 左サイドバー: `FILE/TXT`切替・`タブ表示`・`DL`・`注釈モード`・`↑↓`（前後文書）・`とじる`
- テキスト選択コピー可（PDF.js textLayer有効化）。「実行」「判例」ボタンは再現しない

### テキスト検索タブ（通常検索のみ）
- 本文キーワード＋`検索実行`＋`×クリア`＋絞り込み（タイトル・符号）
- 結果はページ単位スニペットカード（ハイライト付き）→クリックで該当文書・該当ページをビューアで開く
- スペース区切り＝OR検索（ヘルプ1行表示）。スニペットは正規化テキスト由来でよい（原文はビューアで確認）

## アーキテクチャ

```
tools/kiroku-viewer/
├── server.py             # FastAPI（API＋静的配信＋PDFストリーム）
├── requirements.txt      # fastapi / starlette / uvicorn をバージョン固定
├── cases.example.json    # 雛形（実体 cases.json は各Macローカル・gitignore）
├── start.sh              # bash start.sh で起動。ポート8788占有時はlsofで検知して案内
├── static/
│   ├── index.html / app.js / viewer.js / style.css   # 素のJS SPA
│   └── vendor/pdfjs/     # pdfjs-dist@4.x固定（build/pdf.mjs, pdf.worker.mjs, cmaps/全部, standard_fonts/全部）
└── README.md             # 起動・佐藤さんMac導入・運用ルール
```

### データ配置（事件フォルダ内＝Drive同期で共有）
- 索引: `{事件フォルダ}/_index/export/evidence_index.sqlite`（evidence_index.pyが生成）
- 注釈: `{事件フォルダ}/_index/annotations/{user}/{sha256}.json`
- メモ・符号上書き・丁数オフセット・回転状態: `{事件フォルダ}/_index/viewer_meta/{user}.json`（**単一共有ファイル禁止**。全ユーザー分をマージし、同一sha256・同一フィールドはupdated_atの新しい方を採用）
- ユーザー名: `getpass.getuser()`。表示名マッピングを設定ファイルに持つ（例 `{"nsato": "佐藤"}`）

## レビューで確定した必須設計判断（実装セッションは逸脱禁止）

### D1. 日本語全文検索の改修（最優先・フェーズ0）
現状の `pages_fts` はトークナイザ未指定（unicode61）で**日本語検索がほぼヒットしない**。`evidence_index.py` の `create_schema`（L377-446付近）を改修する：
1. `pages_fts` を `tokenize='trigram'` に変更（サーバ起動時に `sqlite3.sqlite_version_info >= (3,34,0)` をassert）
2. FTS投入テキストに `unicodedata.normalize('NFKC', text).lower()` を適用（表示用 `pages.text` は原文保持）。正規化は `normalize_for_search(s)` として1箇所に定義し、**検索クエリにも必ず同じ関数を適用**
3. クエリが3文字未満の場合は `pages` への正規化 `LIKE '%q%'` フォールバック（`text_norm` 列を追加）
4. スペース区切りはFTS5のORに展開（`escape_fts_query` L833を流用し各語を`"..."`で囲む）
5. ひらがな⇔カタカナ統合・異体字・和暦/西暦クエリ展開はフェーズ3（やらない）
6. pytest: 「文中の2文字語／3文字語／全角半角ゆれ」3パターンのヒットを検証

### D2. 文書識別子は全API・全永続データで sha256 に統一
- `documents.id` は再索引のたびに振り直されるため**永続キー使用禁止**（メモが別文書に付く事故になる）。`abs_path` はマシン依存のため**使用禁止**
- URLは `/api/cases/{id}/pdf/{sha256}`。実パスは `cases.jsonのpath + documents.rel_path` で解決し、解決後パスが事件フォルダ配下であることを `Path.resolve()` で検証
- reindex後はフロントに文書リスト再取得を指示

### D3. 索引DBはDrive上で直接openしない
- リクエスト時にDrive側DBのmtime/サイズを確認し、変化時のみ `~/Library/Caches/kiroku-viewer/{case_id}/` へ `.tmp`→`os.replace` でコピー
- 接続は `sqlite3.connect("file:...?mode=ro&immutable=1", uri=True)`（WAL/shm問題の回避）。Drive側DBへの書き込み・ロック取得は一切しない
- README＋先生確認①の案内に「事件フォルダはDriveで**オフラインアクセス有効**に設定」「一度機内モードで開いて確認」を明記

### D4. 符号パースの拡張（evidence_index.py の `EVIDENCE_FILE_RE` L56）
- `^[甲乙]` → `^[甲乙弁]`、枝番に「の」（甲7の2）、数字クラスに全角`０-９`を追加
- server側で自然順ソートキー（符号種別, 数値, 枝番）を生成（「甲1, 甲10, 甲2」を防ぐ）
- `訴訟書類/資料` タブは viewer_meta の手動上書きのみで振り分け（自動判定しない）

### D5. 未索引ファイルの即時表示（刑事の波状的な証拠追加に対応）
- documents一覧は「索引DB由来のPDF」＋「サーバがフォルダ走査（`os.walk`、`_index/`除外）で見つけた未索引PDF・非PDF（mp4/mov/m4a/jpg等）」をマージ
- 未索引PDFは「未索引（検索不可・閲覧可）」バッジ付きで即閲覧可能。非PDFは `kind:"media"`、クリックで open-file API
- 再索引ボタンは**フェーズ1**に含める（reindexは米谷Macからのみ実行する運用。`cases.json` の `"reindex": false` で無効化可能に）

### D6. PDF.js同梱の具体手順
- `pdfjs-dist@4.x`（バージョン固定）のtarballから `build/pdf.mjs`・`build/pdf.worker.mjs`・`cmaps/`全部・`standard_fonts/`全部 を配置（**cmaps欠落は日本語PDFで即死**）
- `getDocument({url, cMapUrl:'/static/vendor/pdfjs/cmaps/', cMapPacked:true, standardFontDataUrl:'/static/vendor/pdfjs/standard_fonts/'})` を必ず指定
- server.py冒頭で `mimetypes.add_type('text/javascript', '.mjs')`
- textLayer有効化（テキスト選択コピー用）。スキャンPDFはtextLayerが空なのでTXT表示をページ単位にして現在ページと連動

### D7. ビューアのレンダリング戦略
- ページ仮想化必須: 表示中±2ページのみcanvas保持、範囲外はプレースホルダdivで高さ確保。同時レンダリング2枚まで（renderTaskキャンセル可能に）
- 受け入れテスト: 200ページ超のPDFで末尾までスクロール・ジャンプしてもメモリが増え続けない

### D8. 注釈の座標系
- 保存は `viewport.convertToPdfPoint()` で変換した**PDFユーザー空間座標（左下原点）**。描画時は `convertToViewportPoint()` で復元。回転は `page.getViewport({scale, rotation})` 経由で吸収（自前のsin/cos禁止）
- 受け入れテスト: ズーム50/100/200%＋リロード後に注釈が同じ文字上に重なることをスクリーンショット比較

### D9. 注釈・metaの書き込み規律
- 全JSON書き込みは `.tmp`→`os.replace` のアトミック方式。注釈PUTのデバウンスは2秒以上
- ファイル列挙は厳密パターンのみ（注釈 `^[0-9a-f]{64}\.json$` 等）。Drive競合コピー（` (1).json`等）は無視してログ警告
- 楽観ロック: PUTのupdated_atが保存済みより古ければ409を返し再読込を促す
- **他ユーザーの注釈は表示のみ・選択不可**（色分け表示）。サーバ側も自ユーザー分のみPUT受付

### D10. 文書内検索はSQLiteで実装
- PDF.jsの `PDFFindController`・viewer.html は使わない。ビューア内検索ボックスは当該文書の `pages` を正規化LIKEで検索→ヒットページ一覧→ジャンプ。ページ内ハイライトはフェーズ3

### D11. localhostセキュリティ
- uvicornは `--host 127.0.0.1` 明示
- 全APIで `Host` ヘッダ検証（localhost:8788 / 127.0.0.1:8788 以外は403）
- POST/PUT系はフロントが `X-Kiroku-Viewer: 1` ヘッダを付け、無ければ403（CSRF遮断）。CORS有効化しない
- open-file APIは解決後パスの事件フォルダ配下検証＋拡張子ホワイトリスト（mp4/mov/m4a/jpg/png/pdf）

### D12. Range配信の受け入れテスト
- フェーズ0完了条件: `curl -H 'Range: bytes=0-1023'` が **206**＋Content-Range/Accept-Rangesを返すこと。FileResponseで206にならなければ自前Rangeハンドラ（seek＋StreamingResponse）を実装

### D13. 運用上の割り切り（READMEに明文化）
- 矩形注釈は**マスキングではない**。第三者提出用の黒塗りは別途墨消しツールで行う（注釈込み書き出し時に警告表示）
- 画面からの文書削除機能は作らない（Finderで削除→再索引）
- 索引対象フォルダの運用ルール（弁護人メモ・接見メモは検察官開示証拠と索引を分ける既存方針）を1段記載

## API設計（server.py）

- `GET /api/cases` — cases.json一覧
- `GET /api/cases/{id}/documents` — D2/D4/D5仕様の文書一覧（メモ・注釈有無・OCR低信頼フラグ・未索引バッジ・kind付き）
- `GET /api/cases/{id}/pdf/{sha256}` — Range対応PDFストリーム
- `GET /api/cases/{id}/text/{sha256}?page=n` — 抽出テキスト（ページ単位）
- `GET /api/cases/{id}/search?q=...&filter=...` — D1仕様の検索（snippet付き、`{sha256, page_no, snippet, evidence_no, title}`）
- `GET/PUT /api/cases/{id}/annotations/{sha256}` — 注釈（読みは全ユーザーマージ、書きは自ユーザーのみ、楽観ロック）
- `PUT /api/cases/{id}/meta/{sha256}` — メモ・符号上書き・丁数オフセット・回転状態
- `GET /api/cases/{id}/export/{sha256}?pages=1-5` — **注釈込みPDF書き出し**（pypdf＋reportlabで焼き込み。座標はD8のPDF座標をそのまま使用）
- `POST /api/cases/{id}/open-file/{sha256}` — `subprocess.run(["open", path])`（D11検証付き）
- `POST /api/cases/{id}/reindex` — evidence_index.py build をサブプロセス起動（多重起動拒否）

## 実装フェーズ

### フェーズ0: 土台（完了条件つき）
1. `evidence_index.py` 改修: D1（trigram＋NFKC＋text_norm列）＋D4（符号正規表現）→ pytest緑
2. `tools/kiroku-viewer/` scaffold、PDF.js同梱（D6）、`venv/bin/pip install`（バージョンpin）
3. テストデータ作成: `/tmp/kiroku_dev_case/` に (a) reportlab＋`UnicodeCIDFont('HeiseiMin-W3')` で日本語テキストPDF、(b) PyMuPDFで日本語を画像化した**テキスト層なしPDF**（OCR経路用）、(c) 200ページ超PDF、(d) ダミーmp4。甲・乙・弁・枝番・全角数字・符号無しの命名を網羅。`venv/bin/python scripts/evidence_index.py build --ocr` で索引生成
4. 完了条件: サーバ起動・`/api/documents` 応答・Range=206（D12）

### フェーズ1: コア（完成後に先生確認①30〜45分）
- **1-A** API完成（documents/search/pdf/text/open-file/reindex）＋pytest全緑 → コミット
- **1-B** 文書DB表UI（符号タブ・全列ソート・自然順・和暦・絞り込み・件数・未索引バッジ・OCR低信頼マーク）＋preview検証 → コミット
- **1-C** インラインビューア（仮想化レンダリング・ページ送り/ジャンプ・ズーム・回転・textLayer・FILE/TXT切替・文書内検索・DL・↑↓・とじる）＋テキスト検索タブ（スニペット→該当ページへ）＋検索0件時のOCR低信頼注記 ＋preview検証 → コミット

### フェーズ2: 注釈＋運用（完成後に先生確認②30〜45分）
1. 注釈オーバーレイ（付箋★・矩形・ペン・直線・テキスト・コメント・undo/redo・削除。D8座標・D9規律）
2. 自動保存・他ユーザー注釈の色分け表示（選択不可）・`注釈あり`フィルタ
3. メモ列インライン編集・丁数オフセット入力（ページ表示「3 / 12（45丁）」併記）
4. タブ表示（複数文書同時オープン）
5. 注釈込みPDF書き出し（export API＋墨消し警告）
6. 孤児注釈の検出（documentsに無いsha256）＋「この文書に紐付け直す」操作

### フェーズ3: 後回し（着手指示があった場合のみ）
並列2画面／AI検索／ページ内検索ハイライト（textLayer利用）／インクリメンタル索引＋夜間launchd／LaunchAgent自動起動／かな・異体字正規化／佐藤さんMacセットアップ

## 実装セッションへの遵守事項

1. **Bashコマンド行に日本語パス・U+3000を書かない**。事件フォルダパスはcases.json/設定ファイル経由のみ
2. `source` をBashに書かない。Pythonは常に `venv/bin/python`
3. 実事件フォルダ（01_事件記録・[弁護革命system]）への**書き込み禁止**。開発・検証は /tmp のテストケースのみ
4. PDF.jsはローカル同梱、CDN参照・外部リクエストを残さない（preview_networkで外部0件を確認）
5. 索引DBはビューアから読み取り専用（D3）。ユーザーデータはannotations/・viewer_meta/に分離
6. 設計判断D1〜D13からの逸脱禁止。やむを得ない場合は理由をコミットメッセージと進捗ログに記録
7. 機能単位でこまめにコミット。push先は `git push company main` のみ
8. **各フェーズ完了時に `obsidian-vault/60_ツール開発/記録ビューア（弁護革命代替）/進捗ログ.md` へ追記**（日付・実装内容・検証結果・未解決事項）

## 検証（AIの自己検証。先生の時間を使わない）

1. **pytest**: D1検索3パターン／sha256ラウンドトリップ（注釈保存→再索引相当のDB再生成→注釈が同一文書に表示）／楽観ロック409／Range206／open-fileのパス検証拒否
2. **UI（Claude Previewツール）**: preview_start→click/screenshot/snapshotで、符号タブ・ソート・行クリック展開・PDF描画（文字が読めるスクリーンショット）・検索→該当ページ・注釈の保存→リロード残存・ズーム変更後の注釈位置一致を確認
3. **オフライン近似**: preview_networkで外部リクエスト0件
4. **性能**: 200ページPDFのスクロール・ジャンプでメモリ安定
5. 完了後、起動方法（`bash tools/kiroku-viewer/start.sh`）と「Driveオフラインアクセス有効化＋機内モード確認」を先生に案内して確認①を依頼

## 完了の定義

- 先生がテストデータで「一覧→符号タブ→検索→該当ページ→付箋・マーカー→注釈込み書き出し」を弁護革命と同じ流れで操作できる
- 確認時間が①②合計90分（最大2時間）以内
- パイロット新件1案件をcases.jsonに登録して並行運用開始。2〜3ヶ月後に弁護革命解約判断
