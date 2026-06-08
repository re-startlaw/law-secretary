# 裁判記録インデックス管理

## 基本方針

- 正本は開示PDFそのもの。索引、OCRテキスト、時系列表は二次資料。
- PDF全文を毎回LLMへ渡さず、ローカル索引で候補ページを絞ってから必要部分だけ使う。
- NotebookLMは全体把握・探索用の補助に限定し、証拠台帳や時系列表の正本にはしない。

## 作成される成果物

`scripts/evidence_index.py build` は、対象フォルダ直下に `_index/export/` を作り、以下を出力する。

- `evidence_index.sqlite`: SQLite FTS5 のページ単位検索DB
- `manifest.csv`: PDFパス、証拠番号、標目、sha256、ページ数、抽出方式などの管理表
- `timeline.csv`: 日付候補を含む時系列候補表
- `timeline.md`: 閲覧用の時系列候補表
- `llm_usage_log.csv`: Codex、Claude Code、NotebookLM等への投入ログ
- `README.md`: その案件フォルダ内での運用注意

SQLiteは一度ローカル一時領域で構築し、完了後にDrive側へスナップショットとしてコピーする。

## 実行例

対象パスは `.codex/shell_path_utf8.txt` の1行目に書く。

```bash
venv/bin/python scripts/evidence_index.py build
```

OCRも試す場合:

```bash
venv/bin/python scripts/evidence_index.py build --ocr
```

検索:

```bash
venv/bin/python scripts/evidence_index.py search "検索語"
```

LLM/NotebookLM投入ログ:

```bash
venv/bin/python scripts/evidence_index.py log-llm \
  --tool-or-service Codex \
  --purpose "準備書面用の争点整理" \
  --target-files-or-pages "甲1 p.3, 乙8 p.2" \
  --input-chars 4200 \
  --external-ai-used yes \
  --confirmed-by "米谷尚起" \
  --confirmation-note "対象ページと目的を確認済み"
```

## 運用ルール

- 索引DB、OCRテキスト、時系列表、LLM投入用抽出物も開示証拠由来資料として扱う。
- LLM/NotebookLMの出力は参考メモにとどめ、書面・尋問・証拠意見に使う前に必ず原PDF画像で確認する。
- `timeline.*` の `review_status` が `原PDF目視確認済み` でない行は未確認候補。
- 外部AI投入前には、対象ファイル、個人情報、利用目的、投入先、削除予定を明示して確認を取る。
- 検察官開示証拠、弁護人メモ、接見メモ、反対尋問メモ、公開資料は索引・保存場所を分ける。

