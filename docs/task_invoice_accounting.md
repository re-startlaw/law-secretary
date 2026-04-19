# 請求書作成・経理ルール

## 対象業務
- freee請求書の作成
- 取引先への送付
- freee経理連携（PDF取引登録・紐付け）

## 事業区分
- 法人: `corporate`
- 個人: `personal`
- 指示がない場合はユーザーへ確認してから実行

## freee請求書作成フロー
1. 事業区分（法人/個人）を確認
2. 取引先名・請求内容・金額・期日を確認
3. 請求書freee API（`/iv/invoices`）で請求書を作成
4. スクリプトが返すfreee UIのURLを米谷弁護士に報告
5. 米谷弁護士がUIで「URL共有で送付」を1クリックして送付
   （※ メール送信APIは請求書freeeに存在しないため自動化不可）

## 実行コマンド
- 自然文起点: `venv/bin/python instruction_router.py "請求書を作成して"`
- 直接起点（単一品目）:
  `venv/bin/python freee_invoice.py create --business corporate --partner-name "取引先名" --title "件名" --amount 50000 --issue-date 2026-04-17 --due-date 2026-04-30`
- 直接起点（複数品目）:
  `venv/bin/python freee_invoice.py create --business corporate --partner-name "取引先名" --title "件名" --issue-date 2026-04-17 --due-date 2026-04-30 --items '[{"description":"項目A","amount":400000},{"description":"項目B","amount":100000}]'`
- `--amount` と `--items` の金額はいずれも**税別**（`tax_entry_method=out`）

## freee経理連携
- 実行: `venv/bin/python freee_sync.py`
- 入力: `06_分類依頼/freee/` にPDFを配置
- 処理済み: `06_分類依頼/freee/処理済み/`
- 要確認: `06_分類依頼/freee/要確認/`
- ログ: `logs/freee_sync_YYYYMMDD.log`

## 認証
- `.env` に OAuth2 トークンを保存（自動リフレッシュ）
- トークン有効期限6時間、期限5分前に自動更新

## 注意事項
- 金額・期日・送信先メールは実行前に必ず復唱確認する
- 事業区分の誤り（法人/個人違い）を最優先で防ぐ
