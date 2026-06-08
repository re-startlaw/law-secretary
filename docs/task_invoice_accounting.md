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

## freeeファイルボックス自動アップロード（米谷尚起／Re-Start法律事務所名義の経理書類）
- **対象**: 03_経理/ 配下に分類されたファイルのうち、本文に「米谷尚起」または「Re-Start法律事務所」の名義が登場するもの。除外: 預り金・事業主貸・預け金・04_会計勉強
- **動作**: 朝の自動分類バッチ（`secretary.py`）内で `_detect_freee_target` が該当判定→`freee_filebox.upload_receipt` で法人事業所のファイルボックスへ自動アップロード
- **報告**: 同バッチの通知メール（佐藤様・米谷宛）の「佐藤様 本日のお願い」セクションに、アップロード結果（receipt_id・ファイルボックスURL）を一覧表示。佐藤さんはfreee上で仕訳登録のみ行う
- **手動アップロード**:
  `venv/bin/python freee_filebox.py upload --file <path> --issue-date YYYY-MM-DD --description "<摘要>" --business corporate`
- 失敗時は通知メールに「アップロード失敗」と記載されるので、佐藤さんが手動アップロードで対応

## 認証
- `.env` に OAuth2 トークンを保存（自動リフレッシュ）
- トークン有効期限6時間、期限5分前に自動更新

## 請求書Excel修正指示の運用（必須）
- 作成済み請求書Excel(.xlsx)への修正指示は**元ファイルを直接編集しない**
- 元ファイルと同じフォルダに`元ファイル名_2.xlsx`（既にあれば`_3`, `_4`...）としてコピーを作成し、コピー側のみ修正
- 金額訂正・宛名修正・品目追加など、規模に関わらず例外なく別バージョン保存
- 修正後は`open`で新バージョンを開く
- AGENTS.md「Word/Excel修正時の必須ルール」も参照

## 注意事項
- 金額・期日・送信先メールは実行前に必ず復唱確認する
- 事業区分の誤り（法人/個人違い）を最優先で防ぐ
