"""Obsidian Vault 第三弾シード。
- 残りの受任中依頼者プロファイル（骨格）
- システム運用・自動化ジョブの統合ノート
- 書面テンプレート（証拠開示請求書・準備書面・上申書）
- _INDEX.md の再更新
"""
from __future__ import annotations

from pathlib import Path

VAULT = Path(
    "/Users/kometaninaoki/Library/CloudStorage/"
    "GoogleDrive-n.kometani@re-startlaw.com/マイドライブ/Obsidian"
)

FILES: dict[str, str] = {}

# ===========================================================================
# 残りの受任中依頼者プロファイル（骨格）
# ===========================================================================

REMAINING_CLIENTS = [
    {
        "folder": "インバウンドサポート（旧イワキ産業）",
        "drive_folder": "い_インバウンドサポート（旧イワキ産業）",
        "case_type": "",
        "note": "学費訴訟案件と別管理。学費訴訟は [[インバウンドサポート学費訴訟/00_プロファイル]]",
    },
    {
        "folder": "KAZEMI_HOSSEIN",
        "drive_folder": "か_KAZEMI・HOSSEIN",
        "client_name": "KAZEMI HOSSEIN",
        "case_type": "",
        "note": "外国人案件。委任契約書は [[委任契約書テンプレ#翻訳文付き対訳]] の対訳形式を検討",
    },
    {
        "folder": "斉藤慎二",
        "drive_folder": "さ_斉藤慎二",
        "case_type": "",
        "note": "",
    },
    {
        "folder": "鈴木七海",
        "drive_folder": "す_鈴木七海",
        "case_type": "",
        "note": (
            "**書面テンプレートの元案件**\n"
            "本案件の `260507電子内容証明-送付.docx` が、内容証明・通知書系の段落揃え・"
            "差出人ブロックの標準テンプレートとして運用されている（差出人インデント `Emu(4140835)` 等）。\n"
            "詳細は [[Word書面の段落揃え]]。\n\n"
            "⚠️ テンプレ流用時は `w:firstLineChars` の落とし穴に注意（同ノート参照）"
        ),
    },
    {
        "folder": "大研バイオメディカル",
        "drive_folder": "だ_大研バイオメディカル",
        "case_type": "",
        "note": "",
    },
    {
        "folder": "Homewiseonline",
        "drive_folder": "は_Homewiseonline",
        "case_type": "",
        "note": "",
    },
    {
        "folder": "メディアクリエイト_古澤智一",
        "drive_folder": "ふ_メディアクリエイト・古澤智一",
        "case_type": "",
        "note": "",
    },
    {
        "folder": "馬強",
        "drive_folder": "ま_馬強",
        "case_type": "",
        "note": "外国人案件。委任契約書は [[委任契約書テンプレ#翻訳文付き対訳]] の対訳形式を検討",
    },
    {
        "folder": "望月安士",
        "drive_folder": "も_望月安士",
        "case_type": "",
        "note": "",
    },
    {
        "folder": "卢范烨（ロハンヨウ）",
        "drive_folder": "ろ_卢范烨（ロハンヨウ）",
        "case_type": "",
        "note": "外国人案件（中国系）。委任契約書は [[委任契約書テンプレ#翻訳文付き対訳]] の対訳形式を検討",
    },
]


def build_client_profile(c: dict[str, str]) -> str:
    name = c.get("client_name", c["folder"])
    case_type = c.get("case_type", "")
    note_block = ""
    if c.get("note"):
        note_block = f"\n## 運用メモ\n{c['note']}\n"
    return f"""---
client_name: {name}
case_type: {case_type}
status: 受任中
tags: [依頼者プロファイル]
---

# {name}

> **詳細未入力** — 案件概要・事件番号・経過は随時追記する

## 事件記録フォルダ（Drive）
- `01_事件記録/{c['drive_folder']}/`
{note_block}"""


for c in REMAINING_CLIENTS:
    FILES[f"10_事件メモ/{c['folder']}/00_プロファイル.md"] = build_client_profile(c)


# ===========================================================================
# システム運用・自動化ジョブ統合ノート
# ===========================================================================

FILES["20_業務ナレッジ/業務ノウハウ/システム運用_自動化ジョブ統合.md"] = """---
title: システム運用・自動化ジョブ統合ノート
last_reviewed: 2026-05-18
sources: [memory: project_launchd_jobs, project_drive_backup, project_gas_automation, project_classification_llm, project_correction_workflow]
tags: [システム運用, 自動化, launchd, GAS, バックアップ]
---

# システム運用・自動化ジョブ統合ノート

法律事務所秘書エージェントの自動化基盤。
朝のメール添付→ファイル分類→夜のバックアップまでの全体像と、各ジョブの責務・スケジュール・トラブル対応を集約。

## 全体タイムライン（1日）

| 時刻 | 実行主体 | 内容 |
|---|---|---|
| 1時間毎 | GAS | セーブGmailアタッチメント（メール添付の自動保存） |
| 6:00〜7:00 | GAS | ライトスキャンフォルダーアタッチメント（スキャンデータ処理） |
| 08:00 | launchd | `secretary.py`：自動ファイル分類→移動→シート記録→通知メール |
| 21:00（金曜） | launchd | `freee_sync.py`：freee同期処理 |
| 23:00 | launchd | `law-secretary` リポジトリをマイドライブ個人領域へバックアップ |

> 朝のGAS処理（添付保存・スキャン処理）が終わった後の8時に分類が走るので、
> 営業開始時には通知メールが届いている設計。

---

## launchd ジョブ一覧

`~/Library/LaunchAgents/` に登録されているジョブ。
**スリープ中は復帰時にキャッチアップ実行**、電源オフの間は逃す（`StartCalendarInterval` の仕様）。

| Label | 実行内容 | スケジュール | plist |
|---|---|---|---|
| `com.restart-law.classification` | `secretary.py` を実行（06_分類依頼の自動分類→移動→シート記録→通知メール） | 毎日 08:00 | `com.restart-law.classification.plist` |
| `com.kometani.lawsecretary.backup` | `scripts/backup_to_drive.sh` でlaw-secretaryをマイドライブ個人領域へrsyncミラー | 毎日 23:00 | `com.kometani.lawsecretary.backup.plist` |
| `com.restart-law.freee-sync` | `freee_sync.py` でfreee同期 | 毎週金曜 21:00 | `com.restart-law.freee-sync.plist` |

### 状態確認・運用コマンド

```bash
# 登録状態確認
launchctl list | grep -E 'restart-law|lawsecretary'

# 手動実行
launchctl start <Label>

# plist編集後の再読み込み
launchctl unload <plist> && launchctl load <plist>
```

### ログ
- 分類: `~/law-secretary/logs/classification_launchd.log` / `classification_launchd_err.log`
- バックアップ: `~/law-secretary/.sync/backup.log` / `launchd.stdout.log` / `launchd.stderr.log`

> 新規ジョブ追加時はmemoryの `project_launchd_jobs` も更新する

---

## ファイル分類ジョブ（08:00）

### 三段構え

`secretary.py` の自動分類は3段階で動作する。

1. **外形分類**（無料）
   - `rename_file` + `classify_file`（ファイル名・送信者・件名キーワード）
2. **送信者メアド学習**（無料）
   - `build_sender_email_index` が保存ログC列(送信者)→E列(移動先)から自動構築
   - `SENDER_EMAIL_TO_CLIENT` を毎回上書き補強
   - 曖昧マッピング（同一メアド→複数依頼者）は除外
3. **LLM内容分類**（有料）
   - 外形でUNKNOWN_FOLDERに落ちたファイルのみ
   - `scripts/extract_text.py` で1回だけ抽出
   - Anthropic API (claude-sonnet-4-6) で `canonical_filename` / `dest_relative_path` / `accounting_category` / `reasoning` / `confidence` を tool_use で取得

### 設定
- **API鍵**: `~/law-secretary/secrets/.env` の `ANTHROPIC_API_KEY`
- **モデル**: `claude-sonnet-4-6`（コストと精度のバランス）
- **キャッシュ**: `system` プロンプトに `cache_control: ephemeral` を付与
  - 依頼者一覧・フォルダ構造・経理勘定キーワードは静的なのでキャッシュが効く
- **低信頼の扱い**: `confidence == "low"`、不正パス、抽出失敗、OCR必要 → `None` 返却 → `分からなかった` フォルダへ
- **LLM由来の分類は保存ログL列に** `LLM推定: {reasoning}` で記録
- **コスト目安**: 1ファイル数円（米谷弁護士承認済み）

### 関連
- 分類ルール: リポジトリ `docs/task_file_classification.md`
- 実務補足: [[ファイル分類ルール_実務補足]]
- 「分からなかった」フォルダ運用: memory `feedback_unknown_folder`
- 簡易pre_trash判定の禁止: memory `feedback_classification_no_easy_pretrash`

---

## J列修正指示フロー（apply_corrections）

`secretary.py` の `apply_corrections()` がmain()の分類前に実行される。
**シートに「修正指示」を書くだけでファイルが自動再移動される半自動化フロー**。

### 検知条件
F列（確認）が未チェック **かつ** J列（修正指示）が入力済 **かつ** K列（済）が未チェック

### 処理ロジック
1. 現在のファイル位置 = `共有用/{E列}/{D列ファイル名}`
2. J列指示から目的フォルダを推定：
   - ①スラッシュを含むフルパス指示 → そのまま採用
   - ②同じ依頼者内の別サブフォルダ → fuzzy match
   - ③別依頼者名を含む → その依頼者の該当サブフォルダ
   - ④経理科目名 → `03_経理/{科目}/`
3. サブフォルダ fuzzy match の優先順位：
   - 完全一致 → 番号プレフィックス除去一致 → 部分一致 → `CRIMINAL/CIVIL_SUBFOLDER_KEYWORDS` 経由
4. 移動後：K列TRUE化、D列HYPERLINK再生成、E列更新

### J列の書き方
- 「裁判所」「弁護人」「06_裁判所手続」など、ざっくりした記載でOK
- 解釈失敗時はスキップ（print出力あり）→ J列を修正して翌日再実行

### 学習サイクル
- J列指示は[[ファイル分類ルール_実務補足]]に取り込む対象
- memory `feedback_learn_from_j_corrections` 参照

---

## バックアップジョブ（23:00）

### 仕様
- **同期元**: `~/law-secretary/` 全体
- **同期先**: `~/Library/CloudStorage/GoogleDrive-n.kometani@re-startlaw.com/マイドライブ/秘書エージェント_バックアップ/law-secretary/`
  - 共有用の**外**＝本人専用領域（佐藤さん等は参照不可）
- **方式**: rsync -a --delete
- **除外**: `.git/` / `venv/` / `__pycache__/` / `*.pyc` / `.DS_Store` / `.sync/` / `logs/` / `secretary.log`
- **対象に含めるもの**: `.env` / `secrets/` / `.backup_*/` / 依頼者名入りスクリプト
  - **Git未管理のものを保全するのが本目的**

### 手動実行
```bash
bash ~/law-secretary/scripts/backup_to_drive.sh
```

### 設計意図
PC紛失時でも事務所アカウントから復元できるようにする。
当初は15_AI教育用配下も検討したが、共有領域のため**本人専用領域**に変更（2026-04-22）。

### パス変更時の注意
`backup_to_drive.sh` の `DEST` と plist の両方を更新する必要がある。

---

## GAS（Google Apps Script）自動化

### 稼働中のスクリプト
1. **セーブGmailアタッチメント**：1時間に1回。メール添付ファイルを自動保存
2. **ライトスキャンフォルダーアタッチメント**：1日1回（6時〜7時）。スキャンデータの処理

### 履歴・トラブル対応
- **古いGASは停止済み**（2026-01-08）
- **過剰書き込み問題**：トリガーの「入れ子」設定が原因。2026-01-17で解決済み
- **スキップ条件**：「ユニフロー」を追加済み（2025-11-08）

### 編集方法
- GASはローカルから [clasp](https://github.com/google/clasp) で編集
  - `clasp clone` → ローカル編集 → `clasp push`
- 参考：memory `feedback_gas_use_clasp`

### 関連
- 06_分類依頼フォルダに入るファイルの**上流**がこれらのGASスクリプト
- ファイル分類のトラブルシュート時はまずGAS稼働状態を確認

---

## トラブルシューティング早見

| 症状 | 確認手順 |
|---|---|
| 朝の分類通知が来ない | `launchctl list \\| grep classification` → 終了コード確認 → `classification_launchd_err.log` を見る |
| 「分からなかった」が増えた | LLM分類が落ちている可能性。`secrets/.env` の `ANTHROPIC_API_KEY` 確認 |
| J列修正が反映されない | F未チェック・J入力・K未チェックの3条件満たすか確認。fuzzy matchで失敗してる可能性も（ログに print 出力） |
| Driveバックアップが古い | `.sync/backup.log` 確認 → 容量不足・Drive未マウント等の可能性 |
| GAS自動保存が止まった | GASエディタでトリガー設定・実行ログ確認 |

---

## 関連リンク

- launchd ジョブ詳細: memory `project_launchd_jobs`
- Driveバックアップ: memory `project_drive_backup`
- GAS自動化: memory `project_gas_automation`
- LLMフォールバック: memory `project_classification_llm`
- J列修正フロー: memory `project_correction_workflow`
- 基本ルール: リポジトリ `docs/task_system_management.md`
"""


# ===========================================================================
# 書面テンプレート（90_テンプレート/）
# ===========================================================================

FILES["90_テンプレート/証拠開示請求書.md"] = """---
template_for: 証拠開示請求書
related: 証拠開示請求の列挙ルール
tags: [テンプレート, 刑事, 証拠開示]
---

# 証拠開示請求書（テンプレート）

> 必ず [[証拠開示請求の列挙ルール]] を確認してから作成する。
> - 勾留状記載の本名を使用
> - 1人=1項目
> - 紙書類と録音録画媒体は別項目
> - 「添付」等の曖昧表記禁止

---

```
　　　　　　　　　　　　　　　　　　　　　　　　　令和○年○月○日

○○地方検察庁
検察官 ○○ ○○ 殿

　　　　　　　　　　　　　　　　　〒170-6012 東京都豊島区東池袋３丁目１−１
　　　　　　　　　　　　　　　　　　　　　　　　サンシャイン60 12階
　　　　　　　　　　　　　　　　　　　　　　弁護士法人Re-Start法律事務所
　　　　　　　　　　　　　　　　　　　　　　　　弁護士　米谷　尚起
　　　　　　　　　　　　　　　　　　　　　　TEL 03-6820-3815


　　　　　　　　　　　　証拠開示請求書

被疑者　○○○○（勾留状記載本名）

頭書事件について、当職は以下の証拠の開示を請求する。

第１　開示請求する証拠

１　Ａ氏（勾留状記載：○○○○、別記「△△」）の供述調書全て
２　Ａ氏の取調べに係る録音録画記録媒体
３　Ｂ氏（勾留状記載：○○○○）の供述調書全て
４　Ｂ氏の取調べに係る録音録画記録媒体
（人ごと・紙／媒体別に１項目で続ける）

第２　請求の理由

（具体的に記載）

以上
```

## チェックリスト
- [ ] 勾留状の本名を全員分確認したか
- [ ] 通称・別名がある場合は「別記」「自称」で補記したか
- [ ] 1人1項目で番号を振っているか
- [ ] 紙の書類（供述調書・上申書・被疑者ノート・メモ等）と録音録画記録媒体を別項目にしたか
- [ ] 「等」「並びに」で複数人を1項目にまとめていないか
- [ ] 「添付」等の曖昧表記を使っていないか
- [ ] 段落の揃え（日付RIGHT、タイトルCENTER、本文LEFT、以上RIGHT）を設定したか → [[Word書面の段落揃え]]
- [ ] 自動番号付きリストを使っているか → [[Wordの自動番号付きリスト]]
"""

FILES["90_テンプレート/準備書面.md"] = """---
template_for: 準備書面
tags: [テンプレート, 民事, 訴訟]
---

# 準備書面（テンプレート）

---

```
　　　　　　　　　　　　　　　　　　　　　　　　　令和○年（ワ）第○○○号
　　　　　　　　　　　　　　　　　　　　　　　　　○○請求事件
　　　　　　　　　　　　　　　　　　　　　　　　　原告　○○○○
　　　　　　　　　　　　　　　　　　　　　　　　　被告　○○○○

　　　　　　　　　　　　　準備書面（○）

　　　　　　　　　　　　　　　　　　　　　　　　　令和○年○月○日

○○地方裁判所　民事第○部　○係　御中

　　　　　　　　　　　　　　　　　〒170-6012 東京都豊島区東池袋３丁目１−１
　　　　　　　　　　　　　　　　　　　　　　　　サンシャイン60 12階
　　　　　　　　　　　　　　　　　　　　　　弁護士法人Re-Start法律事務所
　　　　　　　　　　　　　　　　　　　　　原告（被告）訴訟代理人
　　　　　　　　　　　　　　　　　　　　　　　　弁護士　米谷　尚起
　　　　　　　　　　　　　　　　　　　　　　TEL 03-6820-3815
　　　　　　　　　　　　　　　　　　　　　　FAX

第１　○○について

１　○○○○

　○○○○○○○○○○○○○○○○○○○○○○○○○○○○○。

２　○○○○

　○○○○○○○○○○○○○○○○○○○○○○○○○○○○○。

第２　○○について

（以下続く）

以上
```

## 作成チェックリスト
- [ ] 事件番号・事件名・当事者表示は正確か
- [ ] 「準備書面（○）」の通番は前回提出分と整合しているか
- [ ] 段落の揃え：日付RIGHT、タイトルCENTER、宛先・差出人LEFT、本文LEFT、以上RIGHT → [[Word書面の段落揃え]]
- [ ] 第１・第２… の項目見出しは自動番号付きスタイルか → [[Wordの自動番号付きリスト]]
- [ ] 内容証明的な機種依存記号は問題ないが、念のためPDF表示崩れに注意
- [ ] 修正時は別バージョン保存（_2, _3...） → [[書面修正の別バージョン保存ルール]]

## 関連
- 民事委任契約書: [[委任契約書テンプレ#民事]]
"""

FILES["90_テンプレート/上申書.md"] = """---
template_for: 上申書
tags: [テンプレート, 刑事, 上申]
---

# 上申書（テンプレート）

---

```
　　　　　　　　　　　　　　　　　　　　　　　　　令和○年○月○日

○○地方裁判所　刑事第○部　○係　御中
（または　○○地方検察庁　検察官　○○ ○○ 殿）

　　　　　　　　　　　　　　　　　〒170-6012 東京都豊島区東池袋３丁目１−１
　　　　　　　　　　　　　　　　　　　　　　　　サンシャイン60 12階
　　　　　　　　　　　　　　　　　　　　　　弁護士法人Re-Start法律事務所
　　　　　　　　　　　　　　　　　　　　　被告人（被疑者）弁護人
　　　　　　　　　　　　　　　　　　　　　　　　弁護士　米谷　尚起
　　　　　　　　　　　　　　　　　　　　　　TEL 03-6820-3815


　　　　　　　　　　　　　上　申　書

被告人（被疑者）　○○○○（勾留状記載本名）

頭書事件について、弁護人は下記のとおり上申する。

　　　　　　　　　　　　　　記

１　○○○○について

　○○○○○○○○○○○○○○○○○○○○○○○○○○○○○。

２　○○○○について

　○○○○○○○○○○○○○○○○○○○○○○○○○○○○○。

以上
```

## 作成チェックリスト
- [ ] 被告人（被疑者）氏名は勾留状記載の本名を使用したか
- [ ] 宛先（裁判所部係 or 検察官）は正確か
- [ ] 段落の揃え：日付RIGHT、タイトルCENTER、「記」中央、本文LEFT、以上RIGHT → [[Word書面の段落揃え]]
- [ ] 自動番号付きリストを使用しているか → [[Wordの自動番号付きリスト]]
- [ ] 修正時は別バージョン保存 → [[書面修正の別バージョン保存ルール]]
- [ ] 複数人物の証拠列挙が含まれる場合は1人1項目・紙／媒体別 → [[証拠開示請求の列挙ルール]]

## 関連
- 刑事委任契約書: [[委任契約書テンプレ#刑事]]
"""


# ===========================================================================
# 索引の再更新
# ===========================================================================

FILES["20_業務ナレッジ/業務ノウハウ/_INDEX.md"] = """---
title: 業務ナレッジ 索引
last_reviewed: 2026-05-18
tags: [索引, ナビゲーション]
---

# 業務ナレッジ 索引

このVaultに格納したナレッジと、Claude Code側memoryに残してある関連項目への対応表。

## 経理・会計

| ナレッジ | 場所 |
|---|---|
| 勘定科目・freee運用 | [[経理・勘定科目ルール]] |
| 法人設立の経緯 | memory: `project_corporate_setup` |
| 役員報酬・社会保険 | memory: `project_officer_compensation` |
| 請求書のデフォルト・件名規則 | memory: `feedback_invoice_defaults` / `feedback_invoice_subject_from_contract` |
| 請求書ワークフロー（承認なし一括） | memory: `feedback_invoice_workflow_autorun` |

## ファイル分類

| ナレッジ | 場所 |
|---|---|
| 基本ルール | リポジトリ `docs/task_file_classification.md` |
| 実務補足（J列指示等から） | [[ファイル分類ルール_実務補足]] |
| 保存ログJ列修正の運用 | memory: `project_correction_workflow` |
| J列指示は学習材料 | memory: `feedback_learn_from_j_corrections` |
| 分類は承認不要 | memory: `feedback_file_classification_autorun` |
| 送信者・件名を最初に確認 | memory: `feedback_classify_sender_first` |
| 判断不能は「分からなかった」フォルダへ | memory: `feedback_unknown_folder` |

## 書面作成

| ナレッジ | 場所 |
|---|---|
| 基本ルール | リポジトリ `docs/task_legal_docs.md` |
| 委任契約書テンプレ（民事／刑事／対訳） | [[委任契約書テンプレ]] |
| 内容証明で使えない文字 | [[内容証明郵便の文字制限]] |
| 通知書（相手方本人宛）スタイル | [[通知書（相手方本人宛）スタイル指針]] |
| Word段落の揃え（python-docx） | [[Word書面の段落揃え]] |
| Word自動番号付きリスト | [[Wordの自動番号付きリスト]] |
| 修正は別バージョン保存 | [[書面修正の別バージョン保存ルール]] |
| 証拠開示請求の列挙ルール | [[証拠開示請求の列挙ルール]] |
| 契約書＋請求書ワークフロー | memory: `project_contract_invoice_workflow` |
| 刑事契約書・整理は一気通貫 | memory: `feedback_criminal_contract_workflow_autorun` |
| 書面作成後は自動で開く | memory: `feedback_auto_open_documents` |
| 本件非該当記述はグレーアウト | memory: `feedback_no_delete_grayout` |

## 書面テンプレート（90_テンプレート/）

| テンプレ | 場所 |
|---|---|
| 打合せ録 | [[90_テンプレート/打合せ録]] |
| 依頼者プロファイル | [[90_テンプレート/依頼者プロファイル]] |
| 判例メモ | [[90_テンプレート/判例メモ]] |
| デイリーノート | [[90_テンプレート/デイリーノート]] |
| 証拠開示請求書 | [[90_テンプレート/証拠開示請求書]] |
| 準備書面 | [[90_テンプレート/準備書面]] |
| 上申書 | [[90_テンプレート/上申書]] |

## 依頼者プロファイル（受任中）

| 依頼者 | 案件種別 | ノート |
|---|---|---|
| ダンイシン | 民事 | [[10_事件メモ/ダンイシン/00_プロファイル\\|ダンイシン]] |
| ファムさん | 在留特別許可 | [[10_事件メモ/ファムさん/00_プロファイル\\|ファムさん]] |
| 大中忠生 | 刑事 | [[10_事件メモ/大中忠生/00_プロファイル\\|大中忠生]] |
| 田村正宣 | -- | [[10_事件メモ/田村正宣/00_プロファイル\\|田村正宣]] |
| 毛さんVSヒメラ | 民事 | [[10_事件メモ/毛さんVSヒメラ/00_プロファイル\\|毛さんVSヒメラ]] |
| フューチャーリーディング | 民事＋刑事 | [[10_事件メモ/フューチャーリーディング/00_プロファイル\\|フューチャーリーディング]] |
| インバウンドサポート学費訴訟 | 民事 | [[10_事件メモ/インバウンドサポート学費訴訟/00_プロファイル\\|学費訴訟]] |
| インバウンドサポート（旧イワキ産業） | -- | [[10_事件メモ/インバウンドサポート（旧イワキ産業）/00_プロファイル\\|旧イワキ]] |
| KAZEMI HOSSEIN | -- | [[10_事件メモ/KAZEMI_HOSSEIN/00_プロファイル\\|KAZEMI]] |
| 斉藤慎二 | -- | [[10_事件メモ/斉藤慎二/00_プロファイル\\|斉藤慎二]] |
| 鈴木七海 | -- | [[10_事件メモ/鈴木七海/00_プロファイル\\|鈴木七海]] |
| 大研バイオメディカル | -- | [[10_事件メモ/大研バイオメディカル/00_プロファイル\\|大研バイオ]] |
| Homewiseonline | -- | [[10_事件メモ/Homewiseonline/00_プロファイル\\|Homewise]] |
| メディアクリエイト・古澤智一 | -- | [[10_事件メモ/メディアクリエイト_古澤智一/00_プロファイル\\|メディアクリエイト]] |
| 馬強 | -- | [[10_事件メモ/馬強/00_プロファイル\\|馬強]] |
| 望月安士 | -- | [[10_事件メモ/望月安士/00_プロファイル\\|望月安士]] |
| 卢范烨（ロハンヨウ） | -- | [[10_事件メモ/卢范烨（ロハンヨウ）/00_プロファイル\\|卢范烨]] |

### 依頼者プロファイル（終了案件のうち雛形元）

| 依頼者 | ノート |
|---|---|
| 常丹丹（対訳契約書テンプレの元） | [[10_事件メモ/常丹丹/00_プロファイル\\|常丹丹]] |

> 終了案件のうち書面テンプレ等の参照元になっているもののみ Obsidian に保持。
> 残りの終了案件は `01_事件記録/ん_終了/` を直接参照する運用。

## システム運用

| 項目 | 場所 |
|---|---|
| 基本ルール | リポジトリ `docs/task_system_management.md` |
| **統合ノート（全体像・トラブル対応）** | [[システム運用_自動化ジョブ統合]] |
| launchd自動化ジョブ一覧 | memory: `project_launchd_jobs` |
| Drive自動バックアップ | memory: `project_drive_backup` |
| GAS自動化スクリプト | memory: `project_gas_automation` |
| 自動分類LLMフォールバック | memory: `project_classification_llm` |
| J列修正フロー | memory: `project_correction_workflow` |
| GASはclaspで編集 | memory: `feedback_gas_use_clasp` |
| Unicode whitespace警告回避 | memory: `feedback_avoid_unicode_whitespace_warning` |

## ワークフロー全般

| 項目 | 場所 |
|---|---|
| 「よろしく」でデイリータスク実行 | memory: `feedback_daily_task_trigger` |
| 承認削減・効率化ルール | memory: `feedback_workflow_efficiency` |
| 分類完了通知は下書きでなく実送信 | memory: `feedback_classification_email_send` |
| 保存ログD列はHYPERLINK必須 | memory: `feedback_sheet_hyperlink_required` |
| タスク難易度によるモデル使い分け | memory: `feedback_model_routing` |
| 依頼時は質問して情報補完 | memory: `feedback_ask_clarifying_questions` |
| LINE返信のフォーマットルール | memory: `feedback_line_reply_format` |
| law-secretaryのpush先 | memory: `feedback_git_push_policy` |

## 個人

| 項目 | 場所 |
|---|---|
| 家族構成 | memory: `user_family` |
| シャツサイズ | memory: `user_shirt_size` |
| 相談フォームURL | memory: `reference_consultation_form` |
| 法人口座（振込先） | memory: `reference_corporate_bank_account` |

---

> **memory vs Obsidianの使い分け**
> - **memory** = Claude Codeが自動参照する機械可読の判断基準
> - **Obsidian** = 人間が読み返す・蓄積する知識層
> 両者は並行更新する。Obsidianの内容を変えたら、対応するmemoryも見直す。
"""


def main() -> None:
    if not VAULT.exists():
        raise SystemExit(f"Vault が存在しません: {VAULT}")

    created: list[Path] = []
    overwritten: list[Path] = []

    for rel, body in FILES.items():
        target = VAULT / rel
        target.parent.mkdir(parents=True, exist_ok=True)
        if target.exists():
            overwritten.append(target)
        else:
            created.append(target)
        target.write_text(body, encoding="utf-8")

    print(f"作成: {len(created)} 件")
    for p in created:
        print(f"  + {p.relative_to(VAULT.parent)}")
    if overwritten:
        print(f"\n上書き: {len(overwritten)} 件")
        for p in overwritten:
            print(f"  ~ {p.relative_to(VAULT.parent)}")


if __name__ == "__main__":
    main()
