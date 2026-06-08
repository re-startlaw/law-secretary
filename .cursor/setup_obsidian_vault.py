"""Obsidian Vault の初期構築。仕事用Googleドライブ マイドライブ配下に作成する。"""
from __future__ import annotations

from pathlib import Path

VAULT = Path(
    "/Users/kometaninaoki/Library/CloudStorage/"
    "GoogleDrive-n.kometani@re-startlaw.com/マイドライブ/Obsidian"
)

FOLDERS = [
    "00_Inbox",
    "10_事件メモ",
    "20_業務ナレッジ/判例",
    "20_業務ナレッジ/書式",
    "20_業務ナレッジ/業務ノウハウ",
    "30_日記",
    "40_タスク",
    "90_テンプレート",
]

README = """# Re-Start法律事務所 Obsidian Vault

弁護士法人Re-Start法律事務所の業務ナレッジ・事件メモ・個人タスクを集約するVault。

## フォルダ構成

| パス | 用途 |
|---|---|
| `00_Inbox/` | 未整理メモの一時受け皿。後で整理して各フォルダへ振り分ける |
| `10_事件メモ/` | 依頼者別の打合せ録・進捗メモ。サブフォルダは依頼者ごとに作成 |
| `20_業務ナレッジ/判例/` | 判例メモ・要旨・参照リンク |
| `20_業務ナレッジ/書式/` | 書面の書き方ノウハウ（書面そのものは `02_ひな形/` 配下に保存） |
| `20_業務ナレッジ/業務ノウハウ/` | 実務ノウハウ・業種別の留意点・経理ルール等 |
| `30_日記/` | デイリーノート（`YYYY-MM-DD.md` 形式） |
| `40_タスク/` | ToDo・進行中プロジェクトの管理 |
| `90_テンプレート/` | ノートのひな形（Templater/コアプラグイン用） |

## 運用ルール

### 守秘・セキュリティ
- 本VaultはGoogleドライブ「マイドライブ」内に配置（共有しない）
- 守秘情報を含むため、共有設定の追加・公開リンク発行は禁止
- 第三者と画面共有する際はVaultを閉じる

### 同期競合の防止
- 複数端末で同時に開かない（Mac→閉じる→同期完了→iPhoneで開く、の順）
- Google Drive同期完了アイコンを確認してから別端末を開く
- `.obsidian/` 配下の設定ファイルは同期競合しやすい。問題が出たら端末ごとに分離検討

### 役割分担
- **本Vault** = 業務で得た知識・案件メモ・個人の思考（"what / why"）
- **law-secretaryリポジトリ** = Claude Codeの動作ルール・コード（"how"）
- **01_事件記録/** = 確定書面・正本（事件記録の正本はObsidianに置かない）

### Claude Codeからの参照
- リポジトリ `/Users/kometaninaoki/law-secretary/obsidian-vault` から
  シンボリックリンクで本Vaultにアクセス可
- 事件メモの追記・参照はClaude Codeに依頼可能

## 命名規則

- 依頼者ノート: `10_事件メモ/{依頼者名}/00_プロファイル.md`（依頼者の基本情報）
- 打合せ録: `10_事件メモ/{依頼者名}/{YYYY-MM-DD}_打合せ録.md`
- 判例メモ: `20_業務ナレッジ/判例/{裁判所}_{判決日}_{事件名}.md`
- デイリー: `30_日記/{YYYY-MM-DD}.md`
"""

TEMPLATES = {
    "90_テンプレート/打合せ録.md": """---
date: {{date:YYYY-MM-DD}}
client:
attendees:
case_phase:
tags: [打合せ録]
---

# {{date:YYYY-MM-DD}} 打合せ録

## 出席者
-

## 議題
-

## 議論内容


## 決定事項
-

## ToDo（米谷）
- [ ]

## ToDo（依頼者）
- [ ]

## 次回予定
- 日時:
- 場所:
""",
    "90_テンプレート/依頼者プロファイル.md": """---
client_name:
client_kana:
case_type:
case_summary:
status: 受任中
opened: {{date:YYYY-MM-DD}}
tags: [依頼者プロファイル]
---

# {{client_name}}

## 基本情報
- 氏名:
- 連絡先:
- 住所:

## 案件概要
- 案件種別:
- 受任日:
- 相手方:
- 事件番号:

## 経緯


## 進捗


## 関連メモ
-

## 関連ファイル
- 事件記録: `01_事件記録/{依頼者フォルダ}/`
""",
    "90_テンプレート/判例メモ.md": """---
court:
decision_date:
case_name:
citation:
tags: [判例]
---

# {{court}} {{decision_date}} {{case_name}}

## 出典
-

## 事案の概要


## 争点


## 判旨


## 実務へのインパクト


## 関連判例
-
""",
    "90_テンプレート/デイリーノート.md": """---
date: {{date:YYYY-MM-DD}}
tags: [daily]
---

# {{date:YYYY-MM-DD dddd}}

## 今日のタスク
- [ ]

## メモ


## 振り返り

""",
}


def main() -> None:
    if not VAULT.parent.exists():
        raise SystemExit(f"親ディレクトリが存在しません: {VAULT.parent}")

    created: list[Path] = []
    skipped: list[Path] = []

    if VAULT.exists():
        print(f"既存Vault検出: {VAULT}")
    else:
        VAULT.mkdir(parents=True)
        created.append(VAULT)

    for rel in FOLDERS:
        target = VAULT / rel
        if target.exists():
            skipped.append(target)
        else:
            target.mkdir(parents=True)
            created.append(target)

    readme_path = VAULT / "README.md"
    if readme_path.exists():
        skipped.append(readme_path)
    else:
        readme_path.write_text(README, encoding="utf-8")
        created.append(readme_path)

    for rel, body in TEMPLATES.items():
        target = VAULT / rel
        if target.exists():
            skipped.append(target)
        else:
            target.write_text(body, encoding="utf-8")
            created.append(target)

    print(f"\n作成: {len(created)} 件")
    for p in created:
        print(f"  + {p.relative_to(VAULT.parent)}")
    if skipped:
        print(f"\nスキップ（既存）: {len(skipped)} 件")
        for p in skipped:
            print(f"  = {p.relative_to(VAULT.parent)}")


if __name__ == "__main__":
    main()
