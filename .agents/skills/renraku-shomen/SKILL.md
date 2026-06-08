---
name: renraku-shomen
description: 通常の連絡書面（損害賠償請求書・催告書・通知書・受任通知などで内容証明郵便以外の普通郵便で送る書面）を、馬様事件由来の標準ひな型 `02_ひな形/00_一般/連絡書面.docx` をコピーしてレイアウトを維持したまま作成するスキル。Word デフォルト余白＋タイトル太字16pt＋差出人右端寄り＋本文1字下げ＋ア箇条書きhanging indent＋実口座記載済の標準書式を継承する。
---

# 通常連絡書面 作成スキル

ユーザーが「連絡書面を作って」「損害賠償請求書を作って」「催告書を作って」「（内容証明ではなく）通知書を作って」等と指示した場合に発動する。

## 基本原則：雛形コピー＋本文置換

**絶対ルール：** 通常書面docxを `Document()` の新規生成で作ってはならない。必ず雛形ファイルを`shutil.copy2`で依頼者フォルダにコピーし、python-docxで段落のrun.textだけを書き換える。

## 1. 雛形ファイル

メイン雛形：

```
/Users/kometaninaoki/Library/CloudStorage/GoogleDrive-n.kometani@re-startlaw.com/マイドライブ/共有用/02_ひな形/00_一般/連絡書面.docx
```

このひな型は、馬様事件 損害賠償請求書（2026-05-22）の最終版（_3.docx）のレイアウト規律を踏襲したもの。米谷弁護士の手作業による調整を反映済み。

## 2. ひな型のレイアウト規律

| 要素 | 揃え | 字下げ | フォント |
|---|---|---|---|
| 日付 | RIGHT | - | 既定 |
| 宛先（〒・住所・名前・先生） | LEFT | - | 既定 |
| （差出人）・差出人ブロック | LEFT | `first_line_indent ≒ 3600450 EMU`（右端寄り） | 既定 |
| タイトル | CENTER | - | **太字・16pt** |
| 見出し（第○、（○）、【○】） | LEFT | **インデント無し（左マージンから開始）** | 既定 |
| 本文 | LEFT | `left_indent = 133350 EMU`（1字）＋ `first_line_indent = 133350 EMU`（1字）→ 1行目2字下げ・2行目以降1字下げ | 既定 |
| 箇条書き ア・イ・ウ等 | LEFT | `left_indent = 400050 EMU`（3字）, `first_line_indent = -266700 EMU`（-2字）= hanging indent | 既定 |
| 表 | - | `tblInd = 210 twips`（≒1字） | 既定 |
| 末尾「以上」 | RIGHT | - | 既定 |

ポイント：
- **セクションタイトル**（第○・（○）・【○】等）は**左マージンから開始**（インデント無し）
- **本文**は `left_indent=1字 + first_line_indent=1字` で**1行目2字下げ・2行目以降1字下げ**（全角スペース先頭は不要、Word のインデント機能で字下げを実現）
- **差出人ブロック**は alignment=LEFT のまま `first_line_indent ≒ 3600450 EMU`（≒10cm）でぐっと右端寄りに配置
- **箇条書きア・イ・ウ**は hanging indent で2行目以降を「ア　」の後ろに揃える
- **表**は `tblInd = 210 twips`（≒1字）でテーブル全体をタイトルより右に配置（本文と同じ位置）
- **マージン**は Word デフォルト 2.54cm 四方
- **振込先**は雛形に実口座（ＧＭＯあおぞらネット銀行 法人第二営業部支店(102) 普通2469137 預かり口）を記載済み

**重要：本文の先頭に全角スペースを入れない。** `first_line_indent` で字下げを実現するため、全角スペース併用はNG（2文字下げと3文字下げが混在してしまう）。

## 3. ファイル作成手順

```python
import shutil
SRC = '/Users/.../02_ひな形/00_一般/連絡書面.docx'
DST = '<依頼者フォルダ>/Wordファイル/<YYMMDD>_<書面名>（<依頼者>）.docx'
shutil.copy2(SRC, DST)
```

修正指示が入ったら、AGENTS.md「Word/Excel修正時の必須ルール」と memory [[feedback_document_versioning]] に従い `_2`・`_3` の連番別バージョンとする。

## 4. 本文の書き換え（python-docx・run温存）

`paragraph.text = ...` を使うとrunが壊れて書式が崩れる。必ず以下のヘルパーでrunを温存して書き換える：

```python
def replace_paragraph(para, new_text):
    if para.runs:
        first_run = para.runs[0]
        for run in para.runs[1:]:
            run.text = ''
        first_run.text = new_text
    else:
        para.add_run(new_text)
```

段落数を増やす場合は、雛形末尾の「以上」段落（RIGHT 揃え）の前に `_element.addprevious()` で挿入する。

## 5. 表を追加する場合

```python
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

def set_table_indent(table, twips=210):
    """テーブル全体の左インデントを設定（≒1字）。"""
    tblPr = table._element.tblPr
    if tblPr is None:
        tblPr = OxmlElement('w:tblPr')
        table._element.insert(0, tblPr)
    existing = tblPr.find(qn('w:tblInd'))
    if existing is not None:
        tblPr.remove(existing)
    tblInd = OxmlElement('w:tblInd')
    tblInd.set(qn('w:w'), str(twips))
    tblInd.set(qn('w:type'), 'dxa')
    tblPr.append(tblInd)
```

セル罫線（雛形に Table Grid スタイルが無いため）：

```python
def _set_cell_borders(cell):
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    for edge in ('top', 'left', 'bottom', 'right'):
        b = OxmlElement(f'w:{edge}')
        b.set(qn('w:val'), 'single')
        b.set(qn('w:sz'), '4')
        b.set(qn('w:color'), '000000')
        tcBorders.append(b)
    existing = tcPr.find(qn('w:tcBorders'))
    if existing is not None:
        tcPr.remove(existing)
    tcPr.append(tcBorders)
```

## 6. 書面固有のルール

### 損害賠償請求書
- 遅延損害金の起算日：**不法行為発生日**（例：令和８年４月１３日）から年３％
- 表は「実費損害一覧＋総合計」程度に限定。慰謝料・弁護士費用が単一項目なら文章で「少なくとも金●円を下ることはないため、同金額を請求する。状況変化に応じて増額予定」と記載
- 「万一」「念のため」等の枕詞は使わない

### 受任通知（普通郵便）
- 内容証明を別途送付する場合は、その旨触れる
- 直接連絡を控えてもらう旨を明記

## 7. 完了処理

1. `.cursor/shell_path_utf8.txt` に保存先パスを書き出して `venv/bin/python scripts/path_ops.py open` で開く（[[feedback_auto_open_documents]]）
2. ユーザーに報告：
   - 反映した本文骨子
   - 空欄（◯）にした項目（金額・期限など米谷弁護士判断待ち）

## 注意点・落とし穴

- **新規Document()で雛形なしに作るのは禁止**
- **`paragraph.text = ...` 直接代入は禁止**（runが壊れて書式崩壊）
- **末尾「以上」のRIGHT揃え・タイトルの太字16pt・差出人の右端寄り配置を温存**
- **左マージン1字（133350 EMU）を見出し・本文に統一**（タイトル・日付・宛先・差出人・以上は除外）
- 修正指示が入ったら必ず `_2`・`_3` 別バージョン保存（[[feedback_document_versioning]]）
- 内容証明郵便で送る場合は本スキルではなく [[skill: naiyou-shoumei]] を使う

## 関連メモリ

- [[feedback_word_document_layout]] — 通常Word書面のレイアウト規律
- [[feedback_document_paragraph_alignment]] — 段落揃え
- [[feedback_document_versioning]] — _2・_3 別バージョン保存
- [[feedback_auto_open_documents]] — 書面作成後の自動open
- [[reference_corporate_bank_account]] — 振込先口座（雛形に記載済）

## 関連スキル

- [[skill: naiyou-shoumei]] — 内容証明郵便で送る書面はこちら
