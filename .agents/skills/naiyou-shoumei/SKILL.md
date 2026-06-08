---
name: naiyou-shoumei
description: 内容証明郵便（受任通知・損害賠償請求・通知書等）を雛形ファイルからコピーしてレイアウトを維持したまま作成するスキル。`02_ひな形/00_一般/内容証明.docx` をベースとして `shutil.copy2` で依頼者フォルダの `Wordファイル/` 配下にコピーし、python-docxでrunを温存しながら段落本文だけを差し替える。新規Document()でゼロから組み立てる方式は禁止（フォント・余白・行間など雛形の体裁が全て失われるため）。
---

# 内容証明郵便 作成スキル

ユーザーが「内容証明を作って」「続編の内容証明を作って」「受任通知（内容証明）を作って」等と指示した場合に発動する。

## 基本原則：雛形のコピー＋本文置換

**絶対ルール：** 内容証明docxを `Document()` の新規生成で作ってはならない。必ず雛形ファイルを`shutil.copy2`で依頼者フォルダにコピーし、python-docxで段落のrun.textだけを書き換える。

理由：
- 雛形には用紙サイズ・余白・行間・既定フォント・段落スタイルが体裁設定として保存されている
- 新規Document()ではこれらが全てWord既定値に落ちて見た目が変わる
- 内容証明は受理可否に関わる体裁要件があるため、雛形のレイアウトを保持することが必須

## 1. 入力の確定

ユーザー指示と直近会話文脈から以下を取得：
- **依頼者**（保存先＝`01_事件記録/<頭文字>_<依頼者名>/Wordファイル/`）
- **書面の種類**（受任通知・損害賠償請求・通知書・催告書 等）
- **宛先**（相手方本人 or 相手方代理人）
- **本文の骨子**（不法行為事実・要求事項・損害項目・期限）

雛形だけでは情報が不足する場合は AskUserQuestion で**最大3問**確認。慰謝料額・期限・送付方法・言語等の戦略判断は推測せず米谷弁護士に確認する。

## 2. 雛形ファイル

メイン雛形：

```
/Users/kometaninaoki/Library/CloudStorage/GoogleDrive-n.kometani@re-startlaw.com/マイドライブ/共有用/02_ひな形/00_一般/内容証明.docx
```

雛形のレイアウト（38段落・テーブルなし）：

| 段落 | 内容 | alignment |
|---|---|---|
| 0 | 日付 | LEFT |
| 1-4 | 宛先 〒・住所・氏名 | LEFT |
| 5 | （差出人） | LEFT |
| 6-10 | 差出人 〒・住所・事務所名・弁護士名 | LEFT |
| 11 | 空行 | LEFT |
| 12-16 | （相手方代理人がいる場合）相手方代理人欄 | LEFT |
| 17 | 空行 | LEFT |
| 18 | タイトル（例：「ご連絡」「受任通知及び御連絡」） | CENTER |
| 19 | 空行 | LEFT |
| 20以降 | 本文（冠省〜草々の流れ書き） | LEFT |
| 末尾 | 草々 | RIGHT |

過去事案（馬様・大規模な構成書面）では、雛形をコピーした上で本文を「第１・第２・第３…」「（１）（２）」の項目立て構成に書き換える。タイトル・差出人・宛先・末尾「草々」位置（RIGHT）等の体裁は雛形のまま温存。

## 3. ファイル作成（雛形コピー）

```python
import shutil
SRC = '/Users/kometaninaoki/Library/CloudStorage/GoogleDrive-n.kometani@re-startlaw.com/マイドライブ/共有用/02_ひな形/00_一般/内容証明.docx'
DST = '<依頼者フォルダ>/Wordファイル/<YYMMDD>電子内容証明案_<件名>.docx'
shutil.copy2(SRC, DST)
```

既存版（同名ファイル）があれば、AGENTS.md「Word/Excel修正時の必須ルール」に従い `_2`・`_3` の連番別バージョンとする。

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

段落数が雛形より増える場合：

- 既存段落を末尾から逆順に削除して必要数に調整する、あるいは
- `doc.paragraphs[i]._element.addnext(new_p._element)` で隣接挿入する

ただし**雛形末尾の「草々」**（RIGHT 揃え）は必ず保持し、本文段落はその前に挿入する。

## 5. 内容証明の文字制限

AGENTS.md `docs/task_legal_docs.md` および memory [[feedback_naiyou_shoumei_chars]] に従い、以下の禁止文字を機械的に置換する：

- ⑴⑵⑶ → （１）（２）（３）
- ①②③ → （１）（２）（３）
- Ⅰ Ⅱ Ⅲ → I II III（半角英字）
- ㎡ ㎥ ㎝ ㎏ → 平方メートル 立方メートル センチメートル キログラム 等
- ℡ → TEL／電話
- № → No.
- ㈱ ㈲ → 株式会社 有限会社（正式名称）
- ℃ → 度
- 〝 〟 〃 → 「」／同上
- ハート・絵文字・装飾矢印 → 削除または言い換え

書き換え完了後、生成docxを下記スキャンで0件チェック：

```python
import re
forbidden = re.compile(r'[⑴-⒇①-⓿㊀-㊉ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩⅰ-ⅹ㎡㎥㎝㎜㎞㎏㎎㍑℃℡№㈱㈲㈳㈴㍻㍼㍽㍾〝〟〃⁄]')
```

## 6. 体裁ルール（memory連携）

- 段落の揃え：日付RIGHT・タイトルCENTER・以上RIGHT・他LEFT（[[feedback_document_paragraph_alignment]]）。**ただし雛形に明示的なalignmentが付いていない段落（=既定LEFT）には手を加えない**。
- 番号付きリストはWord自動番号付きを使う（[[feedback_word_autonumber]]）が、内容証明では「（１）（２）」のテキスト直書きが慣行。
- 元号：令和○年表記（西暦は雛形/相手方文書に合わせる）
- 数字：金額・住所は全角数字、半角数字は引用部のみ

## 7. 完了処理

1. `.cursor/shell_path_utf8.txt` に保存先パスを書き出して `venv/bin/python scripts/path_ops.py open` で開く（[[feedback_auto_open_documents]]）
2. ユーザーに報告：
   - 反映した本文骨子
   - 空欄（◯）にした項目（金額・期限など米谷弁護士判断待ち）
   - 禁止文字スキャン結果（0件であることを明示）

## 注意点・落とし穴

- **新規Document()で雛形なしに作るのは禁止**（このスキルが存在する一番の理由）
- **`paragraph.text = ...` 直接代入は禁止**（runが壊れて書式崩壊）
- **末尾「草々」のRIGHT揃えを温存**（流れ書き内容証明の体裁）
- **タイトル中央揃え位置を温存**（雛形段落18のCENTER）
- 日本語パスを Bash 直書きしない（AGENTS.md：`.cursor/shell_path_utf8.txt` 経由）
- 修正指示が入ったら必ず `_2`・`_3` 別バージョン保存（[[feedback_document_versioning]]）
- 戦略判断（慰謝料額・期限・動画/現場検証への言及可否等）は推測せず米谷弁護士に AskUserQuestion で確認

## 関連雛形

- 受任通知書の参考：`02_ひな形/相続/受任通知書/230216受任通知（石井久枝様） (1).docx`、`230704受任通知書（立石謙太様） (1).docx`
- 馬様事件続編：`01_事件記録/ま_馬強/Wordファイル/260515電子内容証明案_提出版_名前訂正.docx`（既出体裁、構成書面型の参考）

## 関連メモリ

- [[feedback_naiyou_shoumei_chars]] — 内容証明の使用禁止文字一覧
- [[feedback_document_paragraph_alignment]] — 段落揃えルール
- [[feedback_document_versioning]] — `_2`・`_3` 別バージョン保存ルール
- [[feedback_auto_open_documents]] — 書面作成後の自動open
- [[feedback_avoid_unicode_whitespace_warning]] — 日本語パスのBash直書き禁止
