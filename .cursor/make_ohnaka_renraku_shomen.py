"""大中忠生事件 類型証拠に関する連絡書面 作成スクリプト

テンプレート: 250918類型証拠開示請求書表紙_提出版.docx
出力     : 260428類型証拠に関する連絡書面.docx

修正点:
 - 表題「類型証拠開示請求書（１）」→「類型証拠に関する連絡書面」
 - 日付「令和７年９月１９日」→「令和８年４月２８日」
 - 本文ブロックを差し替え
   ・刑事訴訟法316条の15第1項に基づく定型文を削除
   ・代わりに「証拠一覧表番号1254ないし1258は類型証拠開示請求書別紙通番3に該当」旨を記載
"""

from __future__ import annotations

from pathlib import Path
import shutil
import zipfile
import re

WORD_DIR = Path(
    "/Users/kometaninaoki/Library/CloudStorage/"
    "GoogleDrive-n.kometani@re-startlaw.com/"
    ".shortcut-targets-by-id/1pEJgbHuaj4L4ovFEISOuq-0LeHGInxYl/"
    "大中 忠生 : 詐欺未遂/ワードファイル"
)
TEMPLATE = WORD_DIR / "250918類型証拠開示請求書表紙_提出版.docx"
OUTPUT = WORD_DIR / "260428類型証拠に関する連絡書面.docx"


# 連絡書面の本文。テンプレートの本文ブロック全体を 1 段落に置き換える。
NEW_BODY_TEXT = (
    "令和８年４月２３日付け証拠一覧表記載の番号１２５４ないし１２５８の証拠は"
    "いずれも、令和７年９月１９日付け類型証拠開示請求書別紙通番３の証拠に該当"
    "するので、類型証拠として開示されたい。"
)


def edit_document_xml(xml: str) -> str:
    # 1) 表題: 「類型証拠開示請求書」→「類型証拠に関する連絡書面」
    xml = xml.replace(
        "<w:t>類型証拠開示請求書</w:t>",
        "<w:t>類型証拠に関する連絡書面</w:t>",
    )
    # 2) 表題横の「（１）」を空に
    xml = xml.replace("<w:t>（１）</w:t>", "<w:t></w:t>")

    # 3) 日付: 令和７年→令和８年, 月日: ９月１９日→４月２８日
    #    テンプレートでは日付がランごとに分割されているので、それぞれを置換する。
    xml = xml.replace(
        "<w:t>令和７年</w:t>", "<w:t>令和８年</w:t>"
    )
    # 月の数字「９」を「４」に。日付段落の構造は
    #   令和７年 / ９ / 月 / １９ / 日 のラン構成。
    # 日付段落を一括で書き換えるため、段落本体を正規表現で置換する。
    date_pattern = re.compile(
        r'(<w:p [^>]*paraId="350B5CF4"[^>]*>.*?</w:p>)',
        re.DOTALL,
    )

    def _date_repl(m: re.Match[str]) -> str:
        inner = m.group(1)
        # 月: ９→４
        inner = inner.replace("<w:t>９</w:t>", "<w:t>４</w:t>")
        # 日: １９→２８
        inner = inner.replace("<w:t>１９</w:t>", "<w:t>２８</w:t>")
        return inner

    xml = date_pattern.sub(_date_repl, xml)

    # 4) 本文ブロック（刑訴法316条の15第1項…〜…捜査官の作成する供述書を全て含む。）を
    #    1 段落の連絡書面本文に差し替える。
    #    対象パラグラフの paraId は次のとおり:
    #     - 5AB36B73 「刑事訴訟法３１６条の１５第１項…」
    #     - 1815A67A 「全般的な類型証拠開示請求…」
    #     - 4FAE3901 「回答にあたっては…」
    #     - 69CDB619 「なお、用語の定義は以下のとおりである。」
    #     - 13CFDC03 「供述録取書等　：　…」
    #     - 6147915F 「捜査報告書等　：　…」
    #    これら 6 段落をまとめて 1 段落に置換する。
    body_pattern = re.compile(
        r'<w:p [^>]*paraId="5AB36B73"[^>]*>.*?</w:p>'
        r'\s*<w:p [^>]*paraId="1815A67A"[^>]*>.*?</w:p>'
        r'\s*<w:p [^>]*paraId="4FAE3901"[^>]*>.*?</w:p>'
        r'\s*<w:p [^>]*paraId="69CDB619"[^>]*>.*?</w:p>'
        r'\s*<w:p [^>]*paraId="13CFDC03"[^>]*>.*?</w:p>'
        r'\s*<w:p [^>]*paraId="6147915F"[^>]*>.*?</w:p>',
        re.DOTALL,
    )

    new_body_para = (
        '<w:p w14:paraId="5AB36B73" w14:textId="2E03BDE5">'
        '<w:pPr><w:rPr><w:rFonts w:ascii="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"/></w:rPr></w:pPr>'
        '<w:r>'
        '<w:rPr><w:rFonts w:ascii="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝" w:hint="eastAsia"/></w:rPr>'
        '<w:t xml:space="preserve">　</w:t>'
        '</w:r>'
        '<w:r>'
        '<w:rPr><w:rFonts w:ascii="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝" w:hint="eastAsia"/></w:rPr>'
        f'<w:t>{NEW_BODY_TEXT}</w:t>'
        '</w:r>'
        '</w:p>'
    )

    new_xml, n = body_pattern.subn(new_body_para, xml, count=1)
    if n != 1:
        raise RuntimeError("本文ブロックの置換に失敗しました（パターン不一致）")
    return new_xml


def build_doc() -> Path:
    if not TEMPLATE.exists():
        raise FileNotFoundError(TEMPLATE)
    if OUTPUT.exists():
        raise FileExistsError(f"既に存在します（手動で確認）: {OUTPUT}")

    # zipfile で document.xml だけを書き換えて新規 zip を生成
    with zipfile.ZipFile(TEMPLATE, "r") as zin:
        names = zin.namelist()
        with zipfile.ZipFile(OUTPUT, "w", zipfile.ZIP_DEFLATED) as zout:
            for name in names:
                data = zin.read(name)
                if name == "word/document.xml":
                    text = data.decode("utf-8")
                    text = edit_document_xml(text)
                    data = text.encode("utf-8")
                zout.writestr(name, data)
    return OUTPUT


if __name__ == "__main__":
    out = build_doc()
    print(f"created: {out}")
