# -*- coding: utf-8 -*-
"""乙号証リネーム計画の目視確認用HTMLを生成する。
各行：原ファイル名／右肩三角の手書きクロップ／読み取り内容／新ファイル名／確度。
"""
import base64
from pathlib import Path

OUT = Path(".cursor/otsu_thumbs")
REPORT = OUT / "乙_リネーム確認表.html"

# (index, 原名, 乙番号, 種類, 作成日, 陳述者/作成者, 確度, 備考)
ROWS = [
    (0,  "19920326_f上とうZE司司.pdf",        "30", "供述調書", "20260416", "木村広希", "高", "氏名:世羅博樹こと木村広希。旧名は陳述者の生年月日(19920326)を誤用"),
    (1,  "20260222_供述調書.pdf",              "32", "供述調書", "20260330", "木村広希", "高", "氏名:世羅博樹こと木村広希。旧名の日付20260222は誤り"),
    (2,  "20260223_住と迂匠司司.pdf",          "39", "供述調書", "20260430", "三井隆将", "高", "氏名:林松風こと三井隆将"),
    (3,  "20260330_供述調書.pdf",              "25", "供述調書", "20260330", "熊野正起", "高", "氏名:熊野正起(くまのまさおき)"),
    (4,  "20260331_供述調書.pdf",              "27", "供述調書", "20260331", "田村正宣", "中", "氏名:鈴木哲夫こと田村正宣(被告人本人)。三角は「27」と判読"),
    (5,  "20260413_供述調書.pdf",              "31", "供述調書", "20260413", "木村広希", "高", "氏名:世羅博樹こと木村広希"),
    (6,  "20260414_供述調書.pdf",              "19", "供述調書", "20260414", "奥田和也", "高", "氏名:奥山義英こと奥田和也(おくだかずや)"),
    (7,  "20260416_佳進．ラゴミ司司薑壽.pdf",   "16", "供述調書", "20260416", "奥田和也", "高", "氏名:奥山義英こと奥田和也"),
    (8,  "20260416_蓼はこラZiさ言司.pdf",       "17", "供述調書", "20260416", "奥田和也", "低", "氏名:奥山義英こと奥田和也。三角は「17」か「47」か不鮮明＝要確認"),
    (9,  "20260417_供述。調書.pdf",             "40", "供述調書", "20260417", "三井隆将", "高", "氏名:林松風こと三井隆将"),
    (10, "20260426_供述調書.pdf",              "34", "供述調書", "20260426", "木村広希", "高", "氏名:世羅博樹こと木村広希"),
    (11, "20260427_供述調書.pdf",              "3",  "供述調書", "20260427", "鈴木義明", "高", "氏名:鈴木義明(昭和47年生)"),
    (12, "20260604_f生う重重司司.pdf",          "6",  "供述調書", "20260430", "鈴木義明", "低", "氏名:鈴木義明。三角は「6」か「4」か不鮮明＝要確認"),
    (13, "20260604_はとうzE司司.pdf",           "10", "供述調書", "20260417", "田村正宣", "高", "氏名:鈴木哲夫こと田村正宣(被告人本人)"),
    (14, "20260604_イラ生う重重司司.pdf",       "22", "供述調書", "20260428", "木村広希", "高", "氏名:世羅博樹こと木村広希"),
    (15, "20260604_且睾、ユ〆1弧盈工叫訂剥.pdf", "??", "戸籍全部事項証明書", "????????", "大阪市住吉区長", "低", "戸籍全部事項証明書(本籍:大阪府大阪市住吉区／田村正宣・子玲菜)。三角番号・発行日が印影で不鮮明＝要確認"),
    (16, "20260604_佳と五重司司言壽.pdf",       "13", "供述調書", "20260428", "田村正宣", "低", "氏名:鈴木哲夫こと田村正宣。三角は「13」と判読だが不鮮明＝要確認"),
    (17, "20260604_参はと五zE司司三妻.pdf",     "8",  "供述調書", "20260416", "田村正宣", "高", "氏名:鈴木哲夫こと田村正宣(被告人本人)"),
    (18, "20260604_性と式zE・‘司司二言.pdf",    "20", "供述調書", "20260428", "奥田和也", "高", "氏名:奥山義英こと奥田和也"),
    (19, "20260604_産うZE司司.pdf",            "11", "供述調書", "20260417", "田村正宣", "中", "氏名:鈴木哲夫こと田村正宣。三角は「11」と判読"),
    (20, "260525_送付書.pdf",                  "-",  "送付書",   "20260525", "東京地検", "対象外", "証拠開示請求に対する回答の送付状。乙番号なし＝リネーム対象外(任意)"),
]


def b64(path):
    if not path.exists():
        return ""
    return base64.b64encode(path.read_bytes()).decode()


def new_name(otsu, kind, date, who):
    if otsu in ("-", "??") or date.startswith("?"):
        return "(要確認のため保留)"
    return f"乙{otsu} {kind} {date} ＠{who}.pdf"


conf_color = {"高": "#2e7d32", "中": "#ef6c00", "低": "#c62828", "対象外": "#757575"}

cells = []
for idx, orig, otsu, kind, date, who, conf, note in ROWS:
    tri = OUT / f"{idx:02d}_topright.png"
    img = b64(tri)
    imgtag = f'<img src="data:image/png;base64,{img}" style="max-width:320px;border:1px solid #ccc;">' if img else "(画像なし)"
    nn = new_name(otsu, kind, date, who)
    color = conf_color.get(conf, "#000")
    cells.append(f"""
    <tr>
      <td class="idx">{idx}</td>
      <td class="orig">{orig}</td>
      <td class="img">{imgtag}</td>
      <td class="meta">
        <div><b>乙番号:</b> <span class="otsu">{otsu}</span></div>
        <div><b>種類:</b> {kind}</div>
        <div><b>作成日:</b> {date}</div>
        <div><b>陳述者/作成者:</b> {who}</div>
      </td>
      <td class="newname">{nn}</td>
      <td class="conf" style="color:{color};font-weight:bold;">{conf}</td>
      <td class="note">{note}</td>
    </tr>""")

html = f"""<!DOCTYPE html>
<html lang="ja"><head><meta charset="utf-8">
<title>乙号証 リネーム確認表（田村正宣）</title>
<style>
  body {{ font-family: "Hiragino Sans","Yu Gothic",sans-serif; margin: 24px; color:#222; }}
  h1 {{ font-size: 20px; }}
  .lead {{ background:#fff8e1; border:1px solid #ffe082; padding:12px 16px; border-radius:6px; font-size:13px; line-height:1.7; }}
  table {{ border-collapse: collapse; width: 100%; margin-top:16px; }}
  th, td {{ border:1px solid #bbb; padding:8px 10px; vertical-align: middle; font-size:13px; }}
  th {{ background:#37474f; color:#fff; position:sticky; top:0; }}
  td.idx {{ text-align:center; color:#888; }}
  td.orig {{ font-family: monospace; font-size:11px; color:#555; max-width:160px; word-break:break-all; }}
  td.img {{ text-align:center; background:#fafafa; }}
  td.otsu, .otsu {{ font-size:18px; font-weight:bold; color:#1565c0; }}
  td.newname {{ font-weight:bold; font-size:14px; max-width:260px; word-break:break-all; }}
  td.conf {{ text-align:center; font-size:15px; }}
  td.note {{ font-size:11px; color:#666; max-width:220px; }}
  tr:nth-child(even) {{ background:#f7f9fa; }}
</style></head>
<body>
<h1>乙号証 ファイル名 変更プラン／目視確認表</h1>
<div class="lead">
<b>確認方法：</b>各行の「手書きの右肩三角」画像と「乙番号」が一致しているか目視でご確認ください。<br>
新ファイル名の形式 = <code>乙{{番号}} {{種類}} {{作成日YYYYMMDD}} ＠{{陳述者}}.pdf</code>（例：<code>乙8 供述調書 20260416 ＠田村正宣.pdf</code>）<br>
・<b>作成日</b>＝供述調書1枚目の「令和○年○月○日、…において…供述した」の日付。<br>
・<b>＠</b>＝供述者本人（「○○こと△△」の本名△△）。被告人本人＝田村正宣。<br>
・<span style="color:#c62828;font-weight:bold;">確度「低」(No.8,12,15,16)</span> は手書き／印影が不鮮明です。特にご確認ください。<br>
・<b>旧自動処理の失敗との違い</b>：旧版は①日付に陳述者の生年月日を誤用、②標目をOCRで文字化けさせていました。本案は1枚目を画像として目視し、両方を是正しています。<br>
・No.20「送付書」は乙番号がなくリネーム対象外（ご希望あれば実施）。
</div>
{''.join(cells)}
</table>
</body></html>"""

REPORT.write_text(html, encoding="utf-8")
print(f"saved: {REPORT}")
