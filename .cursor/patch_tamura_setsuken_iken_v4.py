"""現在の _2.docm から (1)(2)(3) 並列＋まとめ反論への構造変更だけを適用して _4.docm を作る。

入力：260520_接見等禁止請求に対する意見書_2.docm（米谷弁護士による編集後）
出力：260520_接見等禁止請求に対する意見書_4.docm

変更箇所のみ：第３・２「検察官が示した接見等禁止維持の理由はいずれも具体的事情たり得ないこと」
の直下を、a1×3 を並べた後に aff で番号引用して反論する形に組み替える。
他のパラグラフ（米谷弁護士の編集を含む）は一切触らない。
"""

import shutil
import zipfile
from lxml import etree

SRC = (
    "/Users/kometaninaoki/Library/CloudStorage/"
    "GoogleDrive-n.kometani@re-startlaw.com/マイドライブ/共有用/"
    "01_事件記録/た_田村正宣/09_ワードファイル/"
    "260520_接見等禁止請求に対する意見書_2.docm"
)
DST = (
    "/Users/kometaninaoki/Library/CloudStorage/"
    "GoogleDrive-n.kometani@re-startlaw.com/マイドライブ/共有用/"
    "01_事件記録/た_田村正宣/09_ワードファイル/"
    "260520_接見等禁止請求に対する意見書_4.docm"
)

WNS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = f"{{{WNS}}}"


def mk_p(style_id, text):
    p = etree.Element(f"{W}p")
    pPr = etree.SubElement(p, f"{W}pPr")
    pStyle = etree.SubElement(pPr, f"{W}pStyle")
    pStyle.set(f"{W}val", style_id)
    r = etree.SubElement(p, f"{W}r")
    rPr = etree.SubElement(r, f"{W}rPr")
    rFonts = etree.SubElement(rPr, f"{W}rFonts")
    rFonts.set(f"{W}hint", "eastAsia")
    t = etree.SubElement(r, f"{W}t")
    t.text = text
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    return p


def para_text(p):
    return "".join(tt.text or "" for tt in p.iter(f"{W}t"))


def para_style(p):
    pPr = p.find(f"{W}pPr")
    if pPr is None:
        return None
    pStyle = pPr.find(f"{W}pStyle")
    return pStyle.get(f"{W}val") if pStyle is not None else None


def main():
    shutil.copy(SRC, DST)
    with zipfile.ZipFile(DST, "r") as zin:
        names = zin.namelist()
        contents = {n: zin.read(n) for n in names}

    tree = etree.fromstring(contents["word/document.xml"])
    body = tree.find(f"{W}body")
    paras = list(body.findall(f"{W}p"))

    # 範囲特定：見出し a0「検察官が示した接見等禁止維持…」 の次行（intro）の次から、
    # a0「小括」 の直前まで。
    start_idx = None  # 置換対象の最初の段落 index
    end_idx = None    # 置換対象の最後の段落の次の index（exclusive）
    for i, p in enumerate(paras):
        if para_style(p) == "a0" and "検察官が示した接見等禁止維持" in para_text(p):
            # 直後の intro (aff) はそのまま残す。その次から置換対象とする。
            start_idx = i + 2
        elif start_idx is not None and end_idx is None and para_style(p) == "a0" and "小括" in para_text(p):
            end_idx = i
            break

    if start_idx is None or end_idx is None:
        raise SystemExit(f"section bounds not found: start={start_idx} end={end_idx}")

    print(f"replacing paragraphs [{start_idx}:{end_idx}]")
    for j in range(start_idx, end_idx):
        print(f"  remove [{j}] style={para_style(paras[j])} text={para_text(paras[j])[:60]!r}")

    new_paras = [
        mk_p("a1", "　田村氏が否認ないし黙秘していること"),
        mk_p(
            "a1",
            "　内縁の妻である大宮氏の属性や暴力団との関係性等について不明な"
            "点が多く、接見や書類の授受を通じて、被疑者と共犯者や暴力団関係者との通謀等"
            "に加担する可能性が否定できないこと",
        ),
        mk_p(
            "a1",
            "　共犯者三井氏についての妻子との間の接見禁止一部解除請求につい"
            "ても、職権発動しないとの判断がなされていること",
        ),
        mk_p(
            "aff",
            "しかしながら、上記⑴ないし⑶は、いずれも、勾留に加えて接見等禁止までを"
            "必要とする「相当強度の具体的事由」（前掲京都地裁昭和４３年６月１４日決定）"
            "には到底当たらない。以下、順次反論する。",
        ),
        mk_p(
            "aff",
            "まず⑴については、被疑事実の否認ないし黙秘という正当な防御権の行使そ"
            "れ自体を不利益に評価するものに他ならず、これを接見禁止の根拠として用いるこ"
            "とは、自白を強要する違法捜査に手を貸す結果を招来するものであって、到底許さ"
            "れない。",
        ),
        mk_p(
            "aff",
            "次に⑵については、最も近しい家族である内縁の妻について「属性が分から"
            "ない」「暴力団との関係性が不明である」というのみで接見及び手紙の授受を一切"
            "認めないという結論は、あまりに乱暴である。大宮氏は本件事件とは一切無関係で"
            "あり、共犯者とされる者たちとは全員夫を介して会ったことがあるに過ぎず、連絡"
            "先を交換しているわけでもなく、田村氏を介さずにこれらの者と連絡を取る手段を"
            "持たない。大宮氏は、弁護人から「罪証隠滅行為に加担すれば、あなた自身が犯罪"
            "に問われる」との説明を受け、これらの者と連絡を取ることは決して行ってはいけ"
            "ないと十分に理解した旨述べ、夫に対してもそのようなことを行ってはいけないと"
            "強く言い聞かせると誓約している（添付・上申書、誓約書）。加えて、本件接見は"
            "警察署職員の立会いの下で行われ、手紙も検閲を経るものであるから、これらの手"
            "続的担保を潜脱して罪証隠滅が現実に行われる具体的危険性については、検察官に"
            "おいて何ら主張・疎明がなされていない。一般人との文書の授受はすべて検閲され"
            "、接見も警察署職員などの立ち会いの下に会話の内容も極めて限定されている以上"
            "、通信手段を通じて、あるいは一般人を介して、罪証の隠滅をはかろうとしても、"
            "実際には不可能である。",
        ),
        mk_p(
            "aff",
            "最後に⑶については、そもそも全く理由になっていない。⑶は、共犯者とさ"
            "れる別の被疑者である三井氏に係る妻子との接見禁止一部解除請求について裁判所"
            "が職権発動をしなかったという、本件被疑者とは別個の被疑者に関する別件の判断"
            "結果を引き合いに出すにとどまるものであり、本件被疑者である田村氏と内縁の妻"
            "である大宮氏との接見について、これを禁止し続けなければならない具体的事情を"
            "何ら基礎付けるものではない。およそ別事件における判断結果をもって本件接見等"
            "禁止維持の根拠とすること自体、罪証隠滅の「相当強度の具体的事由」を必要とす"
            "る刑事訴訟法８１条の解釈を誤るものであって、失当というほかない。",
        ),
    ]

    # 削除：start_idx から end_idx-1 まで
    for p in paras[start_idx:end_idx]:
        body.remove(p)
    # 挿入：end_idx の直前（=現 paras[end_idx]、つまり「小括」 a0 の直前）
    anchor = paras[end_idx]
    for np in new_paras:
        anchor.addprevious(np)

    new_doc_xml = etree.tostring(
        tree, xml_declaration=True, encoding="UTF-8", standalone=True
    )
    contents["word/document.xml"] = new_doc_xml

    with zipfile.ZipFile(DST, "w", zipfile.ZIP_DEFLATED) as zout:
        for n in names:
            zout.writestr(n, contents[n])
    print("WROTE", DST)


if __name__ == "__main__":
    main()
