"""田村正宣・接見等禁止請求に対する意見書 v2（岡口テンプレ流用版）。

260404_接見禁止準抗告申立書_岡口.docm の書式・自動採番スタイルを完全踏襲し、
骨子の構成（第１〜第５）に従って本文を再構成する。
法律的議論（刑訴法81条・京都地裁S43.6.14・大阪地判S34.2.17 等）は岡口書面の文言を採用する。
出力先：260520_接見等禁止請求に対する意見書_2.docm
"""

import shutil
import zipfile
from lxml import etree

SRC = (
    "/Users/kometaninaoki/Library/CloudStorage/"
    "GoogleDrive-n.kometani@re-startlaw.com/マイドライブ/共有用/"
    "01_事件記録/た_田村正宣/09_ワードファイル/"
    "260404_接見禁止準抗告申立書_岡口.docm"
)
DST = (
    "/Users/kometaninaoki/Library/CloudStorage/"
    "GoogleDrive-n.kometani@re-startlaw.com/マイドライブ/共有用/"
    "01_事件記録/た_田村正宣/09_ワードファイル/"
    "260520_接見等禁止請求に対する意見書_3.docm"
)

WNS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = f"{{{WNS}}}"
nsmap = {"w": WNS}


def mk_p(style_id=None, text="", align=None, bold=False, sz=None):
    """Build a <w:p> element with given pStyle and one run."""
    p = etree.Element(f"{W}p")
    pPr = etree.SubElement(p, f"{W}pPr")
    if style_id:
        pStyle = etree.SubElement(pPr, f"{W}pStyle")
        pStyle.set(f"{W}val", style_id)
    if align:
        jc = etree.SubElement(pPr, f"{W}jc")
        jc.set(f"{W}val", align)
    if text:
        r = etree.SubElement(p, f"{W}r")
        rPr = etree.SubElement(r, f"{W}rPr")
        rFonts = etree.SubElement(rPr, f"{W}rFonts")
        rFonts.set(f"{W}hint", "eastAsia")
        if bold:
            etree.SubElement(rPr, f"{W}b")
        if sz:
            sz_el = etree.SubElement(rPr, f"{W}sz")
            sz_el.set(f"{W}val", str(sz))
            szCs = etree.SubElement(rPr, f"{W}szCs")
            szCs.set(f"{W}val", str(sz))
        t = etree.SubElement(r, f"{W}t")
        t.text = text
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    return p


def mk_empty(style_id="aff8"):
    return mk_p(style_id=style_id)


def build_body(orig_body):
    """Construct a new body element using paragraphs from build."""
    sectPr = orig_body.find(f"{W}sectPr")  # preserve page setup
    new_body = etree.Element(f"{W}body", nsmap=nsmap)

    paras = []

    # ===== Header block =====
    paras.append(mk_p("aff8", "監禁、強要、強盗、窃盗被疑事件"))
    paras.append(mk_empty())
    paras.append(mk_p("aff8", "被疑者　田村正宣"))
    paras.append(mk_empty())
    paras.append(mk_p("aff8", "接見等禁止請求に対する意見書", align="center", bold=True, sz=28))
    paras.append(mk_empty())
    paras.append(mk_p("aff8", "令和８年５月２０日", align="right"))
    paras.append(mk_empty())
    paras.append(mk_p("aff8", "東京地方裁判所　裁判官　殿"))
    paras.append(mk_empty())
    paras.append(mk_p("aff8", "ＦＡＸ：０３－３５８１－５６３９"))
    paras.append(mk_empty())
    paras.append(mk_p("aff8", "弁護人　米谷尚起", align="right"))
    paras.append(mk_empty())
    paras.append(mk_p("aff8", "（電話面談を希望します。）", align="right"))
    paras.append(mk_empty())
    paras.append(mk_p("aff8", "ＴＥＬ：０８０－３８０８－１５１５", align="right"))
    paras.append(mk_empty())

    # ===== 導入 =====
    paras.append(mk_p(
        "aff8",
        "上記被疑者に対する監禁、強要、強盗、窃盗被疑事件について検察官のなした"
        "接見等禁止請求については、下記の理由から、これを却下し、後記の者との接見"
        "及び手紙の授受を許す旨の裁判を求める。"
    ))
    paras.append(mk_empty())
    paras.append(mk_p("aff8", "記", align="center"))
    paras.append(mk_empty())

    # ===== 第１ 当事者及び事案の概要 =====
    paras.append(mk_p("a", "　当事者及び事案の概要"))

    paras.append(mk_p("a0", "　当事者"))
    paras.append(mk_p(
        "aff",
        "田村氏は５６歳の男性で、不動産解体業の仲介の仕事をしている。内縁の妻"
        "である大宮瑞希氏（以下、「大宮氏」という。）と、田村氏の実子で認知もしてい"
        "る３人の子供（長女・大宮向日凛、令和３年９月２９日生まれ、４歳。長男・大宮"
        "優篤、令和５年１月５日生まれ、３歳。次男・大宮護、令和８年１月２２日生まれ"
        "、生後４か月。）との５人暮らしである。本件意見書において接見及び手紙の授受"
        "を許すよう求める対象者は、上記内縁の妻である大宮氏である。"
    ))

    paras.append(mk_p("a0", "　事案の概要"))
    paras.append(mk_p(
        "aff",
        "本件は、令和８年２月２２日夜に被害者吉田雄二氏に対してなされたとされ"
        "る一連の行為について、田村氏に対する３度目の再逮捕に係るものである。同日の"
        "同一被害者に対する行為については、既に二度の逮捕を経て、いずれも監禁の事実"
        "で起訴されている（以下、起訴済みの２事件をそれぞれ「第１事件」「第２事件」"
        "、本件を「第３事件」という。）。"
    ))
    paras.append(mk_p(
        "aff",
        "このように本件は、同一機会・同一被害者に対する一連の行為を細切れにし"
        "て繰り返し再逮捕するものであり、その結果、田村氏が起訴後速やかに保釈請求を"
        "行う等の権利を実質的に行使し得ない異例の身柄拘束が継続している。接見等禁止"
        "請求の当否についても、こうした特異な経過を踏まえて極めて慎重に判断されるべ"
        "きである。"
    ))
    paras.append(mk_empty())

    # ===== 第２ 接見の必要性 =====
    paras.append(mk_p("a", "　接見の必要性"))
    paras.append(mk_p(
        "afd",
        "裁判例（大阪地判昭和３４年２月１７日）によれば、接見禁止は「被疑者を勾"
        "留していてもなお逃亡し又は罪証を隠滅すると疑うに足りる相当な理由がある場合"
        "に、同法第８０条の例外的措置としてなされるものであり、……被疑者に対する重"
        "大な心理的苦痛をもたらすものである点に鑑み、極めて慎重に、最小限度の運用に"
        "とどめるべき」とされている。"
    ))
    paras.append(mk_p(
        "afd",
        "本件は、暴力団内部の抗争であり、被害者の供述の信用性に相当の疑問が残る"
        "こと、被害者にも相当の落ち度があることが見込まれること、被害者の怪我が全治"
        "１ヶ月と重くないこと、被害金額が１００万円にも満たないこと、田村氏に前科が"
        "ないことを考え合わせると、万が一田村氏が有罪となった場合でも、執行猶予があ"
        "る程度見込まれる事案である。このような事案において、事件と関係がない家族等"
        "の人物との接見までも一律に禁止することは、上記裁判例に照らしても、田村氏の"
        "最低限の権利を奪うものであって、あまりに過酷で、未決勾留の趣旨を逸脱すると"
        "いうべきである。"
    ))

    paras.append(mk_p("a0", "　資金的、精神的援助の必要性"))
    paras.append(mk_p(
        "aff",
        "前記のとおり、大宮氏は３人の子供の子育てで働いておらず、生活費はすべて"
        "田村氏から受け取っており、家族の生活は田村氏に経済的に全面依存している。田"
        "村氏が突然逮捕されたことにより、大宮氏は生活費の工面や今後の生活設計につい"
        "て田村氏と直接話し合う必要に迫られている。"
    ))
    paras.append(mk_p(
        "aff",
        "加えて、令和８年３月３１日の第１事件における逮捕に伴い、田村氏の実名・"
        "顔写真及び自宅の地名（埼玉県志木市館）が全国的に報道され、田村氏が指定暴力"
        "団の構成員であることが幼稚園の保護者や近隣住民の知るところとなった。これに"
        "より、大宮氏と子供たちは志木市の自宅に留まることが事実上不可能となり、現在"
        "は茨城県内の大宮氏の実家に一時的に身を寄せている。新たな転居先の選定、長女"
        "及び長男の幼稚園の転園、生後４か月の次男の保育園入園手続、家計維持の方策、"
        "田村氏の知人からの経済的支援の調整等、いずれも一家の支柱である田村氏本人と"
        "直接協議しなければ決定できない切実な事項である。"
    ))
    paras.append(mk_p(
        "aff",
        "また、田村氏にとっても、生後４か月の乳児を含む幼い３人の子供と内縁の妻"
        "の安否は最大の関心事であり、家族との接見は精神的安定のために必要不可欠であ"
        "る。"
    ))

    paras.append(mk_p("a0", "　田村氏の健康上の問題"))
    paras.append(mk_p(
        "aff",
        "田村氏は糖尿病とそれから併発した白内障を患っており、日常的に投薬治療を"
        "受けている。留置施設における投薬管理や食事制限等について、家族との連携が不"
        "可欠であり、大宮氏との接見により、通院歴や服薬状況等の情報を共有し、適切な"
        "医療を受けるための環境を整える必要がある。"
    ))

    paras.append(mk_p("a0", "　防御権行使の環境保全"))
    paras.append(mk_p(
        "aff",
        "身体拘束をされている田村氏には、十分な防御活動ができる環境が保障されな"
        "ければならないが、接見禁止によって、弁護人が接見する際に弁護人自身が妻との"
        "連絡等日常生活上の連絡もしなければならない事態となり、弁護人に極めて過重な"
        "負担が加わる。"
    ))
    paras.append(mk_p(
        "aff",
        "さらに、被疑者に対する接見禁止は被疑者の供述の自由を奪い、任意性に欠け"
        "る自白を、しかも虚偽の自白をさせるという違法捜査に手を貸すことにほかならな"
        "い。"
    ))
    paras.append(mk_p(
        "aff",
        "したがって、内縁の妻である大宮氏との接見及び手紙の授受を認める必要性は"
        "極めて高い。"
    ))
    paras.append(mk_empty())

    # ===== 第３ 罪証隠滅を疑うに足りる相当な理由がないこと =====
    paras.append(mk_p("a", "　罪証隠滅を疑うに足りる相当な理由がないこと"))

    paras.append(mk_p("a0", "　はじめに"))
    paras.append(mk_p(
        "aff",
        "「罪証を隠滅すると疑うに足りる相当な理由がある」（刑事訴訟法８１条）"
        "とは、「被疑者が拘禁されていても、なお罪証を隠滅すると疑うに足りる相当強度"
        "の具体的事由が存する場合でなければならない」（京都地裁昭和４３年６月１４日"
        "決定・判例時報５２７号９０頁等）と解されるところ、本件においてかかる事由は"
        "認められない。単なる抽象的なおそれをもって接見等を禁止することは許されず、"
        "検察官の側において、勾留に加えて接見等禁止までを必要とする具体的事情を主張"
        "・疎明する必要がある。"
    ))

    paras.append(mk_p("a0", "　検察官が示した接見等禁止維持の理由はいずれも具体的事情たり得ないこと"))
    paras.append(mk_p(
        "aff",
        "弁護人は第１事件について大宮氏との接見等禁止解除を申し立てたが、検察"
        "官は、令和８年４月１０日付の求意見に対する回答において、概ね以下の理由によ"
        "り解除は相当でない旨意見を述べ、裁判所による職権発動もなされなかった。"
    ))
    paras.append(mk_p("a1", "　田村氏が否認ないし黙秘していること"))
    paras.append(mk_p(
        "a1",
        "　内縁の妻である大宮氏の属性や暴力団との関係性等について不明な"
        "点が多く、接見や書類の授受を通じて、被疑者と共犯者や暴力団関係者との通謀等"
        "に加担する可能性が否定できないこと"
    ))
    paras.append(mk_p(
        "a1",
        "　共犯者三井氏についての妻子との間の接見禁止一部解除請求につい"
        "ても、職権発動しないとの判断がなされていること"
    ))
    paras.append(mk_p(
        "aff",
        "しかしながら、上記⑴ないし⑶は、いずれも、勾留に加えて接見等禁止までを"
        "必要とする「相当強度の具体的事由」（前掲京都地裁昭和４３年６月１４日決定）"
        "には到底当たらない。以下、順次反論する。"
    ))
    paras.append(mk_p(
        "aff",
        "まず⑴については、被疑事実の否認ないし黙秘という正当な防御権の行使そ"
        "れ自体を不利益に評価するものに他ならず、これを接見禁止の根拠として用いるこ"
        "とは、自白を強要する違法捜査に手を貸す結果を招来するものであって、到底許さ"
        "れない。"
    ))
    paras.append(mk_p(
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
        "実際には不可能である。"
    ))
    paras.append(mk_p(
        "aff",
        "最後に⑶については、そもそも全く理由になっていない。⑶は、共犯者とさ"
        "れる別の被疑者である三井氏に係る妻子との接見禁止一部解除請求について裁判所"
        "が職権発動をしなかったという、本件被疑者とは別個の被疑者に関する別件の判断"
        "結果を引き合いに出すにとどまるものであり、本件被疑者である田村氏と内縁の妻"
        "である大宮氏との接見について、これを禁止し続けなければならない具体的事情を"
        "何ら基礎付けるものではない。およそ別事件における判断結果をもって本件接見等"
        "禁止維持の根拠とすること自体、罪証隠滅の「相当強度の具体的事由」を必要とす"
        "る刑事訴訟法８１条の解釈を誤るものであって、失当というほかない。"
    ))

    paras.append(mk_p("a0", "　小括"))
    paras.append(mk_p(
        "aff",
        "以上のとおり、検察官が挙げる⑴ないし⑶の各事情は、いずれも、勾留に加"
        "えて接見等禁止までを必要とする「相当強度の具体的事由」（前掲京都地裁昭和４"
        "３年６月１４日決定）たり得ない。立会接見及び手紙の検閲という手続的担保があ"
        "る以上、これを潜脱して罪証隠滅が現実に行われる具体的事情の主張・疎明がない"
        "限り、大宮氏との接見及び手紙の授受を一律に禁止することは到底許されない。"
    ))
    paras.append(mk_empty())

    # ===== 第４ 起訴後勾留中であること =====
    paras.append(mk_p("a", "　起訴後勾留中であること"))
    paras.append(mk_p(
        "afd",
        "田村氏は、既に第１事件及び第２事件について監禁の事実で起訴され、現在も"
        "起訴後勾留中であり、各事件についてそれぞれ接見等禁止決定が付されている。弁"
        "護人は近日中に保釈請求を行う予定であり、これが認められなかった場合には、令"
        "和８年５月１日付でなされた接見禁止決定に対して別途これを争う予定である。"
    ))
    paras.append(mk_p(
        "afd",
        "検察官は、同一の日に同一の被害者に対してなされたとされる一連の行為を細"
        "切れにし、三度にわたる再逮捕を繰り返している。このような身柄拘束の積み重ね"
        "方それ自体が、被疑者・被告人から保釈請求の実効的な機会を奪い、家族との接見"
        "及び手紙の授受を不可能とする結果を生じさせるものであって、これに上乗せして"
        "本件において接見等禁止を付すことは到底認められるべきではない。"
    ))
    paras.append(mk_p(
        "afd",
        "仮に本件において接見等禁止が付された場合、田村氏が大宮氏と接見し、又は"
        "手紙を授受するためには、第１事件・第２事件の起訴後勾留に伴って付されている"
        "接見禁止決定をも別途争わなければならず、家族との接触は二重三重に妨げられる"
        "こととなる。これ以上、家族のつながりを不当に断ち切ることは、被疑者・被告人"
        "の防御権の保障及び乳幼児３人を含む家族の生存・生活維持の観点からも、人道的"
        "見地からも、許されない。"
    ))
    paras.append(mk_empty())

    # ===== 第５ 結語 =====
    paras.append(mk_p("a", "　結語"))
    paras.append(mk_p(
        "afd",
        "以上のとおり、本件において大宮氏との接見及び手紙の授受を一律に禁止しな"
        "ければならない具体的事情は存在せず、他方、これを認めなければ、乳幼児３人を"
        "含む家族の生活そのものが立ち行かなくなる切迫した事情が現に存在する。"
    ))
    paras.append(mk_p(
        "afd",
        "よって、検察官のなした接見等禁止請求については、これを却下し、下記の者"
        "との接見及び手紙の授受を許す旨の裁判を求める。"
    ))
    paras.append(mk_empty())

    paras.append(mk_p("aff8", "氏　名　　大宮瑞希（生年月日：平成９年９月７日）"))
    paras.append(mk_p("aff8", "住　所　　埼玉県志木市館二丁目３番１３－６０２号"))
    paras.append(mk_p("aff8", "関　係　　内縁の妻"))
    paras.append(mk_empty())

    paras.append(mk_p("aff8", "以上", align="right"))
    paras.append(mk_empty())

    paras.append(mk_p("aff8", "添付書類（全て写し）"))
    paras.append(mk_empty())
    paras.append(mk_p("aff8", "１　上申書（大宮瑞希）"))
    paras.append(mk_p("aff8", "２　誓約書（大宮瑞希）"))
    paras.append(mk_p("aff8", "３　妻身分証"))

    for p in paras:
        new_body.append(p)
    if sectPr is not None:
        new_body.append(sectPr)
    return new_body


def main():
    shutil.copy(SRC, DST)
    # Read document.xml
    with zipfile.ZipFile(DST, "r") as zin:
        names = zin.namelist()
        contents = {n: zin.read(n) for n in names}

    tree = etree.fromstring(contents["word/document.xml"])
    body = tree.find(f"{W}body")
    new_body = build_body(body)
    # Replace body in tree
    tree.remove(body)
    tree.append(new_body)
    new_doc_xml = etree.tostring(
        tree,
        xml_declaration=True,
        encoding="UTF-8",
        standalone=True,
    )
    contents["word/document.xml"] = new_doc_xml

    # Write new zip
    with zipfile.ZipFile(DST, "w", zipfile.ZIP_DEFLATED) as zout:
        for n in names:
            zout.writestr(n, contents[n])
    print("WROTE", DST)


if __name__ == "__main__":
    main()
