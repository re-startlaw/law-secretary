"""_4 の修正履歴受諾後の最終本文を抽出して出力する。"""
from docx import Document
from lxml import etree

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
NS = {"w": W[1:-1]}

PATH = (
    "/Users/kometaninaoki/Library/CloudStorage/"
    "GoogleDrive-n.kometani@re-startlaw.com/マイドライブ/"
    "共有用/01_事件記録/ま_馬さん/Wordファイル/260512電子内容証明案_4.docx"
)


def is_in_del(el):
    a = el.getparent()
    while a is not None:
        if a.tag == f"{W}del":
            return True
        a = a.getparent()
    return False


def is_paragraph_deleted(p_elem):
    pPr = p_elem.find(f"{W}pPr", NS)
    if pPr is None:
        return False
    rPr = pPr.find(f"{W}rPr", NS)
    if rPr is None:
        return False
    return rPr.find(f"{W}del", NS) is not None


def final_text(p_elem):
    out = []
    for el in p_elem.iter():
        if el.tag == f"{W}t" and not is_in_del(el):
            out.append(el.text or "")
    return "".join(out)


doc = Document(PATH)
for i, p in enumerate(doc.paragraphs):
    pe = p._element
    if is_paragraph_deleted(pe):
        print(f"{i}: [DELETED]")
        continue
    print(f"{i}: {final_text(pe)}")
