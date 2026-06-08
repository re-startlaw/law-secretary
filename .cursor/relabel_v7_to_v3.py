"""v7.docx のコメント中の「v7」を「v3」へ置換し、v7_2.docx として保存する。
（依頼者向け呼称はv3。社内管理上のファイル名はv7系列を維持。）
"""
from pathlib import Path
import shutil
import zipfile
from lxml import etree

SRC = Path("/Users/kometaninaoki/Downloads/合弁会社設立契約書_株式会社Sophia_v7.docx")
DST = Path("/Users/kometaninaoki/Downloads/合弁会社設立契約書_株式会社Sophia_v7_2.docx")

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W_T = f"{{{W_NS}}}t"

shutil.copy(SRC, DST)

# zip読み出し
TMP = DST.with_suffix(".tmp.docx")
shutil.move(DST, TMP)
with zipfile.ZipFile(TMP, "r") as zin:
    contents = {n: zin.read(n) for n in zin.namelist()}

# comments.xml 内の<w:t>テキストのみ置換
parser = etree.XMLParser(remove_blank_text=False)
root = etree.fromstring(contents["word/comments.xml"], parser)
replaced = 0
for t in root.iter(W_T):
    if t.text and "v7" in t.text:
        before = t.text
        t.text = t.text.replace("v7", "v3")
        replaced += 1
        print(f"[replace] '{before[:30]}...' -> '{t.text[:30]}...'")
contents["word/comments.xml"] = etree.tostring(
    root, xml_declaration=True, encoding="UTF-8", standalone=True
)

with zipfile.ZipFile(DST, "w", compression=zipfile.ZIP_DEFLATED) as zout:
    for n, data in contents.items():
        zout.writestr(n, data)

TMP.unlink(missing_ok=True)
print(f"[done] replaced={replaced} -> {DST}")
