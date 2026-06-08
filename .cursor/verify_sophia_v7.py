"""v7検証: 該当4箇所の本文・コメント・コメントマーカー数を確認。"""
from pathlib import Path
import zipfile
from docx import Document

V7 = Path("/Users/kometaninaoki/Downloads/合弁会社設立契約書_株式会社Sophia_v7.docx")

doc = Document(str(V7))
TARGETS = [
    ("第2条第2項(修正後)", "２．本会社の取締役の員数は、原則として2名以上の偶数とし"),
    ("第2条第4項(修正後)", "４．甲及び乙は、自らが指名した取締役についてのみ"),
    ("第6条第1項(修正後)", "１．本会社における以下の各号に定める事項の決定（以下「重要事項」"),
    ("第10条第2項(修正後)", "２．配当を行うか否か、及びその配当額については、甲及び乙が協議の上、株主総会"),
]
print("=" * 60)
print("【修正箇所の本文確認】")
print("=" * 60)
for label, prefix in TARGETS:
    found = False
    for p in doc.paragraphs:
        if p.text.startswith(prefix):
            print(f"\n■ {label}\n  {p.text}")
            found = True
            break
    if not found:
        print(f"\n!! {label} NOT FOUND")

# コメントマーカー数
print()
print("=" * 60)
print("【コメントマーカー検証】")
print("=" * 60)
with zipfile.ZipFile(V7) as zf:
    doc_xml = zf.read("word/document.xml").decode("utf-8")
    com_xml = zf.read("word/comments.xml").decode("utf-8")
print(f"commentRangeStart count: {doc_xml.count('commentRangeStart')}")
print(f"commentRangeEnd count:   {doc_xml.count('commentRangeEnd')}")
print(f"commentReference count:  {doc_xml.count('commentReference')}")
print(f"<w:comment id=...> count: {com_xml.count('<w:comment')}")
print()
print("【comments.xml 抜粋(先頭500字)】")
print(com_xml[:500])
