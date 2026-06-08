from docx import Document
PATH = (
    "/Users/kometaninaoki/Library/CloudStorage/"
    "GoogleDrive-n.kometani@re-startlaw.com/マイドライブ/"
    "共有用/01_事件記録/ま_馬さん/Wordファイル/260512電子内容証明案_中国語版.docx"
)
doc = Document(PATH)
for i, p in enumerate(doc.paragraphs):
    print(f"{i}: [{p.style.name}|{p.alignment}] {p.text}")
