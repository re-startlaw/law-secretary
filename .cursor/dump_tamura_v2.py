from pathlib import Path
from docx import Document

p = Path("/Users/kometaninaoki/Library/CloudStorage/GoogleDrive-n.kometani@re-startlaw.com/マイドライブ/共有用/01_事件記録/た_田村正宣/09_ワードファイル/260520_上申書_2.docx")
doc = Document(str(p))
for i, par in enumerate(doc.paragraphs):
    align = par.alignment
    print(f"[{i:03d}] align={align}: {par.text!r}")
