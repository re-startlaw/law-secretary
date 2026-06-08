"""段落141を確認。"""
from pathlib import Path
from docx import Document

v7_2_path = Path("/Users/kometaninaoki/Downloads/合弁会社設立契約書_株式会社Sophia_v7_2.docx")

doc = Document(str(v7_2_path))

print("=" * 100)
print("【段落141の完全なテキスト】")
print("=" * 100)

para = doc.paragraphs[141]
print(f"\n[段落141]\n{para.text}")

print("\n" + "=" * 100)
print("【該当部分の抽出】")
print("=" * 100)
import re
matches = re.findall(r'第16条第(\d+)項', para.text)
print(f"段落141に含まれる「第16条第○項」の参照: {matches}")

