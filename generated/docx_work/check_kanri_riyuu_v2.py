from pathlib import Path

from docx import Document


ROOT = Path(__file__).resolve().parents[2]
DOCX = ROOT / "監理措置決定を希望する理由_2.docx"
NG_PHRASES = [
    "確認することができます",
    "明らかにすることができます",
    "おそれは低い",
]

text = "\n".join(p.text for p in Document(DOCX).paragraphs)
for phrase in NG_PHRASES:
    print(f"{phrase}: {text.count(phrase)}")
