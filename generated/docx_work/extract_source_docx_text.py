from pathlib import Path

from docx import Document


ROOT = Path(__file__).resolve().parents[2]
src = Path((ROOT / ".codex" / "shell_path_utf8.txt").read_text(encoding="utf-8").splitlines()[0])

doc = Document(src)
for i, paragraph in enumerate(doc.paragraphs, 1):
    if paragraph.text.strip():
        print(f"{i}: {paragraph.text}")
