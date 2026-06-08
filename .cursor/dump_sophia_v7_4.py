"""Dump paragraphs of v7_4 docx with index, for inspection."""
from __future__ import annotations

import os
import sys
import zipfile
from xml.etree import ElementTree as ET

DOCX = (
    "/Users/kometaninaoki/Library/CloudStorage/"
    "GoogleDrive-n.kometani@re-startlaw.com/マイドライブ/共有用/"
    "01_事件記録/え_えいすた/合弁契約(sophia)/"
    "合弁会社設立契約書_株式会社Sophia_v7_4.docx"
)

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
ET.register_namespace("w", W_NS)


def paragraph_text(p: ET.Element) -> str:
    parts: list[str] = []
    for t in p.iter(f"{{{W_NS}}}t"):
        if t.text:
            parts.append(t.text)
    return "".join(parts)


def main() -> None:
    with zipfile.ZipFile(DOCX) as z:
        with z.open("word/document.xml") as f:
            tree = ET.parse(f)
    root = tree.getroot()
    body = root.find(f"{{{W_NS}}}body")
    assert body is not None
    paragraphs = list(body.iter(f"{{{W_NS}}}p"))
    for i, p in enumerate(paragraphs):
        text = paragraph_text(p)
        # Trim very long lines for readability
        print(f"[{i:03d}] {text}")


if __name__ == "__main__":
    main()
