"""Dump paragraphs of v3_提出版_2.docx for inspection."""
from __future__ import annotations

import zipfile
from xml.etree import ElementTree as ET

DOCX = (
    "/Users/kometaninaoki/Library/CloudStorage/"
    "GoogleDrive-n.kometani@re-startlaw.com/マイドライブ/共有用/"
    "01_事件記録/え_えいすた/合弁契約(sophia)/"
    "合弁会社設立契約書_株式会社Sophia_v3_提出版_2.docx"
)

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
ET.register_namespace("w", W)


def paragraph_text(p: ET.Element) -> str:
    return "".join(t.text or "" for t in p.iter(f"{{{W}}}t"))


def main() -> None:
    with zipfile.ZipFile(DOCX) as z:
        with z.open("word/document.xml") as f:
            tree = ET.parse(f)
    root = tree.getroot()
    body = root.find(f"{{{W}}}body")
    assert body is not None
    for i, p in enumerate(body.iter(f"{{{W}}}p")):
        print(f"[{i:03d}] {paragraph_text(p)}")


if __name__ == "__main__":
    main()
