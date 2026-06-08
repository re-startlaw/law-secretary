""".cursor/shell_path_utf8.txt の各.docxの本文を抽出して標準出力に表示。"""
from __future__ import annotations

import os
import sys

from docx import Document

_REPO_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
PATH_FILE = os.path.join(_REPO_ROOT, ".cursor", "shell_path_utf8.txt")


def main() -> None:
    with open(PATH_FILE, encoding="utf-8") as f:
        paths = [ln.strip() for ln in f if ln.strip()]

    for path in paths:
        if not os.path.isfile(path):
            print(f"[NOT FOUND] {path}", file=sys.stderr)
            continue
        print(f"\n{'=' * 70}")
        print(f"FILE: {os.path.basename(path)}")
        print(f"SIZE: {os.path.getsize(path)} bytes")
        print(f"MTIME: {os.path.getmtime(path)}")
        print("=" * 70)
        doc = Document(path)
        for i, p in enumerate(doc.paragraphs, start=1):
            text = p.text
            if not text:
                print(f"{i:03d}|")
                continue
            print(f"{i:03d}|{text}")


if __name__ == "__main__":
    main()
