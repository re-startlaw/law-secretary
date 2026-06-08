#!/usr/bin/env python3
"""PyMuPDFでPDFのテキストを抽出（日本語OK）"""
import os
import sys
import fitz

with open(".cursor/shell_path_utf8.txt", encoding="utf-8") as f:
    paths = [ln.strip() for ln in f if ln.strip()]

for p in paths:
    print(f"\n========== {os.path.basename(p)} ==========")
    try:
        doc = fitz.open(p)
        for i, page in enumerate(doc):
            print(f"---p{i+1}---")
            print(page.get_text())
        doc.close()
    except Exception as e:
        print("ERR:", e)
