#!/usr/bin/env python3
"""田村案件の準抗告申立書から事件番号・罪名を抽出"""
import os
import zipfile
import re

with open(".cursor/shell_path_utf8.txt", encoding="utf-8") as f:
    paths = [ln.strip() for ln in f if ln.strip()]

for p in paths:
    print(f"\n========== {os.path.basename(p)} ==========")
    try:
        if p.endswith(".docx"):
            from docx import Document
            doc = Document(p)
            for par in doc.paragraphs:
                if par.text.strip():
                    print(par.text)
        else:  # docm or other - extract document.xml directly
            with zipfile.ZipFile(p) as z:
                xml = z.read("word/document.xml").decode("utf-8", errors="replace")
                # strip XML tags
                txt = re.sub(r"<[^>]+>", "", xml)
                txt = re.sub(r"\s+\n", "\n", txt)
                # show first 4000 chars
                print(txt[:4000])
    except Exception as e:
        print("ERR:", e)
