#!/usr/bin/env python3
"""田村正宣フォルダ配下の各サブフォルダを一覧表示"""
import os

with open(".cursor/shell_path_utf8.txt", encoding="utf-8") as f:
    paths = [ln.strip() for ln in f if ln.strip()]

for p in paths:
    print(f"\n=== {os.path.basename(p)} ===")
    try:
        for name in sorted(os.listdir(p)):
            if name == ".DS_Store":
                continue
            full = os.path.join(p, name)
            kind = "DIR" if os.path.isdir(full) else "   "
            print(f"  {kind} {name}")
    except Exception as e:
        print(f"  ERR {e}")
