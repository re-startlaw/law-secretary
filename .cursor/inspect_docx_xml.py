#!/usr/bin/env python3
"""Dump document.xml and comments.xml prettified for inspection."""
from pathlib import Path
import zipfile
from lxml import etree

src = Path('/Users/kometaninaoki/Library/CloudStorage/GoogleDrive-n.kometani@re-startlaw.com/マイドライブ/共有用/01_事件記録/だ_大研バイオメディカル/260317_ジャグー株式会社/260427_【大研バイオメディカル株式会社御中】業務委託契約書_0423_米谷チェック済み_2.docx')

out_dir = Path('/Users/kometaninaoki/law-secretary/.cursor/jaguar_xml')
out_dir.mkdir(exist_ok=True)

with zipfile.ZipFile(src) as z:
    for name in z.namelist():
        if name.endswith('.xml') or name.endswith('.rels'):
            data = z.read(name)
            try:
                tree = etree.fromstring(data)
                pretty = etree.tostring(tree, pretty_print=True, encoding='unicode')
            except Exception:
                pretty = data.decode('utf-8', errors='replace')
            outfile = out_dir / name.replace('/', '__')
            outfile.write_text(pretty, encoding='utf-8')
            print(f'wrote {outfile}')
print('listing:')
with zipfile.ZipFile(src) as z:
    for n in z.namelist():
        print(' ', n)
