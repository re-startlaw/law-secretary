#!/usr/bin/env python3
"""docmからテキストを抽出（zipfileから直接）"""
import zipfile, re, os
with open(".cursor/shell_path_utf8.txt", encoding="utf-8") as f:
    paths = [ln.strip() for ln in f if ln.strip()]
for p in paths:
    print(f"\n========== {os.path.basename(p)} ==========")
    with zipfile.ZipFile(p) as z:
        xml = z.read("word/document.xml").decode("utf-8", errors="replace")
        # extract text between <w:t> tags
        texts = re.findall(r"<w:t[^>]*>([^<]*)</w:t>", xml)
        # join with paragraph breaks where </w:p> appears
        # simple version: just print all texts
        para_starts = [m.start() for m in re.finditer(r"</w:p>", xml)]
        # walk through xml, collect text per <w:p>
        out = []
        for para_match in re.finditer(r"<w:p[^>]*>(.*?)</w:p>", xml, re.S):
            ps = para_match.group(1)
            ts = re.findall(r"<w:t[^>]*>([^<]*)</w:t>", ps)
            line = "".join(ts)
            if line.strip():
                out.append(line)
        for line in out:
            print(line)
