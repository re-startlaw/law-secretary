"""v7_2 内部参照チェック。"""
from pathlib import Path
from docx import Document
import re

v7_2_path = Path("/Users/kometaninaoki/Downloads/合弁会社設立契約書_株式会社Sophia_v7_2.docx")

def load_doc(path):
    return Document(str(path))

doc = load_doc(v7_2_path)

print("=" * 100)
print("【v7_2 内部参照の全一覧】")
print("=" * 100)

all_refs = {}
for i, p in enumerate(doc.paragraphs):
    text = p.text
    # 「第○条」パターン（項指定ありなし両方対応）
    matches = re.findall(r'第(\d+)条(?:第(\d+)項)?', text)
    if matches:
        para_text = text[:100] + ("..." if len(text) > 100 else "")
        print(f"\n[段落{i:03d}] {para_text!r}")
        for article, item in matches:
            print(f"  → 第{article}条 (項={item if item else 'なし'})")
            key = int(article)
            if key not in all_refs:
                all_refs[key] = 0
            all_refs[key] += 1

print(f"\n\n{'='*100}")
print("【参照先条文の統計】")
print(f"{'='*100}")
for cond_num in sorted(all_refs.keys()):
    count = all_refs[cond_num]
    print(f"第{cond_num}条 : {count}回参照")

print(f"\n参照対象となっている条文: {sorted(all_refs.keys())}")

