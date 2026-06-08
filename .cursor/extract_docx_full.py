#!/usr/bin/env python3
"""Extract docx body, comments, and revisions from each path in shell_path_utf8.txt."""
from __future__ import annotations
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}


def text_of(elem):
    parts = []
    for t in elem.iter("{%s}t" % NS["w"]):
        if t.text:
            parts.append(t.text)
    return "".join(parts)


def extract_paragraphs(xml_bytes):
    root = ET.fromstring(xml_bytes)
    out = []
    for p in root.iter("{%s}p" % NS["w"]):
        out.append(text_of(p))
        for d in p.iter("{%s}del" % NS["w"]):
            t = text_of(d)
            if t.strip():
                out.append("  [DEL] " + t)
        for i in p.iter("{%s}ins" % NS["w"]):
            t = text_of(i)
            if t.strip():
                out.append("  [INS] " + t)
    return out


def extract_comments(xml_bytes):
    root = ET.fromstring(xml_bytes)
    out = []
    for c in root.iter("{%s}comment" % NS["w"]):
        cid = c.attrib.get("{%s}id" % NS["w"], "")
        author = c.attrib.get("{%s}author" % NS["w"], "")
        date = c.attrib.get("{%s}date" % NS["w"], "")
        out.append((cid, author, date, text_of(c)))
    return out


def process(path):
    print("=" * 80)
    print("FILE: " + path.name)
    print("=" * 80)
    try:
        with zipfile.ZipFile(path) as z:
            names = z.namelist()
            doc = z.read("word/document.xml")
            paragraphs = extract_paragraphs(doc)
            print("\n--- BODY ---")
            for p in paragraphs:
                print(p)
            if "word/comments.xml" in names:
                print("\n--- COMMENTS ---")
                cdata = z.read("word/comments.xml")
                for cid, author, date, text in extract_comments(cdata):
                    print("[#%s %s %s] %s" % (cid, author, date, text))
            else:
                print("\n--- COMMENTS --- (none)")
    except Exception as e:
        print("ERROR: %s" % e, file=sys.stderr)


def main():
    list_file = Path(__file__).resolve().parent / "shell_path_utf8.txt"
    paths = [Path(line.strip()) for line in list_file.read_text(encoding="utf-8").splitlines() if line.strip()]
    for p in paths:
        process(p)


if __name__ == "__main__":
    main()
