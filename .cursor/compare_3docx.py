"""Extract plain text from up to 3 .docx files for comparison.

Reads the source paths from shell_path_utf8.txt (one path per line).
Outputs each file's text separated by clear markers.
"""
import os
import subprocess
import sys
import tempfile

src_file = os.path.join(os.path.dirname(__file__), 'shell_path_utf8.txt')
with open(src_file, encoding='utf-8') as f:
    paths = [line.rstrip('\n') for line in f if line.strip()]

for i, src in enumerate(paths, start=1):
    with tempfile.NamedTemporaryFile(suffix='.txt', delete=False) as tmp:
        out_path = tmp.name
    result = subprocess.run(
        ['textutil', '-convert', 'txt', '-encoding', 'UTF-8', '-output', out_path, src],
        capture_output=True, text=True
    )
    print(f'\n===== FILE {i}: {os.path.basename(src)} =====\n')
    if result.returncode != 0:
        print('STDERR:', result.stderr, file=sys.stderr)
        continue
    with open(out_path, encoding='utf-8') as f:
        print(f.read())
    os.unlink(out_path)
