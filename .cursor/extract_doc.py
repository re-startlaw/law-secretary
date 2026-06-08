"""Extract text from a .doc file via textutil, reading source path from shell_path_utf8.txt."""
import subprocess
import sys
import os
import tempfile

src_file = os.path.join(os.path.dirname(__file__), 'shell_path_utf8.txt')
with open(src_file, encoding='utf-8') as f:
    src = f.readline().rstrip('\n')

with tempfile.NamedTemporaryFile(suffix='.txt', delete=False) as tmp:
    out_path = tmp.name

result = subprocess.run(
    ['textutil', '-convert', 'txt', '-encoding', 'UTF-8', '-output', out_path, src],
    capture_output=True, text=True
)
if result.returncode != 0:
    print('STDERR:', result.stderr, file=sys.stderr)
    sys.exit(result.returncode)

with open(out_path, encoding='utf-8') as f:
    print(f.read())

os.unlink(out_path)
