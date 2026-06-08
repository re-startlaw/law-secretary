"""黒塗り版PDFを開く。"""
import subprocess
from pathlib import Path

src = Path(Path(".cursor/shell_path_utf8.txt").read_text().strip().splitlines()[0])
dst = src.with_name("甲１_黒塗り.pdf")
subprocess.run(["open", str(dst)], check=True)
print(f"opened: {dst}")
