from pathlib import Path
d = Path("/Users/kometaninaoki/Library/CloudStorage/GoogleDrive-n.kometani@re-startlaw.com/マイドライブ/共有用/01_事件記録/た_田村正宣/05_検察官証拠/乙")
for f in sorted(d.glob("*.pdf")):
    print(f.name)
