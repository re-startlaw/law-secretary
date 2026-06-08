"""2026/5/7 ファイル分類実行スクリプト。

スキャナーからフォルダおよび06_分類依頼直下のファイルを所定の移動先へ振り分ける。
- 馬さん（KIST事案）: ま_馬さん/06資料/
- 田村正宣（刑事）: た_田村正宣/01_選任関係/, 02_身体拘束関係/
- 事務所経費: 03_経理/工具器具備品/, 03_経理/支払手数料 (1)/
- ファムさん: ふ_ファム・ティ・フォン/
- 不要: 06_分類依頼/pre_trash/
"""
import shutil
from pathlib import Path

ROOT = Path(
    "/Users/kometaninaoki/Library/CloudStorage/"
    "GoogleDrive-n.kometani@re-startlaw.com/マイドライブ/共有用"
)
SCAN = ROOT / "06_分類依頼" / "スキャナーから"
INBOX = ROOT / "06_分類依頼"

MA = ROOT / "01_事件記録" / "ま_馬さん" / "06資料"
TAMURA_SENNIN = ROOT / "01_事件記録" / "た_田村正宣" / "01_選任関係"
TAMURA_KOURYU = ROOT / "01_事件記録" / "た_田村正宣" / "02_身体拘束関係"
KEIRI_BIHIN = ROOT / "03_経理" / "工具器具備品"
KEIRI_TESURYO = ROOT / "03_経理" / "支払手数料 (1)"
PHAM = ROOT / "01_事件記録" / "ふ_ファム・ティ・フォン"
PRE_TRASH = INBOX / "pre_trash"

# 移動先未作成のフォルダを作る
for d in [MA]:
    d.mkdir(parents=True, exist_ok=True)

# (元フォルダ, ファイル名, 移動先) のリスト
PLAN = [
    # 馬さん（KIST事案）
    (SCAN, "20260413_古目的(1)に加えて、事案発生時の背景および現場全体の状況を示すもの（中核事実十補助.pdf", MA),
    (SCAN, "20260413_本日朝、弟が体育館で9時20分からクロスカントリーの競技に参加しておりました。1時間目の.pdf", MA),
    (SCAN, "20260413_本資料は、当方が現在保有する書面陳述、学校との往復メール、面談録音およびその整理記録、動.pdf", MA),
    (SCAN, "20260414_⑥二、主要証拠抜粋（時系列）.pdf", MA),
    (SCAN, "20260414_⑥二、主要証拠抜粋（時系列）_001.pdf", MA),
    (SCAN, "20260428_CGAClaDetaiIsEnrolmentProcess.pdf", MA),
    (SCAN, "20260507.pdf", MA),
    (SCAN, "20260507_1番1号.pdf", MA),
    (SCAN, "20260507_KISTKarenDonaldGodfrey.pdf", MA),
    (SCAN, "20260507_KISTKarenDonaldGodfrey_001.pdf", MA),
    (SCAN, "20260507_RECEIPT.pdf", MA),
    (SCAN, "20260507_Stairs.pdf", MA),
    (SCAN, "20260507_●收件人@ParentsofG9AAngelinaFeiliMaK2AGeorgeManTzar.pdf", MA),
    (SCAN, "20260507_宗在鴫？単.pdf", MA),
    (SCAN, "20260507_收件人OParentsofG9AAngelinaFeiliMaK2AGeorgeManTZar.pdf", MA),
    (SCAN, "20260507_本件は、KInternationalSchoolTokyoにおいて発生した事案である。.pdf", MA),
    (SCAN, "20260507_相談カード.pdf", MA),
    # 田村正宣（刑事）
    (SCAN, "20260501_勾留通知.pdf", TAMURA_KOURYU),
    (SCAN, "20260502_弁護人選任届.pdf", TAMURA_SENNIN),
    # 事務所経費（Acerモニター → 工具器具備品）
    (SCAN, "20260502.pdf", KEIRI_BIHIN),
    (SCAN, "20260502_001.pdf", KEIRI_BIHIN),
    (SCAN, "20260502_PM161Q.pdf", KEIRI_BIHIN),
    # 経理（Anthropic → 支払手数料）
    (SCAN, "260505_Receipt-2184-1374-4017.pdf", KEIRI_TESURYO),
    # 不要（取扱説明書）
    (SCAN, "20260502_故障かな？と思ったら.pdf", PRE_TRASH),
    # 06_分類依頼直下
    (INBOX, "260502_63_配偶者の陳述書（ファム様）.docx", PHAM),
    (INBOX, "5553711356.pdf", KEIRI_TESURYO),
]


def safe_dest(dst_dir: Path, name: str) -> Path:
    """移動先で同名がある場合は (2), (3), ... を付けて衝突回避。"""
    p = dst_dir / name
    if not p.exists():
        return p
    stem = p.stem
    suf = p.suffix
    i = 2
    while True:
        cand = dst_dir / f"{stem} ({i}){suf}"
        if not cand.exists():
            return cand
        i += 1


results = []
for src_dir, name, dst_dir in PLAN:
    src = src_dir / name
    if not src.exists():
        results.append(("MISSING", name, str(dst_dir)))
        continue
    dst_dir.mkdir(parents=True, exist_ok=True)
    dst = safe_dest(dst_dir, name)
    shutil.move(str(src), str(dst))
    results.append(("OK", name, str(dst)))

for status, name, dst in results:
    print(f"[{status}] {name} -> {dst}")
