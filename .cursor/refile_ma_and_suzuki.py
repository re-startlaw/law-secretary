"""1番1号.pdfを鈴木七海さんへ移動、馬さん06資料をリネームして260507受領サブフォルダへ整理する。

実施内容:
- 20260507_1番1号.pdf -> 01_事件記録/す_鈴木七海/03連絡文書/1番1号書留通知（封筒）.pdf
- 馬さんの06資料配下16ファイルを「260507受領」サブフォルダにリネームして移動
"""
import shutil
from pathlib import Path

ROOT = Path(
    "/Users/kometaninaoki/Library/CloudStorage/"
    "GoogleDrive-n.kometani@re-startlaw.com/マイドライブ/共有用"
)

MA_RES = ROOT / "01_事件記録" / "ま_馬さん" / "06資料"
MA_RECEIVED = MA_RES / "260507受領"
SUZUKI = ROOT / "01_事件記録" / "す_鈴木七海" / "03連絡文書"

# (現ファイル名（馬さん06資料配下）, 新ファイル名)
RENAMES = [
    ("20260413_古目的(1)に加えて、事案発生時の背景および現場全体の状況を示すもの（中核事実十補助.pdf",
     "証拠一覧と各証拠の目的説明.pdf"),
    ("20260413_本日朝、弟が体育館で9時20分からクロスカントリーの競技に参加しておりました。1時間目の.pdf",
     "Angelina書面陳述（4月13日付）.pdf"),
    ("20260413_本資料は、当方が現在保有する書面陳述、学校との往復メール、面談録音およびその整理記録、動.pdf",
     "事案整理書面（依頼者作成）.pdf"),
    ("20260414_⑥二、主要証拠抜粋（時系列）.pdf",
     "面談録音主要抜粋（4月14日学校面談）.pdf"),
    ("20260414_⑥二、主要証拠抜粋（時系列）_001.pdf",
     "面談録音主要抜粋（4月14日学校面談）_2.pdf"),
    ("20260428_CGAClaDetaiIsEnrolmentProcess.pdf",
     "CGA入学プロセス資料.pdf"),
    ("20260507.pdf",
     "証拠映像説明書面.pdf"),
    ("20260507_KISTKarenDonaldGodfrey.pdf",
     "学校メール_KarenDonaldGodfreyスチューデントケアコーディネーター.pdf"),
    ("20260507_KISTKarenDonaldGodfrey_001.pdf",
     "学校メール_KarenDonaldGodfreyスチューデントケアコーディネーター_2.pdf"),
    ("20260507_RECEIPT.pdf",
     "医療費領収書（AmericanClinicTokyo_4月16日精神科）.pdf"),
    ("20260507_Stairs.pdf",
     "事案発生現場の位置関係図.pdf"),
    ("20260507_●收件人@ParentsofG9AAngelinaFeiliMaK2AGeorgeManTzar.pdf",
     "学校メール_MarkCowe（4月13日事案当日）.pdf"),
    ("20260507_宗在鴫？単.pdf",
     "WeChat記録（証拠動画取得経緯）.pdf"),
    ("20260507_收件人OParentsofG9AAngelinaFeiliMaK2AGeorgeManTZar.pdf",
     "学校法務顧問メール_寺井隼人（4月20日）.pdf"),
    ("20260507_本件は、KInternationalSchoolTokyoにおいて発生した事案である。.pdf",
     "案件概要.pdf"),
    ("20260507_相談カード.pdf",
     "相談カード.pdf"),
]

SUZUKI_MOVE = ("20260507_1番1号.pdf", "1番1号書留通知（封筒）.pdf")


def safe_dest(dst_dir: Path, name: str) -> Path:
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


# 鈴木七海さんへ移動
SUZUKI.mkdir(parents=True, exist_ok=True)
src = MA_RES / SUZUKI_MOVE[0]
if src.exists():
    dst = safe_dest(SUZUKI, SUZUKI_MOVE[1])
    shutil.move(str(src), str(dst))
    print(f"[OK] {SUZUKI_MOVE[0]} -> {dst}")
else:
    print(f"[MISSING] {SUZUKI_MOVE[0]}")

# 馬さん 260507受領 サブフォルダへ
MA_RECEIVED.mkdir(parents=True, exist_ok=True)
for old, new in RENAMES:
    src = MA_RES / old
    if not src.exists():
        print(f"[MISSING] {old}")
        continue
    dst = safe_dest(MA_RECEIVED, new)
    shutil.move(str(src), str(dst))
    print(f"[OK] {old}\n     -> {dst.name}")
