"""田村フォルダの直近1週間ファイル8件を内容に応じて分類・リネーム。

8件目は事件番号令和6年(ワ)2386・原告毛崇熙の受領書FAXで、田村フォルダに誤入り。
毛さんVSヒメラ案件の04事務フォルダへ移す。
"""

from pathlib import Path
import shutil

BASE = Path(
    "/Users/kometaninaoki/Library/CloudStorage/"
    "GoogleDrive-n.kometani@re-startlaw.com/マイドライブ/共有用/01_事件記録"
)
TAMURA = BASE / "た_田村正宣"
MAO = BASE / "も_毛さんVSヒメラ合同会社"

MOVES = [
    (
        TAMURA / "260327_逮捕状.pdf",
        TAMURA / "02_身体拘束関係" / "260327_逮捕状（監禁・強盗致傷被疑事件）.pdf",
    ),
    (
        TAMURA / "260402_被疑者鈴木哲夫こと田村.pdf",
        TAMURA / "03_検察官提出書面" / "260410_検察官意見書（接見等禁止解除申立に対する意見・不相当）.pdf",
    ),
    (
        TAMURA / "260407_上申.pdf",
        TAMURA / "04_弁護人提出書面" / "260407_上申書（内縁の妻・大宮瑞希作成／接見等禁止解除申立疎明資料）.pdf",
    ),
    (
        TAMURA / "260501_起訴状・勾留質問調書・接見禁止決定.pdf",
        TAMURA / "02_身体拘束関係" / "260501_起訴状・勾留質問調書・接見等禁止決定（合本）.pdf",
    ),
    (
        TAMURA / "260508_鈴木哲夫こと田村正宣.pdf",
        TAMURA / "02_身体拘束関係" / "260501_勾留状（起訴後・被告人用・監禁被告事件）.pdf",
    ),
    (
        TAMURA / "691018_鈴木哲夫こと田村正宣.pdf",
        TAMURA / "02_身体拘束関係" / "260423_勾留状（再逮捕・被疑者用・監禁脅迫被疑事件）.pdf",
    ),
    (
        TAMURA / "260514_FAX_20260512_1778560600_780.pdf",
        TAMURA / "99_事務" / "260512_刑事事件記録閲覧謄写票ひな形（司法協会送付・FAX）.pdf",
    ),
    (
        TAMURA / "260518_FAX_20260518_1779068955_718.pdf",
        MAO / "訴訟" / "04事務" / "260518受領書（相手方・準備書面8）.pdf",
    ),
]


def main() -> None:
    for src, dst in MOVES:
        if not src.exists():
            print(f"[SKIP] 元ファイルなし: {src.name}")
            continue
        dst.parent.mkdir(parents=True, exist_ok=True)
        if dst.exists():
            # 衝突したら _2, _3 などで退避
            stem, suf = dst.stem, dst.suffix
            n = 2
            while True:
                candidate = dst.with_name(f"{stem}_{n}{suf}")
                if not candidate.exists():
                    dst = candidate
                    break
                n += 1
        shutil.move(str(src), str(dst))
        print(f"[OK] {src.name}\n  -> {dst.parent.name}/{dst.name}")


if __name__ == "__main__":
    main()
