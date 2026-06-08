# -*- coding: utf-8 -*-
"""
大中忠生事件 証拠PDF 資料番号・頁番号スタンプ付与スクリプト。

- 横長ページ: set_rotation(270) で見本同様に回転
- 1ページ目: 表示ビュー右上に「資料番号」朱色16pt（ヒラギノ角ゴW3）
- 2頁以上: 全ページ表示ビュー下端中央に頁番号（黒10.5pt）
- 元ファイルは温存し、提出正式版_資料番号入り/ へ出力

日本語パスはこのファイル内の文字列リテラルに閉じ込める（CLAUDE.md準拠）。
"""
import os
import sys
import shutil
import fitz  # PyMuPDF

FONT_PATH = "/System/Library/Fonts/ヒラギノ角ゴシック W3.ttc"
RED = (254 / 255, 0, 0)
BLACK = (0, 0, 0)
STAMP_SIZE = 16
PAGENO_SIZE = 10.5

BASE = ("/Users/kometaninaoki/Library/CloudStorage/"
        "GoogleDrive-n.kometani@re-startlaw.com/.shortcut-targets-by-id/"
        "1pEJgbHuaj4L4ovFEISOuq-0LeHGInxYl/大中 忠生 : 詐欺未遂/資料/"
        "★大中さんと共有/証拠アップロード用フォルダ")

FOLDERS = [
    os.path.join(BASE, "京阪事務機器", "提出正式版"),
    os.path.join(BASE, "レント", "提出正式版"),
]

# 加工不要（見本）: このファイル名は除外し、出力フォルダへそのままコピー
SAMPLE_SKIP = "資料１の１_京阪事務機器_備品管理台帳_ラベル表示（サンプル）.pdf"
OUT_DIRNAME = "提出正式版_資料番号入り"


def shiryo_label(filename):
    """ファイル名先頭の資料番号（最初の '_' まで）。'_' が無ければ拡張子を除いた全体。"""
    stem = os.path.splitext(filename)[0]
    return stem.split("_")[0]


def is_landscape(page):
    w, h = page.rect.width, page.rect.height
    return w > h


def _word_rects(page):
    """既存テキストの矩形（未回転座標系）一覧。"""
    return [fitz.Rect(w[:4]) for w in page.get_text("words")]


def _find_clear_topright(page, box_w, box_h, margin=24, gap=6):
    """表示ビュー右上から下方向に走査し、既存テキストと重ならない最初の矩形
    （表示座標系）を返す。見つからなければ最上段を返す。"""
    disp = page.rect
    words = _word_rects(page)
    x1 = disp.width - margin
    x0 = x1 - box_w
    top = _disp_rect = None
    y0 = margin
    limit = disp.height * 0.5  # 上半分まで探索
    while y0 + box_h <= limit:
        disp_rect = fitz.Rect(x0, y0, x1, y0 + box_h)
        if top is None:
            top = disp_rect
        unrot = (disp_rect * page.derotation_matrix) + (-3, -3, 3, 3)  # 余白付き
        if not any(unrot.intersects(wr) for wr in words):
            return disp_rect
        y0 += box_h + gap
    return top


def stamp_shiryo_number(page, text):
    """1ページ目: 表示ビュー右上に資料番号を水平・上向きで配置（既存文字と非重複）。
    insert_textbox の rect は未回転座標系なので derotation_matrix で変換し、
    rotate=回転角 を与えて表示上で上向きにする。"""
    box_w = STAMP_SIZE * len(text) + 10
    box_h = STAMP_SIZE + 8
    disp_rect = _find_clear_topright(page, box_w, box_h)
    rot = page.rotation
    unrot_rect = disp_rect * page.derotation_matrix
    page.insert_textbox(
        unrot_rect, text,
        fontname="hkg", fontfile=FONT_PATH,
        fontsize=STAMP_SIZE, color=RED,
        align=fitz.TEXT_ALIGN_RIGHT, rotate=rot,
    )


def stamp_page_number(page, n):
    """全ページ: 表示ビュー下端中央に頁番号。"""
    disp = page.rect
    box_w = 80
    box_h = PAGENO_SIZE + 8
    margin_bottom = 18
    x0 = disp.width / 2 - box_w / 2
    x1 = disp.width / 2 + box_w / 2
    y1 = disp.height - margin_bottom
    y0 = y1 - box_h
    disp_rect = fitz.Rect(x0, y0, x1, y1)
    rot = page.rotation
    unrot_rect = disp_rect * page.derotation_matrix
    page.insert_textbox(
        unrot_rect, str(n),
        fontname="hkg", fontfile=FONT_PATH,
        fontsize=PAGENO_SIZE, color=BLACK,
        align=fitz.TEXT_ALIGN_CENTER, rotate=rot,
    )


def process_pdf(src_path, out_path, label):
    doc = fitz.open(src_path)
    npages = doc.page_count
    for i, page in enumerate(doc):
        if is_landscape(page):
            page.set_rotation(270)
        if i == 0:
            stamp_shiryo_number(page, label)
        if npages >= 2:
            stamp_page_number(page, i + 1)
    doc.save(out_path, garbage=4, deflate=True)
    doc.close()
    return npages


def main(test_only=False):
    targets = []
    for folder in FOLDERS:
        out_dir = os.path.join(folder, OUT_DIRNAME)
        for f in sorted(os.listdir(folder)):
            if not f.lower().endswith(".pdf"):
                continue
            targets.append((folder, out_dir, f))

    if test_only:
        # 横長1件＋縦長1件だけ加工してPNG出力
        test_files = {
            "資料１の２_京阪事務機器_備品管理台帳_写真表示（サンプル）.pdf",
            "資料３の１_見積書及び請求書（サンプル）.pdf",
        }
        for folder, out_dir, f in targets:
            if f not in test_files:
                continue
            src = os.path.join(folder, f)
            tmp_out = os.path.join("/tmp", "stamp_test_" + f)
            label = shiryo_label(f)
            process_pdf(src, tmp_out, label)
            d = fitz.open(tmp_out)
            png = os.path.join("/tmp", "stamp_test_" + os.path.splitext(f)[0] + "_p1.png")
            d[0].get_pixmap(dpi=110).save(png)
            print("TEST", f, "label=", label, "->", png)
            d.close()
        return

    for folder, out_dir, f in targets:
        os.makedirs(out_dir, exist_ok=True)
        src = os.path.join(folder, f)
        out = os.path.join(out_dir, f)
        if f == SAMPLE_SKIP:
            shutil.copy2(src, out)
            print("COPY(sample)", f)
            continue
        label = shiryo_label(f)
        npages = process_pdf(src, out, label)
        print("STAMP", f, "label=", label, "pages=", npages)


if __name__ == "__main__":
    main(test_only="--test" in sys.argv)
