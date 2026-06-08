"""甲1.pdfの個人情報を黒塗りした版を作成する(v3: 部分矩形マッチング)。

word内のキーワード文字位置を計算し、その部分だけを矩形として塗る。
OCRが長い文字列を1つのwordとして認識した場合の過剰塗りを防ぐ。
"""
from pathlib import Path
import re
import unicodedata
import fitz


def nfkc_each(text):
    """各文字を独立にNFKC正規化(同じ長さの文字列を返す)。"""
    out = []
    for c in text:
        n = unicodedata.normalize("NFKC", c)
        if len(n) != 1:
            out.append(c)
        else:
            out.append(n)
    return "".join(out)

PATH_FILE = Path(".cursor/shell_path_utf8.txt")
src = Path(PATH_FILE.read_text().strip().splitlines()[0])
dst = src.with_name("甲１_黒塗り.pdf")
print(f"src: {src.name}")
print(f"dst: {dst.name}")

KEYWORDS = [
    # 個人氏名
    "玉本真也", "玉本", "真也",
    "太田義弘", "太田", "義弘",
    "樫尾恒卓", "樫尾", "恒卓",
    "古市哲也", "古市", "哲也",
    "イケダシオン", "イガハルナ", "ホリウチキョウヘイ", "モリリョウスケ",
    # 会社・サイト
    "株式会社七福堂", "七福堂",
    "株式会社ネクストステージ", "ネクストステージ",
    "株式会社OJKアネックス", "ＯＪＫアネックス", "OJKアネックス",
    "OJKアネ", "ＯＪＫ", "OJK", "アネックス",
    "ギフトチェンジ",
    "ビッグローブ", "Biglobe", "BIGLOBE",
    "アルテリア・ネットワークス", "ARTERIA",
    "ファミリーネット・ジャパン", "ファミリーネット",
    "admin.mobachen.net", "mobachen.net", "mobachen",
    "https://admin",
    "Tradeuserdetails/kesaibun", "Tradeuserdetails", "kesaibun", "kBsaibun",
    "/users/login", "/Tradeuserdetails", ".net", "net/",
    "mesh.ad.jp", "ucom.ne.jp", "cyberhome.jp",
    "biglobe.ne.jp", "arteria-net.com", "fnj.co.jp",
    # 住所
    "東京都豊島区東池袋", "豊島区東池袋", "東池袋1丁目", "東池袋１丁目",
    "1丁目36番7号", "１丁目３６番７号", "36番7号", "３６番７号",
    "アルテール池袋", "アルテール", "203号室", "２０３号室", "204号室", "２０４号室",
    "千葉県", "長野県", "東京都",
    "Chiba", "Tokyo", "Nagano", "Ichibachō", "Kashiwa", "Noda", "Osu",
    "Higashinakazawa", "Abiko", "Ikejiri",
    # 郵便番号
    "260-0851", "277-0005", "278-0001", "272-0032",
    "273-0036", "270-1151", "154-0001", "380-0811",
    # IPアドレス(明示パターン)
    "160.237.152.151", "160 237 152 151", "160.237.152", "160 237 152",
    # ID
    "fhfugeyfugjibnfih", "jgijgiuirhigiahif", "Pgij6iuirhiXiahif",
    "gmhEmAw7iFa2", "mcniirkfkmxjvjirkd", "GiftAdmin87695264",
    "ft201304",
    # タイムスタンプ
    "15/Nov/2023:08:00:51",
    # 座標
    "35.5973", "140.1423", "35.8675", "139.9832",
    "35.9718", "139.8935", "35.7178", "139.9054",
    "35.7199", "139.9477", "35.8699", "140.0093",
    "35.6453", "139.6813", "36.6497", "138.2013",
]
KEYWORDS = sorted(set(KEYWORDS), key=lambda s: -len(s))


def widx_set_for_range(all_words, char_start, char_end):
    """joined_all内の文字範囲 → 該当word indexのset。
    joined_allの構築方式: wordテキスト連結 + 各wordの後ろに" "。"""
    indices = set()
    pos = 0
    for idx, w in enumerate(all_words):
        wlen = len(w[4])
        w_start = pos
        w_end = pos + wlen
        if w_end > char_start and w_start < char_end:
            indices.add(idx)
        pos = w_end + 1  # +1 for separator " "
    return indices


def widx_at_pos(all_words, char_pos):
    pos = 0
    for idx, w in enumerate(all_words):
        wlen = len(w[4])
        if pos <= char_pos < pos + wlen:
            return idx
        pos += wlen + 1
    return None


def add_partial_rect(page, word, char_start, char_end):
    """word内の文字範囲[char_start, char_end)に対応する部分矩形を追加。"""
    x0, y0, x1, y1, text = word[0], word[1], word[2], word[3], word[4]
    n = len(text)
    if n == 0:
        return
    cw = (x1 - x0) / n
    rx0 = x0 + cw * char_start - 1
    rx1 = x0 + cw * char_end + 1
    r = fitz.Rect(rx0, y0 - 1, rx1, y1 + 1)
    page.add_redact_annot(r, fill=(0, 0, 0))


def redact_page(page, keywords):
    marks = 0
    words = page.get_text("words")
    if not words:
        return 0

    # 行単位でグループ化
    rows = {}
    for w in words:
        key = (w[5], w[6])
        rows.setdefault(key, []).append(w)

    for key, row in rows.items():
        row.sort(key=lambda w: w[0])
        # joined text と char→word マップ(word間にスペース挿入)
        joined = ""
        char_word_idx = []  # joined各文字 → row内のidx (スペースは -1)
        char_in_word = []   # joined各文字 → word内の文字idx (スペースは -1)
        for ridx, w in enumerate(row):
            if joined:
                joined += " "
                char_word_idx.append(-1)
                char_in_word.append(-1)
            for ci in range(len(w[4])):
                char_word_idx.append(ridx)
                char_in_word.append(ci)
            joined += w[4]
        if not joined:
            continue

        # NFKC正規化(全角→半角)した文字列で検索する。長さは保持。
        norm = nfkc_each(joined)

        ranges = []  # (joined_start, joined_end)
        for kw in keywords:
            kw_norm = nfkc_each(kw)
            start = 0
            while True:
                pos = norm.find(kw_norm, start)
                if pos == -1:
                    break
                ranges.append((pos, pos + len(kw_norm)))
                start = pos + 1

        # IPアドレス(行内連結後、ピリオド/スペース/カンマのみ許可、3桁要素2つ以上)
        ip_pat = re.compile(r"(\d{1,3})[.,\s]+(\d{1,3})[.,\s]+(\d{1,3})[.,\s]+(\d{1,3})")
        for m in ip_pat.finditer(norm):
            parts = [int(m.group(i)) for i in range(1, 5)]
            if not all(0 <= p <= 255 for p in parts):
                continue
            if sum(1 for p in parts if p >= 100) < 2:
                continue
            ranges.append((m.start(), m.end()))

        # ID(英数字10文字以上、特定の許可語を除く)
        # スペース除去版で検索し、元位置に戻す
        norm_ns_chars = []
        ns_to_norm = []
        for i, c in enumerate(norm):
            if not c.isspace():
                norm_ns_chars.append(c)
                ns_to_norm.append(i)
        norm_ns = "".join(norm_ns_chars)

        id_pat = re.compile(r"[A-Za-z0-9]{10,}")
        ALLOWED_IDS = {"hostgator", "applegift", "amazongift"}
        for m in id_pat.finditer(norm_ns):
            if m.group().lower() in ALLOWED_IDS:
                continue
            if m.start() >= len(ns_to_norm) or m.end() - 1 >= len(ns_to_norm):
                continue
            s = ns_to_norm[m.start()]
            e = ns_to_norm[m.end() - 1] + 1
            ranges.append((s, e))

        # 範囲を部分矩形に変換
        for js, je in ranges:
            cur_idx = None
            cur_s = None
            cur_e = None
            for c in range(js, min(je, len(char_word_idx))):
                widx = char_word_idx[c]
                cidx = char_in_word[c]
                if widx < 0:
                    if cur_idx is not None:
                        add_partial_rect(page, row[cur_idx], cur_s, cur_e + 1)
                        marks += 1
                        cur_idx = None
                    continue
                if widx != cur_idx:
                    if cur_idx is not None:
                        add_partial_rect(page, row[cur_idx], cur_s, cur_e + 1)
                        marks += 1
                    cur_idx = widx
                    cur_s = cidx
                    cur_e = cidx
                else:
                    cur_e = cidx
            if cur_idx is not None:
                add_partial_rect(page, row[cur_idx], cur_s, cur_e + 1)
                marks += 1
    # --- 追加: ページ全体のwords連結でIP/タイムスタンプを検出 ---
    # OCRが数字をblock/lineで分散させる場合に対応
    all_words = sorted(words, key=lambda w: (w[5], w[6], w[0]))
    joined_all = ""
    pos_to_widx = []
    for idx, w in enumerate(all_words):
        for _ in w[4]:
            pos_to_widx.append(idx)
            joined_all += "X"  # ダミー(実際の文字は別で扱う)
        # 区切り(別wordとして識別するためのスペース)
        pos_to_widx.append(-1)
        joined_all += " "
    # 正しい文字列を作り直す
    joined_all = ""
    for idx, w in enumerate(all_words):
        joined_all += w[4]
        joined_all += " "
    norm_all = nfkc_each(joined_all)

    ip_global = re.compile(r"\b(\d{1,3})[.,\s\-]+(\d{1,3})[.,\s\-]+(\d{1,3})[.,\s\-]+(\d{1,3})\b")
    ts_global = re.compile(
        r"\d{1,2}\s*/\s*(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)"
        r"\s*/\s*\d{4}",
        re.IGNORECASE,
    )

    for m in ip_global.finditer(norm_all):
        parts = [int(m.group(i)) for i in range(1, 5)]
        if not all(0 <= p <= 255 for p in parts):
            continue
        if sum(1 for p in parts if p >= 100) < 2:
            continue
        # IP範囲妥当 → 該当文字位置のwordsを塗る
        for idx in widx_set_for_range(all_words, m.start(), m.end()):
            w = all_words[idx]
            page.add_redact_annot(
                fitz.Rect(w[0] - 1, w[1] - 1, w[2] + 1, w[3] + 1), fill=(0, 0, 0)
            )
            marks += 1

    for m in ts_global.finditer(norm_all):
        # タイムスタンプ前後を含めて広めに塗る(-0600などの近隣も)
        s_pos = max(0, m.start() - 20)
        e_pos = min(len(norm_all), m.end() + 30)
        for idx in widx_set_for_range(all_words, s_pos, e_pos):
            w = all_words[idx]
            # 同じy座標±30に絞る(他ページのテキスト混入を避ける)
            tm = (m.start() + m.end()) // 2
            target_w = widx_at_pos(all_words, tm)
            if target_w is None:
                continue
            ty = (all_words[target_w][1] + all_words[target_w][3]) / 2
            wy = (w[1] + w[3]) / 2
            if abs(wy - ty) > 30:
                continue
            page.add_redact_annot(
                fitz.Rect(w[0] - 1, w[1] - 1, w[2] + 1, w[3] + 1), fill=(0, 0, 0)
            )
            marks += 1

    # --- 追加: y座標ベースで「IPアドレス/タイムスタンプ単独行」を検出して塗る ---
    # OCRがwordをline/block分断したケースをカバー
    y_rows = {}
    for w in words:
        y_center = (w[1] + w[3]) / 2
        y_key = round(y_center)
        y_rows.setdefault(y_key, []).append(w)

    # 隣接y座標(±3)をマージ
    y_keys = sorted(y_rows.keys())
    merged_rows = []
    cur_key = None
    cur_row = []
    for yk in y_keys:
        if cur_key is None or yk - cur_key <= 3:
            cur_row.extend(y_rows[yk])
            cur_key = yk
        else:
            merged_rows.append(cur_row)
            cur_row = list(y_rows[yk])
            cur_key = yk
    if cur_row:
        merged_rows.append(cur_row)

    ip_re = re.compile(r"(\d{1,3})[.,\s\-]+(\d{1,3})[.,\s\-]+(\d{1,3})[.,\s\-]+(\d{1,3})")
    ts_re = re.compile(
        r"\d{1,2}/(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)/\d{4}"
        r"[:\s]?\d{0,2}[:\s]?\d{0,2}[:\s]?\d{0,2}"
    )
    for prow in merged_rows:
        prow.sort(key=lambda w: w[0])
        text = " ".join(w[4] for w in prow)
        norm = nfkc_each(text)
        hit_ip = ip_re.search(norm)
        if hit_ip:
            parts = [int(hit_ip.group(i)) for i in range(1, 5)]
            if not (all(0 <= p <= 255 for p in parts) and sum(1 for p in parts if p >= 100) >= 2):
                hit_ip = None
        hit_ts = ts_re.search(norm)
        if hit_ip or hit_ts:
            # 行全体を塗る(IP/タイムスタンプ単独行を想定)
            xs0 = [w[0] for w in prow]
            xs1 = [w[2] for w in prow]
            ys0 = [w[1] for w in prow]
            ys1 = [w[3] for w in prow]
            r = fitz.Rect(min(xs0) - 1, min(ys0) - 1, max(xs1) + 1, max(ys1) + 1)
            page.add_redact_annot(r, fill=(0, 0, 0))
            marks += 1

    return marks


doc = fitz.open(str(src))
print(f"pages: {len(doc)}")

total = 0
for i, page in enumerate(doc):
    n = redact_page(page, KEYWORDS)
    total += n
    print(f"page {i+1}: marks={n}")
print(f"total: {total}")

# 手書き署名・指印エリア(相対座標)
HAND_REGIONS = {
    1: [(0.25, 0.18, 0.95, 0.32)],
    # ページ7のIP表示行(2箇所)
    7: [(0.27, 0.14, 0.70, 0.185), (0.27, 0.71, 0.70, 0.76)],
    # ページ10のIP表示行
    10: [(0.27, 0.31, 0.60, 0.36)],
    # ページ11のタイムスタンプ行(上端)
    11: [(0.25, 0.10, 0.85, 0.15)],
    13: [(0.45, 0.78, 0.95, 0.95)],
    14: [(0.40, 0.05, 0.95, 0.30)],
    15: [(0.55, 0.75, 0.95, 0.95)],
    17: [(0.55, 0.75, 0.95, 0.95)],
    18: [(0.65, 0.85, 0.95, 1.0)],
    20: [(0.65, 0.85, 0.95, 1.0)],
    21: [(0.65, 0.85, 0.95, 1.0)],
    22: [(0.65, 0.85, 0.95, 1.0)],
}
for pno, regions in HAND_REGIONS.items():
    if pno - 1 >= len(doc):
        continue
    page = doc[pno - 1]
    w, h = page.rect.width, page.rect.height
    for (rx0, ry0, rx1, ry1) in regions:
        r = fitz.Rect(w * rx0, h * ry0, w * rx1, h * ry1)
        page.add_redact_annot(r, fill=(0, 0, 0))

for page in doc:
    page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_PIXELS)

doc.save(str(dst), garbage=4, deflate=True)
doc.close()
print(f"saved: {dst}")
