"""D1（日本語全文検索）・D4（符号パース/自然順）の検証。"""

from __future__ import annotations

import sqlite3
from pathlib import Path

import evidence_index as ei


# --- D1: 検索3パターン ---------------------------------------------------

def _run(db: Path, query: str):
    conn = sqlite3.connect(db)
    try:
        return ei.search_pages(conn, query)
    finally:
        conn.close()


def test_search_three_char_fts(search_db: Path):
    """3文字語は trigram FTS でヒットする。"""
    hits = _run(search_db, "被害届")
    assert any(h["evidence_no"] == "甲1" for h in hits)


def test_search_two_char_like_fallback(search_db: Path):
    """2文字語は trigram に乗らないため正規化 LIKE フォールバックでヒットする。"""
    hits = _run(search_db, "供述")
    evs = {h["evidence_no"] for h in hits}
    assert "甲1" in evs and "乙3" in evs


def test_search_fullwidth_halfwidth(search_db: Path):
    """全角半角ゆれ: 半角クエリが本文中の全角数字にヒットする。"""
    hits = _run(search_db, "2026")
    assert any(h["evidence_no"] == "甲1" for h in hits)
    # 逆方向（全角クエリ→半角本文に相当）も正規化で吸収される
    hits2 = _run(search_db, "１５日")
    assert any(h["evidence_no"] == "甲1" for h in hits2)


def test_search_space_is_or(search_db: Path):
    """スペース区切りは OR 検索。"""
    hits = _run(search_db, "被害届 員面調書")
    evs = {h["evidence_no"] for h in hits}
    assert "甲1" in evs and "乙3" in evs


def test_normalize_idempotent():
    assert ei.normalize_for_search("ＡＢ１２") == "ab12"
    assert ei.normalize_for_search("被害") == "被害"


# --- D4: 符号パース ------------------------------------------------------

def test_evidence_regex_variants(tmp_path: Path):
    cases = {
        "甲1 被害届.pdf": "甲1",
        "乙3 員面調書.pdf": "乙3",
        "弁１０ 意見書.pdf": "弁１０",   # 全角数字・弁号証
        "甲2の1 実況見分.pdf": "甲2の1",  # 「の」枝番
        "甲7-2 写真.pdf": "甲7-2",       # ハイフン枝番
        "資料メモ.pdf": "",              # 符号無し
    }
    for name, expected in cases.items():
        p = tmp_path / name
        p.write_bytes(b"%PDF-1.4\n")
        meta = ei.parse_pdf_name(p, tmp_path)
        assert meta.evidence_no == expected, f"{name} -> {meta.evidence_no!r}"


def test_evidence_sort_key_natural_order():
    nums = ["甲10", "甲1", "甲2", "乙1", "弁1", "甲2の1", "甲2の2", ""]
    ordered = sorted(nums, key=ei.evidence_sort_key)
    assert ordered == ["甲1", "甲2", "甲2の1", "甲2の2", "甲10", "乙1", "弁1", ""]


def test_evidence_sort_fullwidth_equiv():
    # 全角・半角の番号が同じ順序キーに揃う
    assert ei.evidence_sort_key("弁１０")[:4] == ei.evidence_sort_key("弁10")[:4]


# --- 課題②: 弁護革命IDパーサ -------------------------------------------

REMOVE_CASES = [
    # 観測7例（除去すべき）
    ("甲1 被害届.d73ca7qjaj7b", "甲1", "被害届", "", ""),
    ("乙3 員面調書.6ggnjsv9fvd1", "乙3", "員面調書", "", ""),
    ("弁1 意見書.20pegmr6ib64", "弁1", "意見書", "", ""),
    # ID のみのケース（文書名 + 全角@作成者）
    ("甲5 供述調書＠田中太郎.abc123xyz789", "甲5", "供述調書", "", "田中太郎"),
    # 半角@
    ("乙2 証拠.def456uvw012@佐藤花子", "乙2", "証拠", "", "佐藤花子"),
]

KEEP_CASES = [
    # 除去してはいけない例
    ("報告書.最終版", "報告書.最終版"),        # ドット+日本語 → IDではない
    ("甲1 書面.20240101", "書面"),              # 純数字日付 → IDではない（数字のみ）
    ("甲2 図面", "図面"),                      # IDなし
    ("資料", "資料"),                          # IDなし・符号なし
]


def test_bengokakumei_id_remove():
    """弁護革命IDが正しく除去されること。"""
    for stem, ev, title, date, author in REMOVE_CASES:
        e, t, d, a = ei.parse_name_only(stem)
        assert e == ev, f"stem={stem!r}: evidence_no {e!r} != {ev!r}"
        assert t == title, f"stem={stem!r}: title {t!r} != {title!r}"
        if author:
            assert a == author, f"stem={stem!r}: author {a!r} != {author!r}"
        # ID パターン（英数字10-16桁）がタイトルに残っていないこと
        assert not ei.BENGOKAKUMEI_ID_RE.search(t), f"ID残存 stem={stem!r} -> title={t!r}"


def test_bengokakumei_id_keep():
    """除去すべきでないファイル名は改変されないこと。"""
    for stem, expected_title in KEEP_CASES:
        _, t, _, _ = ei.parse_name_only(stem)
        assert t == expected_title, f"stem={stem!r}: title {t!r} != {expected_title!r}"


def test_bengokakumei_id_not_empty_after_remove():
    """ID除去後に空文字になる場合は元のstemを使うこと。"""
    # IDだけのstem（実用上存在しないがエッジケース）
    stem_only_id = "a1b2c3d4e5f6"  # 12桁英数混在 → IDパターンにマッチするが先頭ドットがないので除去対象外
    _, t, _, _ = ei.parse_name_only(stem_only_id)
    assert t  # タイトルが空にならないこと


def test_at_sign_half_and_full():
    """全角＠・半角@ の両方で作成者分離が働くこと（最後の出現で分割）。"""
    _, _, _, a_full = ei.parse_name_only("被害届＠山田一郎")
    assert a_full == "山田一郎"

    _, _, _, a_half = ei.parse_name_only("被害届@山田一郎")
    assert a_half == "山田一郎"

    # 複数 @ → 最後で分割
    _, t_multi, _, a_multi = ei.parse_name_only("甲1 A＠B@作成者名")
    assert a_multi == "作成者名"
    assert "A＠B" in t_multi


def test_id_removal_prevents_date_contamination():
    """ID除去後に日付抽出が行われるため、IDに含まれる数字列が日付誤抽出されないこと。"""
    # "20pegmr6ib64" の "20" 部分が日付にならない
    _, _, doc_date, _ = ei.parse_name_only("弁1 意見書.20pegmr6ib64")
    assert doc_date == "", f"IDの数字が日付誤抽出された: {doc_date!r}"

    # "6ggnjsv9fvd1" も同様
    _, _, doc_date2, _ = ei.parse_name_only("乙3 員面調書.6ggnjsv9fvd1")
    assert doc_date2 == "", f"IDの数字が日付誤抽出された: {doc_date2!r}"
