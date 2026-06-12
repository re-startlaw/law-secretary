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
