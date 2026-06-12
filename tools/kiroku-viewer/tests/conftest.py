"""pytest 共通設定・フィクスチャ。"""

from __future__ import annotations

import sqlite3
import sys
from pathlib import Path

import pytest

HERE = Path(__file__).resolve().parent
TOOL_DIR = HERE.parent
REPO_ROOT = TOOL_DIR.parents[1]
sys.path.insert(0, str(REPO_ROOT / "scripts"))
sys.path.insert(0, str(TOOL_DIR))

import evidence_index as ei  # noqa: E402


def build_db(db_path: Path, docs: list[dict]) -> None:
    """テスト用に索引DBを直接構築する。

    docs: [{"evidence_no", "title", "rel_path", "sha256", "pages": [text, ...]}]
    本番の build と同じく pages.text_norm と pages_fts には正規化テキストを入れる。
    """
    conn = sqlite3.connect(db_path)
    try:
        ei.create_schema(conn)
        created = ei.now_iso()
        for d in docs:
            cur = conn.execute(
                """
                INSERT INTO documents (
                    rel_path, abs_path, file_name, evidence_no, title,
                    document_date, person_or_source, sha256, file_size,
                    modified_at, acquired_at, page_count, extract_method,
                    extract_status, created_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    d["rel_path"], d.get("abs_path", d["rel_path"]),
                    d.get("file_name", d["rel_path"]), d["evidence_no"], d["title"],
                    d.get("document_date", ""), d.get("author", ""), d["sha256"],
                    d.get("file_size", 0), "", "", len(d["pages"]),
                    "fitz", "ok", created,
                ),
            )
            doc_id = int(cur.lastrowid)
            for i, text in enumerate(d["pages"], start=1):
                norm = ei.normalize_for_search(text)
                cur = conn.execute(
                    """
                    INSERT INTO pages (
                        document_id, page_no, text, text_norm, text_hash,
                        extract_method, ocr_confidence, review_status, error
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (doc_id, i, text, norm, ei.text_hash(text), "fitz", None, "ok", ""),
                )
                page_id = int(cur.lastrowid)
                conn.execute(
                    "INSERT INTO pages_fts (text, evidence_no, title, rel_path, page_id)"
                    " VALUES (?, ?, ?, ?, ?)",
                    (norm, d["evidence_no"], d["title"], d["rel_path"], page_id),
                )
        conn.commit()
    finally:
        conn.close()


@pytest.fixture
def search_db(tmp_path: Path) -> Path:
    db = tmp_path / "evidence_index.sqlite"
    build_db(
        db,
        [
            {
                "evidence_no": "甲1",
                "title": "被害届",
                "rel_path": "甲1 被害届.pdf",
                "sha256": "a" * 64,
                "pages": [
                    "被害届。令和６年１月１５日、被害者は供述調書のとおり申告した。"
                    "実況見分調書の作成日は２０２６年１月１６日である。",
                ],
            },
            {
                "evidence_no": "乙3",
                "title": "員面調書",
                "rel_path": "乙3 員面調書.pdf",
                "sha256": "b" * 64,
                "pages": ["員面調書。田中太郎の供述を録取した。"],
            },
        ],
    )
    return db
