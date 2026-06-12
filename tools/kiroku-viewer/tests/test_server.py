"""サーバAPIの検証（documents 応答・Range=206 / D12・D11 セキュリティ）。"""

from __future__ import annotations

from pathlib import Path

import pytest
from fastapi.testclient import TestClient

import server


@pytest.fixture
def client(tmp_path: Path):
    case_dir = tmp_path / "case"
    case_dir.mkdir()
    # Range 検証用に十分なサイズのダミーPDF（バイト列でよい）。
    pdf = case_dir / "甲1 テスト書面.pdf"
    pdf.write_bytes(b"%PDF-1.4\n" + b"x" * 4096 + b"\n%%EOF")
    # メディア（非PDF）も1件。
    (case_dir / "弁2 録画.mp4").write_bytes(b"\x00\x00\x00\x18ftypmp42" + b"\x00" * 64)

    server.CONFIG = {
        "cases": {
            "test": {"id": "test", "name": "テスト事件", "path": case_dir, "reindex": False}
        },
        "display_names": {},
    }
    server._listing_cache.clear()
    # Host ヘッダが許可値になるよう base_url を設定。
    return TestClient(server.app, base_url="http://127.0.0.1:8788")


def test_documents_lists_unindexed(client):
    r = client.get("/api/cases/test/documents")
    assert r.status_code == 200
    data = r.json()
    docs = {d["evidence_no"]: d for d in data["documents"]}
    assert "甲1" in docs and docs["甲1"]["kind"] == "pdf"
    assert docs["甲1"]["indexed"] is False  # 未索引バッジ
    assert "弁2" in docs and docs["弁2"]["kind"] == "media"
    # 自然順（甲 < 弁）
    order = [d["evidence_no"] for d in data["documents"]]
    assert order.index("甲1") < order.index("弁2")


def test_pdf_range_returns_206(client):
    docs = client.get("/api/cases/test/documents").json()["documents"]
    sha = next(d["sha256"] for d in docs if d["evidence_no"] == "甲1")
    r = client.get(
        f"/api/cases/test/pdf/{sha}", headers={"Range": "bytes=0-1023"}
    )
    assert r.status_code == 206
    assert r.headers["accept-ranges"] == "bytes"
    assert r.headers["content-range"].startswith("bytes 0-1023/")
    assert len(r.content) == 1024


def test_pdf_full_request_has_accept_ranges(client):
    docs = client.get("/api/cases/test/documents").json()["documents"]
    sha = next(d["sha256"] for d in docs if d["evidence_no"] == "甲1")
    r = client.get(f"/api/cases/test/pdf/{sha}")
    assert r.status_code == 200
    assert r.headers.get("accept-ranges") == "bytes"


def test_host_header_rejected():
    client = TestClient(server.app, base_url="http://evil.example.com")
    r = client.get("/api/cases")
    assert r.status_code == 403


def test_unknown_sha_404(client):
    r = client.get("/api/cases/test/pdf/" + "0" * 64)
    assert r.status_code == 404
