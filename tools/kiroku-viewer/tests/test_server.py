"""サーバAPIの検証（documents 応答・Range=206 / D12・D11 セキュリティ）。"""

from __future__ import annotations

from pathlib import Path

import pytest
from fastapi.testclient import TestClient

import server
from conftest import build_db


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


# --- 1-A: search / text / open-file / reindex ---------------------------

@pytest.fixture
def indexed_client(tmp_path: Path):
    case_dir = tmp_path / "icase"
    (case_dir / "_index" / "export").mkdir(parents=True)
    sha1, sha2 = "1" * 64, "2" * 64
    build_db(
        case_dir / "_index" / "export" / "evidence_index.sqlite",
        [
            {
                "evidence_no": "甲1", "title": "被害届",
                "rel_path": "甲1 被害届.pdf", "sha256": sha1,
                "pages": ["被害届。供述調書のとおり申告した。", "2ページ目の本文。"],
            },
            {
                "evidence_no": "乙3", "title": "員面調書",
                "rel_path": "乙3 員面調書.pdf", "sha256": sha2,
                "pages": ["員面調書。田中太郎の供述。"],
            },
        ],
    )
    server.CONFIG = {
        "cases": {
            "ic": {"id": "ic", "name": "索引済", "path": case_dir, "reindex": False}
        },
        "display_names": {},
    }
    server._listing_cache.clear()
    return TestClient(server.app, base_url="http://127.0.0.1:8788"), sha1, sha2


def test_search_endpoint(indexed_client):
    client, sha1, _ = indexed_client
    r = client.get("/api/cases/ic/search", params={"q": "被害届"})
    assert r.status_code == 200
    hits = r.json()["hits"]
    assert any(h["sha256"] == sha1 and h["page_no"] == 1 for h in hits)
    assert all({"sha256", "page_no", "snippet", "evidence_no", "title"} <= h.keys() for h in hits)


def test_search_filter(indexed_client):
    client, sha1, sha2 = indexed_client
    # 「供述」は両文書にヒットするが、符号フィルタ「乙」で乙3のみに絞る
    r = client.get("/api/cases/ic/search", params={"q": "供述", "filter": "乙"})
    shas = {h["sha256"] for h in r.json()["hits"]}
    assert shas == {sha2}


def test_text_endpoint_page(indexed_client):
    client, sha1, _ = indexed_client
    r = client.get(f"/api/cases/ic/text/{sha1}", params={"page": 2})
    data = r.json()
    assert data["page_count"] == 2
    assert data["page_no"] == 2
    assert "2ページ目" in data["text"]


def test_text_endpoint_all_pages(indexed_client):
    client, sha1, _ = indexed_client
    r = client.get(f"/api/cases/ic/text/{sha1}")
    data = r.json()
    assert len(data["pages"]) == 2


def test_open_file_validates_and_calls(client, monkeypatch):
    calls = {}

    def fake_run(cmd, check):
        calls["cmd"] = cmd
        class R:
            returncode = 0
        return R()

    import subprocess
    monkeypatch.setattr(subprocess, "run", fake_run)
    docs = client.get("/api/cases/test/documents").json()["documents"]
    sha = next(d["sha256"] for d in docs if d["evidence_no"] == "弁2")
    r = client.post(
        f"/api/cases/test/open-file/{sha}", headers={"X-Kiroku-Viewer": "1"}
    )
    assert r.status_code == 200
    assert calls["cmd"][0] == "open"


def test_open_file_requires_csrf(client):
    docs = client.get("/api/cases/test/documents").json()["documents"]
    sha = next(d["sha256"] for d in docs if d["evidence_no"] == "弁2")
    r = client.post(f"/api/cases/test/open-file/{sha}")  # CSRFヘッダ無し
    assert r.status_code == 403


def test_reindex_disabled_returns_403(indexed_client):
    client, _, _ = indexed_client
    r = client.post("/api/cases/ic/reindex", headers={"X-Kiroku-Viewer": "1"})
    assert r.status_code == 403
