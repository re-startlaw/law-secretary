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


# --- フェーズ2: 注釈 / メタ / 書き出し / 孤児注釈 -----------------------

CSRF = {"X-Kiroku-Viewer": "1"}


def _real_pdf(path: Path, n_pages: int = 1):
    from pypdf import PdfWriter
    w = PdfWriter()
    for _ in range(n_pages):
        w.add_blank_page(width=595, height=842)
    with path.open("wb") as f:
        w.write(f)


@pytest.fixture
def anno_client(tmp_path: Path):
    case_dir = tmp_path / "acase"
    case_dir.mkdir()
    _real_pdf(case_dir / "甲1 書面.pdf", 2)
    server.CONFIG = {
        "cases": {"a": {"id": "a", "name": "注釈事件", "path": case_dir, "reindex": False}},
        "display_names": {},
    }
    server._listing_cache.clear()
    client = TestClient(server.app, base_url="http://127.0.0.1:8788")
    sha = next(d["sha256"] for d in client.get("/api/cases/a/documents").json()["documents"])
    return client, sha, case_dir


def test_annotation_put_get_roundtrip(anno_client):
    client, sha, _ = anno_client
    ann = {"id": "x1", "type": "rect", "page": 1, "rect": [10, 20, 100, 80], "color": "#ff0000"}
    r = client.put(f"/api/cases/a/annotations/{sha}", headers=CSRF,
                   json={"base_updated_at": "", "annotations": [ann]})
    assert r.status_code == 200
    got = client.get(f"/api/cases/a/annotations/{sha}").json()
    assert len(got["annotations"]) == 1
    a = got["annotations"][0]
    assert a["type"] == "rect" and a["_editable"] is True


def test_annotation_optimistic_lock_409(anno_client):
    client, sha, _ = anno_client
    r1 = client.put(f"/api/cases/a/annotations/{sha}", headers=CSRF,
                    json={"base_updated_at": "", "annotations": []})
    stored = r1.json()["updated_at"]
    # 古い base_updated_at で再PUT → 409
    r2 = client.put(f"/api/cases/a/annotations/{sha}", headers=CSRF,
                    json={"base_updated_at": "2000-01-01T00:00:00+00:00", "annotations": []})
    assert r2.status_code == 409
    # 正しい base なら成功
    r3 = client.put(f"/api/cases/a/annotations/{sha}", headers=CSRF,
                    json={"base_updated_at": stored, "annotations": []})
    assert r3.status_code == 200


def test_annotation_persists_across_reindex(anno_client):
    """sha256キーなので索引再生成相当でも注釈は同一文書に残る（D2）。"""
    client, sha, _ = anno_client
    ann = {"id": "p1", "type": "note", "page": 1, "point": [50, 700], "text": "重要"}
    client.put(f"/api/cases/a/annotations/{sha}", headers=CSRF,
               json={"base_updated_at": "", "annotations": [ann]})
    server._listing_cache.clear()  # 再索引相当でキャッシュ破棄
    got = client.get(f"/api/cases/a/annotations/{sha}").json()
    assert got["annotations"][0]["text"] == "重要"


def test_meta_memo_and_category(anno_client):
    client, sha, _ = anno_client
    r = client.put(f"/api/cases/a/meta/{sha}", headers=CSRF,
                   json={"memo": "メモ書き", "category": "訴訟書類", "cho_offset": 44})
    assert r.status_code == 200
    docs = client.get("/api/cases/a/documents").json()["documents"]
    d = next(x for x in docs if x["sha256"] == sha)
    assert d["memo"] == "メモ書き"
    assert d["category"] == "訴訟書類"
    assert d["cho_offset"] == 44


def test_has_annotations_flag(anno_client):
    client, sha, _ = anno_client
    client.put(f"/api/cases/a/annotations/{sha}", headers=CSRF,
               json={"base_updated_at": "", "annotations": [{"id": "1", "type": "rect", "page": 1, "rect": [0, 0, 1, 1]}]})
    docs = client.get("/api/cases/a/documents").json()["documents"]
    d = next(x for x in docs if x["sha256"] == sha)
    assert d["has_annotations"] is True


def test_export_bakes_annotations(anno_client):
    client, sha, _ = anno_client
    anns = [
        {"id": "r", "type": "rect", "page": 1, "rect": [50, 50, 200, 150], "color": "#ff0000"},
        {"id": "t", "type": "text", "page": 2, "point": [60, 700], "text": "注記テスト"},
    ]
    client.put(f"/api/cases/a/annotations/{sha}", headers=CSRF,
               json={"base_updated_at": "", "annotations": anns})
    r = client.get(f"/api/cases/a/export/{sha}")
    assert r.status_code == 200
    assert r.headers["content-type"] == "application/pdf"
    assert r.content[:4] == b"%PDF"
    # ページ指定（1ページのみ）
    r2 = client.get(f"/api/cases/a/export/{sha}", params={"pages": "1"})
    from pypdf import PdfReader
    import io
    assert len(PdfReader(io.BytesIO(r2.content)).pages) == 1


def test_orphan_and_relink(anno_client):
    client, sha, case_dir = anno_client
    orphan_sha = "f" * 64
    # documents に無い sha の注釈を作る
    client.put(f"/api/cases/a/annotations/{orphan_sha}", headers=CSRF,
               json={"base_updated_at": "", "annotations": [{"id": "o", "type": "rect", "page": 1, "rect": [0, 0, 1, 1]}]})
    orphans = client.get("/api/cases/a/orphan-annotations").json()["orphans"]
    assert any(o["sha256"] == orphan_sha for o in orphans)
    # 実在文書へ紐付け直す
    r = client.post(f"/api/cases/a/annotations/{orphan_sha}/relink", headers=CSRF,
                    json={"target_sha": sha})
    assert r.status_code == 200
    got = client.get(f"/api/cases/a/annotations/{sha}").json()
    assert any(a["id"] == "o" for a in got["annotations"])
    # 孤児は解消
    orphans2 = client.get("/api/cases/a/orphan-annotations").json()["orphans"]
    assert not any(o["sha256"] == orphan_sha for o in orphans2)


def test_other_user_annotations_readonly(anno_client, monkeypatch):
    client, sha, case_dir = anno_client
    # 別ユーザーの注釈ファイルを直接配置
    import json as _json
    other_dir = case_dir / "_index" / "annotations" / "otheruser"
    other_dir.mkdir(parents=True)
    (other_dir / f"{sha}.json").write_text(
        _json.dumps({"updated_at": "2026-01-01T00:00:00+00:00",
                     "annotations": [{"id": "z", "type": "rect", "page": 1, "rect": [0, 0, 5, 5]}]}),
        encoding="utf-8")
    got = client.get(f"/api/cases/a/annotations/{sha}").json()
    other = [a for a in got["annotations"] if a["_user"] == "otheruser"]
    assert other and other[0]["_editable"] is False


# ---- 事件管理API（追加・削除・フォルダ選択・再索引ログ） -----------------

CSRF = {"X-Kiroku-Viewer": "1"}


@pytest.fixture
def mgmt_client(tmp_path: Path, monkeypatch):
    """実 cases.json を保護しつつ管理APIをテストするフィクスチャ。"""
    monkeypatch.setattr(server, "CONFIG_PATH", tmp_path / "cases.json")
    server.CONFIG = {
        "cases": {
            "existing": {
                "id": "existing",
                "name": "既存事件",
                "path": (tmp_path / "existing_case").resolve(),
                "reindex": False,
            }
        },
        "display_names": {"nsato": "佐藤"},
    }
    (tmp_path / "existing_case").mkdir()
    server._listing_cache.clear()
    return TestClient(server.app, base_url="http://127.0.0.1:8788"), tmp_path


def test_add_case_ok(mgmt_client, tmp_path):
    client, tdir = mgmt_client
    new_case = tdir / "new_case"
    new_case.mkdir()
    r = client.post("/api/cases", headers=CSRF, json={"name": "新事件", "path": str(new_case), "reindex": False})
    assert r.status_code == 200, r.text
    data = r.json()
    assert data["name"] == "新事件"
    cid = data["id"]
    # GET /api/cases に現れること
    cases = client.get("/api/cases").json()["cases"]
    assert any(c["id"] == cid for c in cases)
    # cases.json に永続化されていること
    cfg_path = tdir / "cases.json"
    assert cfg_path.exists()
    import json as _json
    saved = _json.loads(cfg_path.read_text())
    assert any(c["id"] == cid for c in saved["cases"])


def test_add_case_nonexistent_path(mgmt_client):
    client, tdir = mgmt_client
    r = client.post("/api/cases", headers=CSRF,
                    json={"name": "", "path": str(tdir / "no_such_dir"), "reindex": False})
    assert r.status_code == 400


def test_add_case_duplicate(mgmt_client):
    client, tdir = mgmt_client
    new_case = tdir / "dup_case"
    new_case.mkdir()
    r1 = client.post("/api/cases", headers=CSRF, json={"name": "dup", "path": str(new_case), "reindex": False})
    assert r1.status_code == 200
    # 末尾スラッシュを付けても重複扱い
    r2 = client.post("/api/cases", headers=CSRF, json={"name": "dup2", "path": str(new_case) + "/", "reindex": False})
    assert r2.status_code == 409


def test_add_case_symlink_duplicate(mgmt_client, tmp_path):
    client, tdir = mgmt_client
    real = tdir / "real_case"
    real.mkdir()
    link = tdir / "link_case"
    link.symlink_to(real)
    r1 = client.post("/api/cases", headers=CSRF, json={"name": "real", "path": str(real), "reindex": False})
    assert r1.status_code == 200
    r2 = client.post("/api/cases", headers=CSRF, json={"name": "link", "path": str(link), "reindex": False})
    assert r2.status_code == 409


def test_delete_case_ok(mgmt_client):
    client, tdir = mgmt_client
    new_case = tdir / "del_case"
    new_case.mkdir()
    add_r = client.post("/api/cases", headers=CSRF, json={"name": "削除テスト", "path": str(new_case), "reindex": False})
    cid = add_r.json()["id"]
    del_r = client.delete(f"/api/cases/{cid}", headers=CSRF)
    assert del_r.status_code == 200
    cases = client.get("/api/cases").json()["cases"]
    assert not any(c["id"] == cid for c in cases)
    import json as _json
    saved = _json.loads((tdir / "cases.json").read_text())
    assert not any(c["id"] == cid for c in saved["cases"])


def test_delete_case_unknown_404(mgmt_client):
    client, _ = mgmt_client
    r = client.delete("/api/cases/no_such_id", headers=CSRF)
    assert r.status_code == 404


def test_delete_while_reindexing_409(mgmt_client, monkeypatch):
    client, tdir = mgmt_client
    new_case = tdir / "reindex_case"
    new_case.mkdir()
    add_r = client.post("/api/cases", headers=CSRF, json={"name": "再索引中", "path": str(new_case), "reindex": False})
    cid = add_r.json()["id"]

    class FakeProc:
        def poll(self):
            return None  # 実行中を模擬

    server._reindex_procs[cid] = FakeProc()
    try:
        r = client.delete(f"/api/cases/{cid}", headers=CSRF)
        assert r.status_code == 409
    finally:
        server._reindex_procs.pop(cid, None)


def test_add_case_requires_csrf(mgmt_client, tmp_path):
    client, tdir = mgmt_client
    new_case = tdir / "csrf_case"
    new_case.mkdir()
    r = client.post("/api/cases", json={"name": "x", "path": str(new_case), "reindex": False})
    assert r.status_code == 403


def test_add_case_default_name_from_parent(mgmt_client, tmp_path):
    """__Document__ 配下を指定した場合、親フォルダ名が事件名になる。"""
    client, tdir = mgmt_client
    parent = tdir / "高山岩男[弁護革命system]"
    parent.mkdir()
    doc_dir = parent / "__Document__"
    doc_dir.mkdir()
    r = client.post("/api/cases", headers=CSRF, json={"name": "", "path": str(doc_dir), "reindex": False})
    assert r.status_code == 200, r.text
    assert r.json()["name"] == "高山岩男"


def test_display_names_preserved_after_add_delete(mgmt_client, tmp_path):
    """add/delete を経ても display_names が保持される。"""
    client, tdir = mgmt_client
    import json as _json
    new_case = tdir / "dn_case"
    new_case.mkdir()
    add_r = client.post("/api/cases", headers=CSRF, json={"name": "dn", "path": str(new_case), "reindex": False})
    cid = add_r.json()["id"]
    client.delete(f"/api/cases/{cid}", headers=CSRF)
    saved = _json.loads((tdir / "cases.json").read_text())
    assert saved["display_names"].get("nsato") == "佐藤"


def test_existing_case_id_and_reindex_unchanged(mgmt_client, tmp_path):
    """既存事件の id/reindex が add/delete を経ても不変。"""
    client, tdir = mgmt_client
    new_case = tdir / "extra_case"
    new_case.mkdir()
    client.post("/api/cases", headers=CSRF, json={"name": "extra", "path": str(new_case), "reindex": False})
    import json as _json
    saved = _json.loads((tdir / "cases.json").read_text())
    existing = next(c for c in saved["cases"] if c["id"] == "existing")
    assert existing["reindex"] is False


def test_api_cases_has_index_field(mgmt_client, tmp_path):
    """GET /api/cases の has_index フィールドが正しく返る。"""
    client, tdir = mgmt_client
    # 索引なし → False
    cases = client.get("/api/cases").json()["cases"]
    ex = next(c for c in cases if c["id"] == "existing")
    assert ex["has_index"] is False
    # 索引を作って → True
    import server as sv
    export_dir = sv.CONFIG["cases"]["existing"]["path"] / server.EXPORT_REL
    export_dir.parent.mkdir(parents=True, exist_ok=True)
    export_dir.touch()
    cases2 = client.get("/api/cases").json()["cases"]
    ex2 = next(c for c in cases2 if c["id"] == "existing")
    assert ex2["has_index"] is True


def test_add_case_from_empty_config(tmp_path, monkeypatch):
    """CONFIG_PATH 不在からの初回 add でサンプル事件が混入しない。"""
    monkeypatch.setattr(server, "CONFIG_PATH", tmp_path / "cases.json")
    monkeypatch.setattr(server, "CONFIG", {"cases": {}, "display_names": {}})
    server._listing_cache.clear()
    c = TestClient(server.app, base_url="http://127.0.0.1:8788")
    new_case = tmp_path / "first_case"
    new_case.mkdir()
    r = c.post("/api/cases", headers=CSRF, json={"name": "初回事件", "path": str(new_case), "reindex": False})
    assert r.status_code == 200
    cases = c.get("/api/cases").json()["cases"]
    assert len(cases) == 1
    assert cases[0]["name"] == "初回事件"


def test_pick_folder_normal(mgmt_client, monkeypatch):
    import subprocess
    client, _ = mgmt_client

    class FakeProc:
        returncode = 0
        stdout = "/tmp/test_folder\n"

    monkeypatch.setattr(subprocess, "run", lambda *a, **kw: FakeProc())
    r = client.post("/api/pick-folder", headers=CSRF)
    assert r.status_code == 200
    assert r.json()["path"] == "/tmp/test_folder"


def test_pick_folder_cancelled(mgmt_client, monkeypatch):
    import subprocess
    client, _ = mgmt_client

    class FakeCancelled:
        returncode = 1
        stdout = ""

    monkeypatch.setattr(subprocess, "run", lambda *a, **kw: FakeCancelled())
    r = client.post("/api/pick-folder", headers=CSRF)
    assert r.status_code == 200
    assert r.json()["cancelled"] is True


def test_reindex_log_empty(mgmt_client):
    client, tdir = mgmt_client
    r = client.get("/api/cases/existing/reindex/log")
    assert r.status_code == 200
    assert r.json()["log"] == ""
