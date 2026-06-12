#!/usr/bin/env python3
"""記録ビューア（弁護革命代替）ローカルサーバ。

FastAPI で索引DB（読み取り専用）・PDFストリーム・静的配信を提供する。
設計判断 D1〜D13（docs/kiroku_viewer_plan.md）に準拠する。

- 索引DBは Drive 上で直接 open しない（D3）。リクエスト時にローカルキャッシュへ
  コピーして immutable 接続する。
- 文書識別子は全 API で sha256 に統一（D2）。
- 未索引PDF・非PDFもフォルダ走査でマージ表示（D5）。
- localhost 限定・Host ヘッダ検証・CSRF ヘッダ必須（D11）。
"""

from __future__ import annotations

import getpass
import json
import mimetypes
import os
import sqlite3
import sys
import time
from pathlib import Path

from fastapi import FastAPI, HTTPException, Request, Response
from fastapi.responses import FileResponse, JSONResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles

# scripts/evidence_index.py のヘルパーを再利用する。
HERE = Path(__file__).resolve().parent
REPO_ROOT = HERE.parents[1]
sys.path.insert(0, str(REPO_ROOT / "scripts"))
import evidence_index as ei  # noqa: E402

# .mjs を JS として配信（D6）。
mimetypes.add_type("text/javascript", ".mjs")

# trigram トークナイザ必須（D1 #1）。
assert sqlite3.sqlite_version_info >= (3, 34, 0), (
    f"SQLite 3.34.0+ が必要です（trigram FTS）。現在: {sqlite3.sqlite_version}"
)

HOST = "127.0.0.1"
PORT = 8788
ALLOWED_HOSTS = {f"localhost:{PORT}", f"127.0.0.1:{PORT}"}
CSRF_HEADER = "x-kiroku-viewer"
INDEX_DIR_NAME = "_index"
EXPORT_REL = f"{INDEX_DIR_NAME}/export/evidence_index.sqlite"
MEDIA_EXTS = {".mp4", ".mov", ".m4a", ".jpg", ".jpeg", ".png"}
OPENABLE_EXTS = MEDIA_EXTS | {".pdf"}
OCR_LOW_CONFIDENCE = 60.0
LISTING_TTL_SECONDS = 5.0

CACHE_ROOT = Path.home() / "Library" / "Caches" / "kiroku-viewer"

CONFIG_PATH = HERE / "cases.json"
EXAMPLE_CONFIG_PATH = HERE / "cases.example.json"


# --------------------------------------------------------------------------
# 設定（cases.json）
# --------------------------------------------------------------------------

def load_config() -> dict:
    path = CONFIG_PATH if CONFIG_PATH.exists() else EXAMPLE_CONFIG_PATH
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise RuntimeError(f"cases 設定を読めません: {path}: {exc}") from exc
    cases = {}
    for c in data.get("cases", []):
        cid = str(c["id"])
        cases[cid] = {
            "id": cid,
            "name": c.get("name", cid),
            "path": Path(c["path"]).expanduser(),
            "reindex": bool(c.get("reindex", False)),
        }
    return {"cases": cases, "display_names": data.get("display_names", {})}


def get_case(config: dict, case_id: str) -> dict:
    case = config["cases"].get(case_id)
    if case is None:
        raise HTTPException(status_code=404, detail=f"unknown case: {case_id}")
    return case


def current_user() -> str:
    return getpass.getuser()


def display_name(config: dict, user: str) -> str:
    return config["display_names"].get(user, user)


# --------------------------------------------------------------------------
# 索引DBのローカルキャッシュ（D3）
# --------------------------------------------------------------------------

def cached_db_path(case: dict) -> Path | None:
    """Drive側DBを必要時のみローカルキャッシュへコピーし、キャッシュパスを返す。

    Drive側DBは一切ロック・書き込みしない。mtime/サイズ変化時のみ
    .tmp→os.replace でコピーする。索引DBが無ければ None。
    """
    src = case["path"] / EXPORT_REL
    if not src.exists():
        return None
    try:
        st = src.stat()
    except OSError:
        return None
    dst_dir = CACHE_ROOT / case["id"]
    dst_dir.mkdir(parents=True, exist_ok=True)
    dst = dst_dir / "evidence_index.sqlite"
    needs_copy = True
    if dst.exists():
        dstat = dst.stat()
        # 元と同じ mtime・サイズなら再コピー不要。
        if dstat.st_size == st.st_size and int(dstat.st_mtime) == int(st.st_mtime):
            needs_copy = False
    if needs_copy:
        tmp = dst.with_name(dst.name + ".tmp")
        import shutil

        shutil.copy2(src, tmp)
        os.replace(tmp, dst)
    return dst


def open_db_ro(db_path: Path) -> sqlite3.Connection:
    """読み取り専用・immutable 接続（WAL/shm を触らない・D3）。"""
    uri = f"file:{db_path}?mode=ro&immutable=1"
    conn = sqlite3.connect(uri, uri=True)
    conn.row_factory = sqlite3.Row
    return conn


# --------------------------------------------------------------------------
# 文書一覧（D2/D4/D5）
# --------------------------------------------------------------------------

_listing_cache: dict[str, tuple[float, list[dict], dict[str, Path]]] = {}
_reindex_procs: dict = {}


def safe_resolve(case: dict, rel_path: str) -> Path:
    """rel_path を事件フォルダ配下に解決し、配下であることを検証する（D2/D11）。"""
    base = case["path"].resolve()
    target = (base / rel_path).resolve()
    if base not in target.parents and target != base:
        raise HTTPException(status_code=403, detail="path escapes case folder")
    return target


def iter_case_files(case: dict):
    """事件フォルダ配下の PDF・メディアを走査（_index/ 除外・D5）。"""
    base = case["path"]
    index_dir = base / INDEX_DIR_NAME
    for root, dirs, files in os.walk(base):
        root_path = Path(root)
        # _index/ 配下はスキップ
        try:
            root_path.relative_to(index_dir)
            continue
        except ValueError:
            pass
        # _index ディレクトリ自体への降下を止める
        dirs[:] = [d for d in dirs if (root_path / d) != index_dir]
        for name in files:
            if name.startswith("."):
                continue
            ext = Path(name).suffix.lower()
            if ext == ".pdf" or ext in MEDIA_EXTS:
                yield root_path / name


def build_listing(case: dict) -> tuple[list[dict], dict[str, Path]]:
    """索引DB由来＋未索引フォルダ走査をマージした文書一覧を返す。

    返り値: (documents, sha256→絶対パス の解決マップ)
    """
    base = case["path"]
    docs: list[dict] = []
    sha_to_path: dict[str, Path] = {}
    indexed_rel: set[str] = set()

    db_path = cached_db_path(case)
    if db_path is not None:
        conn = open_db_ro(db_path)
        try:
            rows = conn.execute(
                """
                SELECT d.id, d.rel_path, d.file_name, d.evidence_no, d.title,
                       d.document_date, d.person_or_source, d.sha256,
                       d.page_count, d.extract_status,
                       (SELECT MIN(p.ocr_confidence) FROM pages p
                        WHERE p.document_id = d.id AND p.ocr_confidence IS NOT NULL)
                         AS min_conf,
                       (SELECT COUNT(*) FROM pages p
                        WHERE p.document_id = d.id AND p.extract_method
                          IN ('apple-vision','tesseract')) AS ocr_pages
                FROM documents d
                """
            ).fetchall()
        finally:
            conn.close()
        for r in rows:
            rel = r["rel_path"]
            indexed_rel.add(rel)
            target = (base / rel).resolve()
            sha_to_path[r["sha256"]] = target
            min_conf = r["min_conf"]
            low_conf = (
                (min_conf is not None and min_conf < OCR_LOW_CONFIDENCE)
                or r["extract_status"] == ei.STATUS_NEEDS_REVIEW
            )
            docs.append(
                {
                    "sha256": r["sha256"],
                    "evidence_no": r["evidence_no"],
                    "title": r["title"],
                    "document_date": r["document_date"],
                    "author": r["person_or_source"],
                    "rel_path": rel,
                    "file_name": r["file_name"],
                    "page_count": r["page_count"],
                    "kind": "pdf",
                    "indexed": True,
                    "ocr_low_confidence": bool(low_conf),
                    "ocr_pages": r["ocr_pages"],
                }
            )

    # 未索引ファイル（PDF・メディア）をマージ（D5）。
    for fpath in iter_case_files(case):
        rel = fpath.relative_to(base).as_posix()
        if rel in indexed_rel:
            continue
        sha = ei.sha256_file(fpath)
        sha_to_path[sha] = fpath.resolve()
        ext = fpath.suffix.lower()
        kind = "pdf" if ext == ".pdf" else "media"
        evidence_no, title, doc_date, author = ei.parse_name_only(fpath.stem)
        docs.append(
            {
                "sha256": sha,
                "evidence_no": evidence_no,
                "title": title,
                "document_date": doc_date,
                "author": author,
                "rel_path": rel,
                "file_name": fpath.name,
                "page_count": None,
                "kind": kind,
                "indexed": False,
                "ocr_low_confidence": False,
                "ocr_pages": 0,
            }
        )

    docs.sort(key=lambda d: ei.evidence_sort_key(d["evidence_no"]))
    return docs, sha_to_path


def get_listing(case: dict) -> tuple[list[dict], dict[str, Path]]:
    now = time.monotonic()
    cached = _listing_cache.get(case["id"])
    if cached and (now - cached[0]) < LISTING_TTL_SECONDS:
        return cached[1], cached[2]
    docs, sha_to_path = build_listing(case)
    _listing_cache[case["id"]] = (now, docs, sha_to_path)
    return docs, sha_to_path


def resolve_sha(case: dict, sha256: str) -> Path:
    _docs, sha_to_path = get_listing(case)
    target = sha_to_path.get(sha256)
    if target is None:
        # キャッシュ失効の可能性。強制再構築して再試行。
        _listing_cache.pop(case["id"], None)
        _docs, sha_to_path = get_listing(case)
        target = sha_to_path.get(sha256)
    if target is None:
        raise HTTPException(status_code=404, detail="unknown document")
    # 事件フォルダ配下であることを再検証（D2/D11）。
    base = case["path"].resolve()
    if base not in target.parents and target != base:
        raise HTTPException(status_code=403, detail="path escapes case folder")
    if not target.exists():
        raise HTTPException(status_code=404, detail="file missing")
    return target


# --------------------------------------------------------------------------
# Range 配信（D12）
# --------------------------------------------------------------------------

def parse_range(range_header: str, file_size: int) -> tuple[int, int] | None:
    """`bytes=start-end` を解釈。不正・未対応は None（呼び出し側で 200 全送）。"""
    if not range_header or not range_header.startswith("bytes="):
        return None
    spec = range_header[len("bytes="):].split(",")[0].strip()
    if "-" not in spec:
        return None
    start_s, _, end_s = spec.partition("-")
    try:
        if start_s == "":
            # 末尾 N バイト
            n = int(end_s)
            if n <= 0:
                return None
            start = max(0, file_size - n)
            end = file_size - 1
        else:
            start = int(start_s)
            end = int(end_s) if end_s else file_size - 1
    except ValueError:
        return None
    if start > end or start >= file_size:
        return None
    end = min(end, file_size - 1)
    return start, end


def range_response(path: Path, request: Request, content_type: str) -> Response:
    file_size = path.stat().st_size
    range_header = request.headers.get("range", "")
    rng = parse_range(range_header, file_size)
    if rng is None:
        # Range 無し・不正 → 200 全送（Accept-Ranges を必ず付ける）。
        return FileResponse(
            path,
            media_type=content_type,
            headers={"Accept-Ranges": "bytes"},
        )
    start, end = rng
    length = end - start + 1
    chunk = 256 * 1024

    def streamer():
        with path.open("rb") as f:
            f.seek(start)
            remaining = length
            while remaining > 0:
                data = f.read(min(chunk, remaining))
                if not data:
                    break
                remaining -= len(data)
                yield data

    headers = {
        "Content-Range": f"bytes {start}-{end}/{file_size}",
        "Accept-Ranges": "bytes",
        "Content-Length": str(length),
    }
    return StreamingResponse(
        streamer(), status_code=206, media_type=content_type, headers=headers
    )


# --------------------------------------------------------------------------
# アプリ
# --------------------------------------------------------------------------

app = FastAPI(title="記録ビューア", docs_url=None, redoc_url=None, openapi_url=None)
CONFIG = load_config()


@app.middleware("http")
async def security_middleware(request: Request, call_next):
    # Host ヘッダ検証（D11）。
    host = request.headers.get("host", "")
    if host not in ALLOWED_HOSTS:
        return JSONResponse({"detail": "host not allowed"}, status_code=403)
    # POST/PUT/DELETE は CSRF ヘッダ必須（D11）。
    if request.method in {"POST", "PUT", "DELETE", "PATCH"}:
        if request.headers.get(CSRF_HEADER) != "1":
            return JSONResponse({"detail": "missing csrf header"}, status_code=403)
    return await call_next(request)


@app.get("/api/cases")
def api_cases():
    return {
        "cases": [
            {"id": c["id"], "name": c["name"], "reindex": c["reindex"]}
            for c in CONFIG["cases"].values()
        ],
        "user": current_user(),
        "user_display": display_name(CONFIG, current_user()),
    }


@app.get("/api/cases/{case_id}/documents")
def api_documents(case_id: str):
    case = get_case(CONFIG, case_id)
    docs, _ = get_listing(case)
    has_index = (case["path"] / EXPORT_REL).exists()
    return {"documents": docs, "has_index": has_index, "reindex": case["reindex"]}


@app.get("/api/cases/{case_id}/pdf/{sha256}")
def api_pdf(case_id: str, sha256: str, request: Request):
    case = get_case(CONFIG, case_id)
    path = resolve_sha(case, sha256)
    ext = path.suffix.lower()
    ctype = "application/pdf" if ext == ".pdf" else (
        mimetypes.guess_type(path.name)[0] or "application/octet-stream"
    )
    return range_response(path, request, ctype)


@app.get("/api/cases/{case_id}/search")
def api_search(case_id: str, q: str = "", filter: str = "", limit: int = 100):
    """本文検索（D1）。snippet 付きでページ単位ヒットを返す。

    filter はタイトル・符号への正規化部分一致で結果を絞り込む。
    """
    case = get_case(CONFIG, case_id)
    db_path = cached_db_path(case)
    if db_path is None or not q.strip():
        return {"hits": [], "indexed": db_path is not None}
    conn = open_db_ro(db_path)
    try:
        raw_limit = 500 if filter.strip() else limit
        hits = ei.search_pages(conn, q, raw_limit)
    finally:
        conn.close()
    if filter.strip():
        nf = ei.normalize_for_search(filter)
        hits = [
            h for h in hits
            if nf in ei.normalize_for_search(f"{h['evidence_no']} {h['title']}")
        ]
        hits = hits[:limit]
    return {"hits": hits, "indexed": True}


@app.get("/api/cases/{case_id}/text/{sha256}")
def api_text(case_id: str, sha256: str, page: int | None = None):
    """抽出テキスト（ページ単位／D6 の TXT 表示用）。"""
    case = get_case(CONFIG, case_id)
    db_path = cached_db_path(case)
    if db_path is None:
        return {"indexed": False, "page_count": 0, "pages": []}
    conn = open_db_ro(db_path)
    try:
        doc = conn.execute(
            "SELECT id, page_count FROM documents WHERE sha256 = ? LIMIT 1", (sha256,)
        ).fetchone()
        if doc is None:
            return {"indexed": False, "page_count": 0, "pages": []}
        doc_id = doc["id"]
        page_count = doc["page_count"]
        if page is not None:
            row = conn.execute(
                "SELECT page_no, text FROM pages WHERE document_id = ? AND page_no = ?",
                (doc_id, page),
            ).fetchone()
            return {
                "indexed": True,
                "page_count": page_count,
                "page_no": page,
                "text": row["text"] if row else "",
            }
        rows = conn.execute(
            "SELECT page_no, text FROM pages WHERE document_id = ? ORDER BY page_no",
            (doc_id,),
        ).fetchall()
    finally:
        conn.close()
    return {
        "indexed": True,
        "page_count": page_count,
        "pages": [{"page_no": r["page_no"], "text": r["text"]} for r in rows],
    }


@app.post("/api/cases/{case_id}/open-file/{sha256}")
def api_open_file(case_id: str, sha256: str):
    """ファイルを既定アプリで開く（mp4 は QuickTime 等／D11 検証付き）。"""
    case = get_case(CONFIG, case_id)
    path = resolve_sha(case, sha256)
    if path.suffix.lower() not in OPENABLE_EXTS:
        raise HTTPException(status_code=403, detail="extension not allowed")
    import subprocess

    try:
        subprocess.run(["open", str(path)], check=True)
    except (subprocess.CalledProcessError, OSError) as exc:
        raise HTTPException(status_code=500, detail=f"open failed: {exc}") from exc
    return {"ok": True, "file_name": path.name}


@app.post("/api/cases/{case_id}/reindex")
def api_reindex(case_id: str):
    """索引を再構築（evidence_index.py build をサブプロセス起動・多重起動拒否）。"""
    case = get_case(CONFIG, case_id)
    if not case["reindex"]:
        raise HTTPException(status_code=403, detail="reindex disabled for this case")
    proc = _reindex_procs.get(case_id)
    if proc is not None and proc.poll() is None:
        raise HTTPException(status_code=409, detail="reindex already running")
    import subprocess

    script = REPO_ROOT / "scripts" / "evidence_index.py"
    cmd = [
        sys.executable, str(script), "build", "--ocr",
        "--evidence-dir", str(case["path"]),
    ]
    log_dir = CACHE_ROOT / case_id
    log_dir.mkdir(parents=True, exist_ok=True)
    log_file = (log_dir / "reindex.log").open("w", encoding="utf-8")
    _reindex_procs[case_id] = subprocess.Popen(
        cmd, stdout=log_file, stderr=subprocess.STDOUT
    )
    _listing_cache.pop(case_id, None)
    return {"started": True}


@app.get("/api/cases/{case_id}/reindex")
def api_reindex_status(case_id: str):
    get_case(CONFIG, case_id)
    proc = _reindex_procs.get(case_id)
    if proc is None:
        return {"running": False, "ran": False}
    running = proc.poll() is None
    return {"running": running, "ran": True, "returncode": proc.returncode}


# 静的配信（SPA）。最後にマウントして API ルートを優先する。
app.mount("/static", StaticFiles(directory=str(HERE / "static")), name="static")


@app.get("/")
def index():
    return FileResponse(HERE / "static" / "index.html")


def main():
    import uvicorn

    uvicorn.run(app, host=HOST, port=PORT, log_level="info")


if __name__ == "__main__":
    main()
