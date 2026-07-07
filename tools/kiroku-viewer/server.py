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
import hashlib
import json
import logging
import mimetypes
import os
import re
import sqlite3
import sys
import threading
import time
from pathlib import Path

from fastapi import Body, FastAPI, File, HTTPException, Request, Response, UploadFile
from fastapi.responses import FileResponse, JSONResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles

log = logging.getLogger("kiroku-viewer")

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
ANNOTATIONS_DIR = "annotations"
VIEWER_META_DIR = "viewer_meta"
SHA_RE = re.compile(r"^[0-9a-f]{64}\.json$")
USER_RE = re.compile(r"^[A-Za-z0-9._-]+$")
META_FIELDS = ("memo", "category", "cho_offset", "rotation", "evidence_no")
# 他ユーザー注釈の色分け（自ユーザーは固定色、他は順に割当）。
USER_COLORS = ["#16a34a", "#9333ea", "#ea580c", "#0891b2", "#be123c"]
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

_config_lock = threading.Lock()


def load_config() -> dict:
    """cases.json を読み込む。不在時は cases 空・display_names のみ example から継承。"""
    if CONFIG_PATH.exists():
        try:
            data = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError) as exc:
            raise RuntimeError(f"cases 設定を読めません: {CONFIG_PATH}: {exc}") from exc
        cases_src = data.get("cases", [])
        names = data.get("display_names", {})
    else:
        # 初回起動: 事件は空。表示名の雛形だけ example から引き継ぐ。
        try:
            names = json.loads(EXAMPLE_CONFIG_PATH.read_text(encoding="utf-8")).get("display_names", {})
        except (OSError, json.JSONDecodeError):
            names = {}
        cases_src = []
    cases = {}
    for c in cases_src:
        cid = str(c["id"])
        cases[cid] = {
            "id": cid,
            "name": c.get("name", cid),
            "path": Path(c["path"]).expanduser().resolve(),
            "reindex": bool(c.get("reindex", False)),
            "upload_subdir": c.get("upload_subdir", ""),
        }
    return {"cases": cases, "display_names": names}


def save_cases_config() -> None:
    """CONFIG を cases.json へアトミック書き込み（.tmp→os.replace）。
    呼び出し前に _config_lock を取得すること。
    """
    data = {
        "cases": [
            {
                "id": c["id"],
                "name": c["name"],
                "path": str(c["path"]),
                "reindex": c["reindex"],
                **({"upload_subdir": c["upload_subdir"]} if c.get("upload_subdir") else {}),
            }
            for c in CONFIG["cases"].values()
        ],
        "display_names": CONFIG["display_names"],
    }
    atomic_write_json(CONFIG_PATH, data)


def gen_case_id(resolved_path: Path) -> str:
    """フォルダパスの SHA1 先頭8文字をケースIDにする。"""
    return hashlib.sha1(str(resolved_path).encode("utf-8")).hexdigest()[:8]


def _default_case_name(target: Path) -> str:
    """事件名の既定値: アンダースコア始まりのフォルダは親名から生成。"""
    name = target.name
    if name.startswith("_"):
        name = target.parent.name
    # [弁護革命system] 等のブラケット以降を除去。
    name = re.sub(r"[\[［].*$", "", name).strip()
    return name or target.name


def get_case(config: dict, case_id: str) -> dict:
    case = config["cases"].get(case_id)
    if case is None:
        raise HTTPException(status_code=404, detail=f"unknown case: {case_id}")
    return case


def current_user() -> str:
    return getpass.getuser()


def display_name(config: dict, user: str) -> str:
    return config["display_names"].get(user, user)


def now_iso() -> str:
    return ei.now_iso()


# --------------------------------------------------------------------------
# 注釈・メタの永続化（D8 座標・D9 規律）
# --------------------------------------------------------------------------

def atomic_write_json(path: Path, data) -> None:
    """`.tmp`→`os.replace` のアトミック書き込み（D9）。"""
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_name(path.name + ".tmp")
    tmp.write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8")
    os.replace(tmp, path)


def validate_sha(sha256: str) -> str:
    if not re.fullmatch(r"[0-9a-f]{64}", sha256):
        raise HTTPException(status_code=400, detail="invalid sha256")
    return sha256


def annotations_dir(case: dict, user: str) -> Path:
    if not USER_RE.match(user):
        raise HTTPException(status_code=400, detail="invalid user")
    return case["path"] / INDEX_DIR_NAME / ANNOTATIONS_DIR / user


def annotation_file(case: dict, user: str, sha256: str) -> Path:
    return annotations_dir(case, user) / f"{validate_sha(sha256)}.json"


def iter_annotation_users(case: dict):
    base = case["path"] / INDEX_DIR_NAME / ANNOTATIONS_DIR
    if not base.is_dir():
        return
    for child in base.iterdir():
        if child.is_dir() and USER_RE.match(child.name):
            yield child.name


def read_annotations(case: dict, sha256: str, current: str) -> dict:
    """全ユーザーの注釈をマージして返す（読みはマージ・D9）。

    自ユーザー分は editable=True、他ユーザー分は表示のみ（色分け）。
    Drive 競合コピー等、厳密パターンに合わないファイルは無視してログ警告。
    """
    validate_sha(sha256)
    merged: list[dict] = []
    own = {"updated_at": "", "annotations": []}
    color_idx = 0
    for user in sorted(iter_annotation_users(case)):
        fpath = annotations_dir(case, user) / f"{sha256}.json"
        if not fpath.exists():
            continue
        if not SHA_RE.match(fpath.name):
            log.warning("ignoring non-conforming annotation file: %s", fpath)
            continue
        try:
            payload = json.loads(fpath.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            log.warning("unreadable annotation file: %s", fpath)
            continue
        anns = payload.get("annotations", [])
        is_own = user == current
        color = "#2563eb" if is_own else USER_COLORS[color_idx % len(USER_COLORS)]
        if not is_own:
            color_idx += 1
        for a in anns:
            item = dict(a)
            item["_user"] = user
            item["_user_display"] = display_name(CONFIG, user)
            item["_editable"] = is_own
            if not is_own and "color" not in item:
                item["color"] = color
            merged.append(item)
        if is_own:
            own = payload
    return {"annotations": merged, "own_updated_at": own.get("updated_at", "")}


def read_merged_meta(case: dict) -> dict:
    """全ユーザーの viewer_meta をマージ（同一sha・同一フィールドは新しい方／D9）。"""
    base = case["path"] / INDEX_DIR_NAME / VIEWER_META_DIR
    result: dict[str, dict] = {}
    if not base.is_dir():
        return result
    for fpath in base.glob("*.json"):
        if not USER_RE.match(fpath.stem):
            log.warning("ignoring non-conforming meta file: %s", fpath)
            continue
        try:
            data = json.loads(fpath.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            log.warning("unreadable meta file: %s", fpath)
            continue
        for sha, fields in data.items():
            if not re.fullmatch(r"[0-9a-f]{64}", sha):
                continue
            cur = result.setdefault(sha, {})
            for field in META_FIELDS:
                if field not in fields:
                    continue
                ts = fields.get(f"{field}_at", "")
                prev_ts = cur.get(f"{field}_at", "")
                if ts >= prev_ts:
                    cur[field] = fields[field]
                    cur[f"{field}_at"] = ts
    return result


def annotated_shas(case: dict) -> set[str]:
    """注釈ファイルが存在し中身が空でない sha256 の集合。"""
    base = case["path"] / INDEX_DIR_NAME / ANNOTATIONS_DIR
    found: set[str] = set()
    if not base.is_dir():
        return found
    for user in iter_annotation_users(case):
        for fpath in annotations_dir(case, user).iterdir():
            if not SHA_RE.match(fpath.name):
                if fpath.is_file():
                    log.warning("ignoring non-conforming annotation file: %s", fpath)
                continue
            try:
                payload = json.loads(fpath.read_text(encoding="utf-8"))
            except (OSError, json.JSONDecodeError):
                continue
            if payload.get("annotations"):
                found.add(fpath.stem)
    return found


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


def _sanitize_indexed_doc_fields(doc: dict) -> dict:
    """索引DB由来の文書について、弁護革命IDが含まれるファイル名を再パースして
    title / document_date / author を上書きする（再索引不要・D3整合）。"""
    file_name = doc.get("file_name", "")
    stem = Path(file_name).stem
    if ei.BENGOKAKUMEI_ID_RE.search(stem):
        ev, title, doc_date, author = ei.parse_name_only(stem)
        doc["title"] = title
        doc["document_date"] = doc_date
        doc["author"] = author
        # 符号は索引DBの値を優先（上書きしない）
    return doc


def build_listing(case: dict) -> tuple[list[dict], dict[str, Path]]:
    """索引DB由来＋未索引フォルダ走査をマージした文書一覧を返す。

    重複排除: 索引行のファイルが実在しない場合のみ、同 sha の walk 行を採用する。
    両ファイルが正規に実在する場合は両行を残す。

    返り値: (documents, sha256→絶対パス の解決マップ)
    """
    base = case["path"]
    sha_to_path: dict[str, Path] = {}
    indexed_rel: set[str] = set()
    meta = read_merged_meta(case)
    anno = annotated_shas(case)

    def apply_meta(doc: dict) -> dict:
        m = meta.get(doc["sha256"], {})
        if m.get("evidence_no"):
            doc["evidence_no"] = m["evidence_no"]
        doc["category"] = m.get("category", "")
        doc["memo"] = m.get("memo", "")
        doc["cho_offset"] = m.get("cho_offset", 0)
        doc["rotation"] = m.get("rotation", 0)
        doc["has_annotations"] = doc["sha256"] in anno
        return doc

    # sha256 → indexed doc (file が存在しない行)
    indexed_sha_missing: dict[str, dict] = {}
    # sha256 → indexed doc (file が存在する行)
    indexed_sha_existing: set[str] = set()
    indexed_docs: list[dict] = []

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
            doc = _sanitize_indexed_doc_fields(apply_meta(
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
            ))
            if target.exists():
                indexed_sha_existing.add(r["sha256"])
            else:
                indexed_sha_missing[r["sha256"]] = doc
            indexed_docs.append(doc)

    # 未索引ファイル（PDF・メディア）をマージ（D5）。
    # 同 sha で索引行ファイルが欠落している行は walk 行で置換する。
    replaced_by_walk: set[str] = set()
    walk_docs: list[dict] = []
    for fpath in iter_case_files(case):
        rel = fpath.relative_to(base).as_posix()
        if rel in indexed_rel:
            continue
        sha = ei.sha256_file(fpath)
        sha_to_path[sha] = fpath.resolve()
        ext = fpath.suffix.lower()
        kind = "pdf" if ext == ".pdf" else "media"
        evidence_no, title, doc_date, author = ei.parse_name_only(fpath.stem)
        walk_doc = apply_meta(
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
        if sha in indexed_sha_missing:
            # ファイル欠落の索引行を walk 行で置換
            replaced_by_walk.add(sha)
        walk_docs.append(walk_doc)

    # 索引行から「walk 行に置換された欠落行」を除外してマージ
    docs: list[dict] = [d for d in indexed_docs if d["sha256"] not in replaced_by_walk]
    docs.extend(walk_docs)
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
            {
                "id": c["id"],
                "name": c["name"],
                "reindex": c["reindex"],
                "path": str(c["path"]),
                "has_index": (c["path"] / EXPORT_REL).exists(),
            }
            for c in CONFIG["cases"].values()
        ],
        "user": current_user(),
        "user_display": display_name(CONFIG, current_user()),
    }


def _register_case(target: Path, name: str, reindex: bool, upload_subdir: str = "") -> dict:
    """target フォルダを事件として登録する共通ロジック（_config_lock 取得後に呼ぶこと）。"""
    for c in CONFIG["cases"].values():
        if c["path"] == target:
            raise HTTPException(
                status_code=409,
                detail=f"同じフォルダが事件「{c['name']}」として登録済みです",
            )
    cid = gen_case_id(target)
    base_id = cid
    suffix = 0
    while cid in CONFIG["cases"]:
        suffix += 1
        cid = f"{base_id}{suffix}"
    CONFIG["cases"][cid] = {
        "id": cid,
        "name": name,
        "path": target,
        "reindex": reindex,
        "upload_subdir": upload_subdir,
    }
    save_cases_config()
    return {
        "id": cid,
        "name": name,
        "reindex": reindex,
        "upload_subdir": upload_subdir,
        "has_index": (target / EXPORT_REL).exists(),
    }


@app.post("/api/cases")
def api_add_case(payload: dict = Body(...)):
    """事件フォルダを登録する。"""
    raw_path = payload.get("path", "").strip()
    if not raw_path:
        raise HTTPException(status_code=400, detail="path は必須です")
    target = Path(raw_path).expanduser().resolve()
    if not target.is_dir():
        raise HTTPException(status_code=400, detail=f"フォルダが見つかりません: {raw_path}")
    name = (payload.get("name") or "").strip() or _default_case_name(target)
    reindex = bool(payload.get("reindex", False))
    upload_subdir = (payload.get("upload_subdir") or "").strip()
    with _config_lock:
        return _register_case(target, name, reindex, upload_subdir)


_FOLDER_NAME_RE = re.compile(r'[/\\:\x00]')
_FOLDER_NAME_DOTDOT = re.compile(r'(?:^|[/\\])\.\.(?:[/\\]|$)')


@app.post("/api/cases/folders")
def api_create_case_folder(payload: dict = Body(...)):
    """新規フォルダを作成して事件として登録する。

    parent_path: 親フォルダの絶対パス
    folder_name: 作成するフォルダ名
    name: 事件名（省略時はフォルダ名から自動設定）
    reindex: 索引作成を許可するか
    upload_subdir: D&D 保存先サブフォルダ（省略時は事件フォルダ直下）
    """
    raw_parent = payload.get("parent_path", "").strip()
    folder_name = (payload.get("folder_name") or "").strip()

    if not raw_parent:
        raise HTTPException(status_code=400, detail="parent_path は必須です")
    if not folder_name:
        raise HTTPException(status_code=400, detail="folder_name は必須です")

    # folder_name バリデーション
    if _FOLDER_NAME_RE.search(folder_name):
        raise HTTPException(status_code=400, detail="フォルダ名にパス区切り文字または NUL は使えません")
    if folder_name.startswith("."):
        raise HTTPException(status_code=400, detail="フォルダ名は先頭ドット不可です")
    if folder_name.startswith("_"):
        raise HTTPException(status_code=400, detail="フォルダ名は先頭アンダースコア不可です（_index 衝突防止）")
    if ".." in folder_name:
        raise HTTPException(status_code=400, detail="フォルダ名に .. は使えません")

    parent = Path(raw_parent).expanduser().resolve()
    if not parent.is_dir():
        raise HTTPException(status_code=400, detail=f"親フォルダが見つかりません: {raw_parent}")

    target = parent / folder_name
    if target.exists():
        raise HTTPException(status_code=409, detail=f"同名フォルダが既に存在します: {folder_name}")

    # 入れ子チェック: parent が既登録事件フォルダの配下（または一致）でないこと
    with _config_lock:
        for c in CONFIG["cases"].values():
            case_path = c["path"].resolve()
            try:
                parent.relative_to(case_path)
                raise HTTPException(
                    status_code=409,
                    detail=f"親フォルダが既登録事件「{c['name']}」の配下です（入れ子登録禁止）",
                )
            except ValueError:
                pass
            # parent が case_path 自身でもエラー（同上）
            if parent == case_path:
                raise HTTPException(
                    status_code=409,
                    detail=f"親フォルダが既登録事件「{c['name']}」です（入れ子登録禁止）",
                )

        target.mkdir(parents=False)
        name = (payload.get("name") or "").strip() or folder_name
        reindex = bool(payload.get("reindex", False))
        upload_subdir = (payload.get("upload_subdir") or "").strip()
        return _register_case(target, name, reindex, upload_subdir)


@app.delete("/api/cases/{case_id}")
def api_delete_case(case_id: str):
    """事件の登録を解除する（ファイル・索引・注釈は削除しない）。"""
    with _config_lock:
        if case_id not in CONFIG["cases"]:
            raise HTTPException(status_code=404, detail=f"unknown case: {case_id}")
        proc = _reindex_procs.get(case_id)
        if proc is not None and proc.poll() is None:
            raise HTTPException(status_code=409, detail="索引作成中は登録解除できません")
        CONFIG["cases"].pop(case_id)
        save_cases_config()
    _listing_cache.pop(case_id, None)
    _reindex_procs.pop(case_id, None)
    return {"ok": True, "note": "登録解除のみ。ファイルは削除されていません"}


_pick_folder_lock = threading.Lock()


@app.post("/api/pick-folder")
def api_pick_folder():
    """macOS の Finder フォルダ選択ダイアログを開いてパスを返す。"""
    import subprocess

    if not _pick_folder_lock.acquire(blocking=False):
        raise HTTPException(status_code=409, detail="folder dialog already open")
    try:
        proc = subprocess.run(
            [
                "osascript",
                "-e", 'tell application "System Events" to activate',
                "-e", 'POSIX path of (choose folder with prompt "事件フォルダを選択してください")',
            ],
            capture_output=True, text=True, timeout=300,
        )
    except subprocess.TimeoutExpired:
        return {"cancelled": True}
    finally:
        _pick_folder_lock.release()
    if proc.returncode != 0:
        return {"cancelled": True}
    return {"path": proc.stdout.strip()}


@app.get("/api/cases/{case_id}/reindex/log")
def api_reindex_log(case_id: str):
    """索引生成ログの末尾50行を返す。"""
    get_case(CONFIG, case_id)
    log_path = CACHE_ROOT / case_id / "reindex.log"
    if not log_path.exists():
        return {"log": ""}
    try:
        lines = log_path.read_text(encoding="utf-8", errors="replace").splitlines()
        return {"log": "\n".join(lines[-50:])}
    except OSError:
        return {"log": ""}


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


def _sanitize_search_hit(hit: dict) -> dict:
    """検索ヒットの title を弁護革命ID除去後の表示値に正規化する。"""
    # search_pages は file_name を返さないため、title から直接サニタイズは難しい。
    # ここでは title 文字列自体に BENGOKAKUMEI_ID_RE パターンが残っている場合を処理する。
    # （build_listing 側でサニタイズ済みの DB があれば不要だが保険として適用）
    title = hit.get("title", "")
    sanitized = ei.BENGOKAKUMEI_ID_RE.sub("", title)
    if sanitized and sanitized != title:
        hit = dict(hit)
        hit["title"] = sanitized.strip()
    return hit


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
    hits = [_sanitize_search_hit(h) for h in hits]
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


# --------------------------------------------------------------------------
# 注釈（D8 座標・D9 規律）
# --------------------------------------------------------------------------

@app.get("/api/cases/{case_id}/annotations/{sha256}")
def api_get_annotations(case_id: str, sha256: str):
    case = get_case(CONFIG, case_id)
    user = current_user()
    data = read_annotations(case, sha256, user)
    return {**data, "user": user, "user_display": display_name(CONFIG, user)}


@app.put("/api/cases/{case_id}/annotations/{sha256}")
def api_put_annotations(case_id: str, sha256: str, payload: dict = Body(...)):
    """自ユーザーの注釈のみ保存（楽観ロック409・アトミック書き込み／D9）。"""
    case = get_case(CONFIG, case_id)
    user = current_user()
    validate_sha(sha256)
    fpath = annotation_file(case, user, sha256)
    stored_at = ""
    if fpath.exists():
        try:
            stored_at = json.loads(fpath.read_text(encoding="utf-8")).get("updated_at", "")
        except (OSError, json.JSONDecodeError):
            stored_at = ""
    base = payload.get("base_updated_at", "")
    if stored_at and base and stored_at > base:
        raise HTTPException(status_code=409, detail="annotations changed; reload required")
    updated = now_iso()
    anns = payload.get("annotations", [])
    if not isinstance(anns, list):
        raise HTTPException(status_code=400, detail="annotations must be a list")
    # 他ユーザー用の表示フィールドは保存しない。
    clean = [{k: v for k, v in a.items() if not k.startswith("_")} for a in anns]
    atomic_write_json(fpath, {"updated_at": updated, "annotations": clean})
    _listing_cache.pop(case_id, None)
    return {"updated_at": updated, "count": len(clean)}


# --------------------------------------------------------------------------
# メタ（メモ・符号上書き・丁数オフセット・回転状態／D9）
# --------------------------------------------------------------------------

@app.put("/api/cases/{case_id}/meta/{sha256}")
def api_put_meta(case_id: str, sha256: str, payload: dict = Body(...)):
    case = get_case(CONFIG, case_id)
    user = current_user()
    validate_sha(sha256)
    if not USER_RE.match(user):
        raise HTTPException(status_code=400, detail="invalid user")
    mfile = case["path"] / INDEX_DIR_NAME / VIEWER_META_DIR / f"{user}.json"
    data = {}
    if mfile.exists():
        try:
            data = json.loads(mfile.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            data = {}
    entry = data.get(sha256, {})
    ts = now_iso()
    for field in META_FIELDS:
        if field in payload:
            entry[field] = payload[field]
            entry[f"{field}_at"] = ts
    data[sha256] = entry
    atomic_write_json(mfile, data)
    _listing_cache.pop(case_id, None)
    return {"ok": True, "updated_at": ts}


# --------------------------------------------------------------------------
# 注釈込みPDF書き出し（D8 座標をそのまま焼き込み）
# --------------------------------------------------------------------------

def parse_pages_param(pages: str, total: int) -> set[int] | None:
    if not pages:
        return None
    out: set[int] = set()
    for part in pages.split(","):
        part = part.strip()
        if "-" in part:
            a, _, b = part.partition("-")
            try:
                lo, hi = int(a), int(b)
            except ValueError:
                continue
            out.update(range(max(1, lo), min(total, hi) + 1))
        elif part:
            try:
                out.add(int(part))
            except ValueError:
                continue
    return out or None


@app.get("/api/cases/{case_id}/export/{sha256}")
def api_export(case_id: str, sha256: str, pages: str = ""):
    """注釈をPDF座標のまま焼き込んだPDFを返す（pypdf＋reportlab）。"""
    case = get_case(CONFIG, case_id)
    path = resolve_sha(case, sha256)
    if path.suffix.lower() != ".pdf":
        raise HTTPException(status_code=400, detail="not a pdf")
    anns = read_annotations(case, sha256, current_user())["annotations"]
    data = bake_annotations(path, anns, pages)
    from urllib.parse import quote

    fname = quote(path.stem + "_注釈付き.pdf")
    return Response(
        content=data,
        media_type="application/pdf",
        headers={"Content-Disposition": f"attachment; filename=export.pdf; filename*=UTF-8''{fname}"},
    )


def bake_annotations(pdf_path: Path, anns: list[dict], pages_param: str) -> bytes:
    import io

    from pypdf import PdfReader, PdfWriter
    from reportlab.pdfgen import canvas as rl_canvas

    reader = PdfReader(str(pdf_path))
    total = len(reader.pages)
    keep = parse_pages_param(pages_param, total)
    by_page: dict[int, list[dict]] = {}
    for a in anns:
        by_page.setdefault(int(a.get("page", 0)), []).append(a)

    writer = PdfWriter()
    for idx, page in enumerate(reader.pages, start=1):
        if keep is not None and idx not in keep:
            continue
        wp = writer.add_page(page)  # 先にwriterへ付けてからmerge（pypdf推奨）
        page_anns = by_page.get(idx, [])
        if page_anns:
            mb = wp.mediabox
            w = float(mb.width)
            h = float(mb.height)
            buf = io.BytesIO()
            c = rl_canvas.Canvas(buf, pagesize=(w, h))
            for a in page_anns:
                draw_annotation(c, a)
            c.save()
            buf.seek(0)
            overlay = PdfReader(buf).pages[0]
            wp.merge_page(overlay)

    out = io.BytesIO()
    writer.write(out)
    return out.getvalue()


def _rgb(color: str):
    color = (color or "#cc0000").lstrip("#")
    if len(color) != 6:
        color = "cc0000"
    return tuple(int(color[i:i + 2], 16) / 255 for i in (0, 2, 4))


def draw_annotation(c, a: dict) -> None:
    """注釈をPDFユーザー空間（左下原点）にそのまま描画する（D8）。"""
    t = a.get("type")
    r, g, b = _rgb(a.get("color"))
    c.setStrokeColorRGB(r, g, b)
    c.setFillColorRGB(r, g, b)
    c.setLineWidth(float(a.get("width", 2)))
    if t == "rect":
        x0, y0, x1, y1 = a["rect"]
        c.setFillAlpha(0.0)
        c.rect(min(x0, x1), min(y0, y1), abs(x1 - x0), abs(y1 - y0), stroke=1, fill=0)
    elif t == "line":
        (x0, y0), (x1, y1) = a["points"][0], a["points"][-1]
        c.line(x0, y0, x1, y1)
    elif t == "pen":
        pts = a.get("points", [])
        if len(pts) >= 2:
            p = c.beginPath()
            p.moveTo(pts[0][0], pts[0][1])
            for x, y in pts[1:]:
                p.lineTo(x, y)
            c.drawPath(p, stroke=1, fill=0)
    elif t in ("text", "comment", "note"):
        x, y = annotation_anchor(a)
        if t == "note":
            c.setFont("Helvetica", 14)
            c.drawString(x, y, "★")
        txt = a.get("text", "")
        if txt:
            try:
                from reportlab.pdfbase import pdfmetrics
                from reportlab.pdfbase.cidfonts import UnicodeCIDFont
                pdfmetrics.registerFont(UnicodeCIDFont("HeiseiKakuGo-W5"))
                c.setFont("HeiseiKakuGo-W5", 10)
            except Exception:
                c.setFont("Helvetica", 10)
            c.drawString(x + (14 if t == "note" else 0), y, txt)


def annotation_anchor(a: dict):
    if "rect" in a:
        x0, y0, x1, y1 = a["rect"]
        return min(x0, x1), max(y0, y1)
    if "point" in a:
        return a["point"][0], a["point"][1]
    pts = a.get("points")
    if pts:
        return pts[0][0], pts[0][1]
    return 36, 36


# --------------------------------------------------------------------------
# 孤児注釈（documents に無い sha256）と紐付け直し
# --------------------------------------------------------------------------

@app.get("/api/cases/{case_id}/orphan-annotations")
def api_orphans(case_id: str):
    case = get_case(CONFIG, case_id)
    user = current_user()
    docs, _ = get_listing(case)
    known = {d["sha256"] for d in docs}
    orphans = []
    adir = annotations_dir(case, user)
    if adir.is_dir():
        for fpath in adir.iterdir():
            if not SHA_RE.match(fpath.name):
                continue
            sha = fpath.stem
            if sha in known:
                continue
            try:
                payload = json.loads(fpath.read_text(encoding="utf-8"))
            except (OSError, json.JSONDecodeError):
                continue
            if payload.get("annotations"):
                orphans.append({"sha256": sha, "count": len(payload["annotations"])})
    return {"orphans": orphans, "user": user}


@app.post("/api/cases/{case_id}/annotations/{sha256}/relink")
def api_relink(case_id: str, sha256: str, payload: dict = Body(...)):
    """孤児注釈を別文書(sha)へ紐付け直す（自ユーザー分のみ）。"""
    case = get_case(CONFIG, case_id)
    user = current_user()
    validate_sha(sha256)
    target = validate_sha(payload.get("target_sha", ""))
    docs, _ = get_listing(case)
    if target not in {d["sha256"] for d in docs}:
        raise HTTPException(status_code=404, detail="target document not found")
    src = annotation_file(case, user, sha256)
    if not src.exists():
        raise HTTPException(status_code=404, detail="source annotations not found")
    src_data = json.loads(src.read_text(encoding="utf-8"))
    dst = annotation_file(case, user, target)
    dst_data = {"updated_at": "", "annotations": []}
    if dst.exists():
        try:
            dst_data = json.loads(dst.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            pass
    merged = (dst_data.get("annotations", []) or []) + (src_data.get("annotations", []) or [])
    atomic_write_json(dst, {"updated_at": now_iso(), "annotations": merged})
    src.unlink()
    _listing_cache.pop(case_id, None)
    return {"ok": True, "moved": len(src_data.get("annotations", [])), "target": target}


# --------------------------------------------------------------------------
# ファイルアップロード（D&D / ④b）
# --------------------------------------------------------------------------

UPLOAD_MAX_BYTES = 500 * 1024 * 1024  # 500 MB
IMPORT_LOG_NAME = "import_log.json"
_UNSAFE_FILENAME_RE = re.compile(r'[/\\:\x00]|^\.')


def _upload_dir(case: dict) -> Path:
    """D&D の保存先ディレクトリを返す。upload_subdir が設定されていればその配下。"""
    subdir = case.get("upload_subdir", "").strip()
    if subdir:
        return (case["path"] / subdir).resolve()
    return case["path"].resolve()


def _append_import_log(case: dict, entry: dict) -> None:
    """_index/import_log.json にアトミック追記する。"""
    log_path = case["path"] / INDEX_DIR_NAME / IMPORT_LOG_NAME
    log_path.parent.mkdir(parents=True, exist_ok=True)
    existing: list[dict] = []
    if log_path.exists():
        try:
            existing = json.loads(log_path.read_text(encoding="utf-8"))
            if not isinstance(existing, list):
                existing = []
        except (OSError, json.JSONDecodeError):
            existing = []
    existing.append(entry)
    atomic_write_json(log_path, existing)


@app.post("/api/cases/{case_id}/upload")
async def api_upload(case_id: str, file: UploadFile = File(...)):
    """事件フォルダへファイルをアップロードする（D&D ドロップ先）。

    - 拡張子ホワイトリスト: OPENABLE_EXTS
    - sha256 重複は拒否
    - 同名・別内容は 409
    - インポートログを _index/import_log.json に追記
    """
    case = get_case(CONFIG, case_id)

    # ファイル名サニタイズ
    orig_name = file.filename or ""
    safe_name = Path(orig_name).name  # basename のみ
    if not safe_name or _UNSAFE_FILENAME_RE.search(safe_name) or ".." in safe_name:
        raise HTTPException(status_code=400, detail="ファイル名が不正です")

    ext = Path(safe_name).suffix.lower()
    if ext not in OPENABLE_EXTS:
        raise HTTPException(
            status_code=400,
            detail=f"許可されていない拡張子です: {ext}（許可: {', '.join(sorted(OPENABLE_EXTS))}）",
        )

    # 保存先ディレクトリの検証
    dest_dir = _upload_dir(case)
    if not dest_dir.exists():
        raise HTTPException(status_code=400, detail=f"保存先フォルダが見つかりません: {dest_dir}")

    dest = (dest_dir / safe_name).resolve()
    base = case["path"].resolve()
    index_dir = (base / INDEX_DIR_NAME).resolve()

    # パストラバーサル検証
    if base not in dest.parents and dest != base:
        raise HTTPException(status_code=403, detail="保存先が事件フォルダ外です")
    # _index 配下への書き込み禁止
    try:
        dest.relative_to(index_dir)
        raise HTTPException(status_code=403, detail="_index フォルダへの書き込みは禁止です")
    except ValueError:
        pass

    # データ読み込み（サイズ上限）
    chunks: list[bytes] = []
    total = 0
    while True:
        chunk = await file.read(256 * 1024)
        if not chunk:
            break
        total += len(chunk)
        if total > UPLOAD_MAX_BYTES:
            raise HTTPException(status_code=413, detail=f"ファイルサイズが上限（500MB）を超えています")
        chunks.append(chunk)
    data = b"".join(chunks)

    incoming_sha = hashlib.sha256(data).hexdigest()

    # sha256 重複チェック（既存文書と同一内容）
    docs, _ = get_listing(case)
    for d in docs:
        if d["sha256"] == incoming_sha:
            raise HTTPException(
                status_code=409,
                detail=f"既に登録済みです（同一内容: {d.get('file_name', d['sha256'][:12])}）",
            )

    # 同名・別内容チェック
    if dest.exists():
        existing_sha = ei.sha256_file(dest)
        if existing_sha != incoming_sha:
            raise HTTPException(
                status_code=409,
                detail="同名の別内容ファイルがあります（差し替え証拠の可能性）。別名で保存するか中止してください。",
            )
        # 同名・同内容はここには来ない（sha重複で弾かれている）

    # .tmp → os.replace（D9）
    tmp = dest.with_name(dest.name + ".tmp")
    tmp.write_bytes(data)
    os.replace(tmp, dest)

    # インポートログ追記
    _append_import_log(case, {
        "time": now_iso(),
        "user": current_user(),
        "file_name": safe_name,
        "sha256": incoming_sha,
        "size": total,
    })

    # キャッシュ無効化（TTL 5 秒の間に一覧に出ず再ドロップ→事故を防ぐ）
    _listing_cache.pop(case_id, None)

    return {
        "ok": True,
        "file_name": safe_name,
        "sha256": incoming_sha,
        "size": total,
        "note": "Drive 同期はバックグラウンドで行われます。",
    }


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
