#!/usr/bin/env python3
"""Build and query a local evidence index for criminal case records.

The source PDFs remain the record originals. This script creates secondary
search artifacts under an `_index/export/` folder:

- evidence_index.sqlite: SQLite database with FTS5 page search.
- manifest.csv: PDF inventory and extraction metadata.
- timeline.csv / timeline.md: date candidates that require human review.
- llm_usage_log.csv: audit log header for later LLM/NotebookLM use.

To keep Google Drive sync safer, the SQLite database is built in a temporary
local workspace and copied to the evidence folder only after the build
finishes.
"""

from __future__ import annotations

import argparse
import csv
import datetime as dt
import hashlib
import json
import os
import re
import shutil
import sqlite3
import subprocess
import sys
import tempfile
import unicodedata
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable


def normalize_for_search(text: str) -> str:
    """Normalize text for full-text search (D1).

    Applies NFKC normalization and lowercasing so that full-width / half-width
    and case variations match. This MUST be applied identically to both the
    text inserted into the FTS index and to the search query, otherwise the
    trigram tokens will not line up.
    """
    return unicodedata.normalize("NFKC", text).lower()


REPO_ROOT = Path(__file__).resolve().parents[1]
DEFAULT_PATH_FILE = REPO_ROOT / ".codex" / "shell_path_utf8.txt"
ALT_PATH_FILE = REPO_ROOT / ".cursor" / "shell_path_utf8.txt"
VISION_OCR_SCRIPT = REPO_ROOT / "scripts" / "evidence_vision_ocr.swift"

INDEX_DIR_NAME = "_index"
EXPORT_DIR_NAME = "export"
LOCK_NAME = ".lock"

STATUS_OCR_ONLY = "OCRのみ"
STATUS_NEEDS_REVIEW = "要再確認"
STATUS_VISUAL_CONFIRMED = "原PDF目視確認済み"

DATE_RE = re.compile(
    r"(?P<yyyy>20\d{2}|19\d{2})[./年-]\s*(?P<m>\d{1,2})[./月-]\s*(?P<d>\d{1,2})日?"
    r"|令和\s*(?P<reiwa>\d{1,2})\s*年\s*(?P<rm>\d{1,2})\s*月\s*(?P<rd>\d{1,2})\s*日"
    r"|R\s*(?P<rnum>\d{1,2})[./-](?P<rnm>\d{1,2})[./-](?P<rnd>\d{1,2})"
    r"|(?<!\d)(?P<compact>20\d{6}|19\d{6}|\d{6})(?!\d)"
)

EVIDENCE_FILE_RE = re.compile(
    r"^(?P<evidence_no>[甲乙弁][A-Za-zＡ-Ｚａ-ｚ]?[0-9０-９]+(?:[-_－―の][0-9０-９]+)?)"
    r"\s*(?P<rest>.*)$"
)

# 弁護革命が付加する末尾ID（英字・数字混在必須・10〜16桁）を除去するパターン。
# 先頭ドット（拡張子区切り）を含む形で末尾にマッチ。
BENGOKAKUMEI_ID_RE = re.compile(
    r"\.(?=[a-z0-9]*[0-9])(?=[a-z0-9]*[a-z])[a-z0-9]{10,16}(?=$|[@＠])"
)


EVIDENCE_KIND_ORDER = {"甲": 0, "乙": 1, "弁": 2}
_FW_DIGITS = str.maketrans("０１２３４５６７８９", "0123456789")


def evidence_sort_key(evidence_no: str) -> tuple:
    """Natural-order sort key for evidence numbers (D4).

    Returns (kind_order, alpha, main_number, branch_number, raw) so that
    甲1, 甲2, 甲10 sort numerically rather than lexically, and 甲 < 乙 < 弁.
    Documents without a recognizable evidence number sort last.
    """
    if not evidence_no:
        return (9, "", 0, 0, "")
    s = evidence_no.translate(_FW_DIGITS)
    kind = s[0] if s else ""
    kind_order = EVIDENCE_KIND_ORDER.get(kind, 8)
    rest = s[1:]
    m = re.match(r"([A-Za-zＡ-Ｚａ-ｚ]?)(\d+)(?:[-_－―の](\d+))?", rest)
    if not m:
        return (kind_order, rest, 0, 0, s)
    alpha = m.group(1) or ""
    main = int(m.group(2)) if m.group(2) else 0
    branch = int(m.group(3)) if m.group(3) else 0
    return (kind_order, alpha, main, branch, s)


@dataclass(frozen=True)
class PageText:
    page_no: int
    text: str
    method: str
    ocr_confidence: float | None
    error: str


@dataclass(frozen=True)
class PdfMetadata:
    path: Path
    rel_path: str
    file_name: str
    evidence_no: str
    title: str
    document_date: str
    person_or_source: str
    sha256: str
    file_size: int
    modified_at: str


def now_iso() -> str:
    return dt.datetime.now(dt.timezone.utc).astimezone().isoformat(timespec="seconds")


def normalize_date(year: int, month: int, day: int) -> str:
    try:
        return dt.date(year, month, day).isoformat()
    except ValueError:
        return ""


def compact_date_to_iso(value: str) -> str:
    if len(value) == 8:
        return normalize_date(int(value[:4]), int(value[4:6]), int(value[6:8]))
    if len(value) == 6:
        return normalize_date(2000 + int(value[:2]), int(value[2:4]), int(value[4:6]))
    return ""


def normalize_date_match(match: re.Match[str]) -> str:
    gd = match.groupdict()
    if gd.get("yyyy"):
        return normalize_date(int(gd["yyyy"]), int(gd["m"]), int(gd["d"]))
    if gd.get("reiwa"):
        return normalize_date(2018 + int(gd["reiwa"]), int(gd["rm"]), int(gd["rd"]))
    if gd.get("rnum"):
        return normalize_date(2018 + int(gd["rnum"]), int(gd["rnm"]), int(gd["rnd"]))
    if gd.get("compact"):
        return compact_date_to_iso(gd["compact"])
    return ""


def read_path_file(path_file: Path) -> Path:
    if not path_file.exists() and path_file == DEFAULT_PATH_FILE and ALT_PATH_FILE.exists():
        path_file = ALT_PATH_FILE
    try:
        lines = [line.strip() for line in path_file.read_text(encoding="utf-8").splitlines()]
    except OSError as exc:
        raise SystemExit(f"cannot read path file: {path_file}: {exc}") from exc
    paths = [Path(line).expanduser() for line in lines if line.strip()]
    if not paths:
        raise SystemExit(f"path file is empty: {path_file}")
    return paths[0]


def resolve_evidence_dir(args: argparse.Namespace) -> Path:
    if args.evidence_dir:
        return Path(args.evidence_dir).expanduser().resolve()
    return read_path_file(Path(args.path_file).resolve()).resolve()


def sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def parse_name_only(stem: str) -> tuple[str, str, str, str]:
    """ファイル名（拡張子なし）から符号・標目・作成者・日付を解析する。

    ファイルをハッシュ・stat しないので PDF・メディア双方に使える。
    返り値: (evidence_no, title, document_date, person_or_source)

    処理順: 弁護革命ID除去 → 符号 → @分離（全角・半角の最後の出現） → 日付抽出
    """
    # Step 1: 弁護革命末尾IDを除去（英字・数字混在必須・10〜16桁）
    cleaned = BENGOKAKUMEI_ID_RE.sub("", stem)
    if not cleaned:
        cleaned = stem  # 除去後に空になる場合は元に戻す

    evidence_no = ""
    title_part = cleaned

    m = EVIDENCE_FILE_RE.match(cleaned)
    if m:
        evidence_no = m.group("evidence_no")
        title_part = m.group("rest").strip()

    # Step 2: @分離（全角＠と半角@の両方、最後の出現で rsplit）
    person_or_source = ""
    last_full = title_part.rfind("＠")  # ＠ = U+FF20
    last_half = title_part.rfind("@")
    last_at = max(last_full, last_half)
    if last_at >= 0:
        person_or_source = title_part[last_at + 1 :].strip()
        title_part = title_part[:last_at].strip()

    # Step 3: 日付抽出（ID除去後に行うことで誤日付汚染を防ぐ）
    document_date = ""
    date_hits = list(DATE_RE.finditer(title_part))
    if date_hits:
        last = date_hits[-1]
        document_date = normalize_date_match(last)
        title_part = (title_part[: last.start()] + title_part[last.end() :]).strip()

    title = re.sub(r"\s+", " ", title_part).strip().rstrip(".").strip() or stem
    return evidence_no, title, document_date, person_or_source


def parse_pdf_name(path: Path, evidence_dir: Path) -> PdfMetadata:
    stat = path.stat()
    rel_path = path.relative_to(evidence_dir).as_posix()
    evidence_no, title, document_date, person_or_source = parse_name_only(path.stem)
    return PdfMetadata(
        path=path,
        rel_path=rel_path,
        file_name=path.name,
        evidence_no=evidence_no,
        title=title,
        document_date=document_date,
        person_or_source=person_or_source,
        sha256=sha256_file(path),
        file_size=stat.st_size,
        modified_at=dt.datetime.fromtimestamp(stat.st_mtime, dt.timezone.utc)
        .astimezone()
        .isoformat(timespec="seconds"),
    )


# PDFs excluded from indexing even when found by rglob.
# - 分冊マスターPDF: navigation bundles that duplicate all individual files
# - 重複？…: explicitly marked duplicate bundle
_EXCLUDED_PDF_NAMES: frozenset[str] = frozenset(
    [
        "第４分冊.pdf",
        "第５分冊.pdf",
        "重複？（１２３〜）.pdf",
    ]
)


def iter_pdfs(evidence_dir: Path) -> Iterable[Path]:
    for path in sorted(evidence_dir.rglob("*.pdf")):
        # Exclude _index directories at any depth in the path
        if INDEX_DIR_NAME in path.parts:
            continue
        if path.name in _EXCLUDED_PDF_NAMES:
            continue
        if path.is_file():
            yield path


def extract_with_fitz(
    path: Path,
    enable_ocr: bool,
    ocr_engine: str,
    force_ocr: bool,
) -> tuple[list[PageText], str]:
    try:
        import fitz  # type: ignore
    except ImportError as exc:
        raise RuntimeError("fitz-not-installed") from exc

    pages: list[PageText] = []
    doc = fitz.open(path)
    try:
        for index, page in enumerate(doc, start=1):
            text = page.get_text() or ""
            method = "fitz"
            confidence: float | None = None
            error = ""
            if enable_ocr and (force_ocr or not text.strip()):
                ocr_text, confidence, method, error = ocr_page(page, ocr_engine)
                if ocr_text.strip():
                    text = ocr_text
                elif not text.strip():
                    method = "fitz-empty"
            elif not text.strip():
                error = "empty-text"
            pages.append(
                PageText(
                    page_no=index,
                    text=text.strip(),
                    method=method,
                    ocr_confidence=confidence,
                    error=error,
                )
            )
    finally:
        doc.close()
    return pages, "fitz"


def ocr_page(page: object, ocr_engine: str) -> tuple[str, float | None, str, str]:
    if ocr_engine in {"auto", "vision"}:
        text, confidence, error = ocr_page_with_vision(page)
        if text.strip() or ocr_engine == "vision":
            return text, confidence, "apple-vision", error
    if ocr_engine in {"auto", "tesseract"}:
        text, confidence, error = ocr_page_with_tesseract(page)
        return text, confidence, "tesseract", error
    return "", None, "ocr-disabled", f"unknown-ocr-engine:{ocr_engine}"


def ocr_page_with_vision(page: object) -> tuple[str, float | None, str]:
    if not VISION_OCR_SCRIPT.exists():
        return "", None, "vision-helper-not-found"
    if shutil.which("swift") is None:
        return "", None, "swift-not-found"

    with tempfile.TemporaryDirectory(prefix="evidence-vision-") as tmp:
        image_path = Path(tmp) / "page.png"
        try:
            pix = page.get_pixmap(dpi=300)  # type: ignore[attr-defined]
            pix.save(str(image_path))
        except Exception as exc:  # noqa: BLE001
            return "", None, f"vision-render-error:{exc.__class__.__name__}"

        try:
            proc = subprocess.run(
                ["swift", str(VISION_OCR_SCRIPT), str(image_path)],
                check=False,
                capture_output=True,
                text=True,
                timeout=120,
            )
        except subprocess.TimeoutExpired:
            return "", None, "vision-timeout"
        except OSError as exc:
            return "", None, f"vision-run-error:{exc.__class__.__name__}"

        if proc.returncode != 0:
            msg = (proc.stderr or proc.stdout or "").strip()[:200]
            return "", None, f"vision-error:{msg}"

        lines = (proc.stdout or "").strip().splitlines()
        if not lines:
            return "", None, "vision-empty-output"
        try:
            payload = json.loads(lines[-1])
        except json.JSONDecodeError:
            return "", None, "vision-json-error"
        text = str(payload.get("text") or "").strip()
        confidence_raw = payload.get("confidence")
        confidence = round(float(confidence_raw) * 100, 2) if confidence_raw is not None else None
        error = str(payload.get("error") or "")
        return text, confidence, error


def ocr_page_with_tesseract(page: object) -> tuple[str, float | None, str]:
    try:
        import io

        import pytesseract  # type: ignore
        from PIL import Image  # type: ignore
    except ImportError:
        return "", None, "ocr-dependencies-not-installed"

    try:
        langs = set(pytesseract.get_languages(config=""))
    except Exception as exc:  # noqa: BLE001
        return "", None, f"tesseract-language-error:{exc.__class__.__name__}"

    if "jpn" not in langs:
        return "", None, "tesseract-jpn-not-installed"

    lang_parts = ["jpn"]
    if "jpn_vert" in langs:
        lang_parts.append("jpn_vert")
    if "eng" in langs:
        lang_parts.append("eng")
    lang = "+".join(lang_parts)

    try:
        pix = page.get_pixmap(dpi=300)  # type: ignore[attr-defined]
        image = Image.open(io.BytesIO(pix.tobytes("png")))
        text = pytesseract.image_to_string(image, lang=lang) or ""
        confidence = mean_tesseract_confidence(pytesseract, image, lang)
        return text.strip(), confidence, ""
    except Exception as exc:  # noqa: BLE001
        return "", None, f"ocr-runtime-error:{exc.__class__.__name__}"


def mean_tesseract_confidence(pytesseract: object, image: object, lang: str) -> float | None:
    try:
        data = pytesseract.image_to_data(  # type: ignore[attr-defined]
            image,
            lang=lang,
            output_type=pytesseract.Output.DICT,  # type: ignore[attr-defined]
        )
        values = [float(v) for v in data.get("conf", []) if str(v) not in {"-1", ""}]
        if values:
            return round(sum(values) / len(values), 2)
    except Exception:
        return None
    return None


def extract_with_pdftotext(path: Path) -> tuple[list[PageText], str]:
    proc = subprocess.run(
        ["pdftotext", "-layout", "-enc", "UTF-8", str(path), "-"],
        check=False,
        capture_output=True,
        text=True,
        timeout=120,
    )
    if proc.returncode != 0:
        msg = (proc.stderr or "").strip()[:200]
        return [PageText(1, "", "pdftotext", None, f"pdftotext-error:{msg}")], "pdftotext"
    raw_pages = proc.stdout.split("\f")
    pages = []
    for i, text in enumerate(raw_pages, start=1):
        if i == len(raw_pages) and not text.strip():
            continue
        pages.append(PageText(i, text.strip(), "pdftotext", None, "" if text.strip() else "empty-text"))
    if not pages:
        pages.append(PageText(1, "", "pdftotext", None, "empty-text"))
    return pages, "pdftotext"


def extract_pdf_pages(
    path: Path,
    enable_ocr: bool,
    ocr_engine: str,
    force_ocr: bool,
) -> tuple[list[PageText], str]:
    try:
        return extract_with_fitz(path, enable_ocr, ocr_engine, force_ocr)
    except Exception:
        return extract_with_pdftotext(path)


def create_schema(conn: sqlite3.Connection) -> None:
    conn.executescript(
        """
        PRAGMA journal_mode = WAL;

        CREATE TABLE documents (
            id INTEGER PRIMARY KEY,
            rel_path TEXT NOT NULL UNIQUE,
            abs_path TEXT NOT NULL,
            file_name TEXT NOT NULL,
            evidence_no TEXT NOT NULL,
            title TEXT NOT NULL,
            document_date TEXT NOT NULL,
            person_or_source TEXT NOT NULL,
            sha256 TEXT NOT NULL,
            file_size INTEGER NOT NULL,
            modified_at TEXT NOT NULL,
            acquired_at TEXT NOT NULL,
            page_count INTEGER NOT NULL,
            extract_method TEXT NOT NULL,
            extract_status TEXT NOT NULL,
            created_at TEXT NOT NULL
        );

        CREATE TABLE pages (
            id INTEGER PRIMARY KEY,
            document_id INTEGER NOT NULL,
            page_no INTEGER NOT NULL,
            material_page_label TEXT NOT NULL DEFAULT '',
            text TEXT NOT NULL,
            text_norm TEXT NOT NULL DEFAULT '',
            text_hash TEXT NOT NULL,
            extract_method TEXT NOT NULL,
            ocr_confidence REAL,
            review_status TEXT NOT NULL,
            error TEXT NOT NULL,
            FOREIGN KEY(document_id) REFERENCES documents(id)
        );

        CREATE VIRTUAL TABLE pages_fts USING fts5(
            text,
            evidence_no UNINDEXED,
            title UNINDEXED,
            rel_path UNINDEXED,
            page_id UNINDEXED,
            tokenize='trigram'
        );

        CREATE TABLE timeline_events (
            id INTEGER PRIMARY KEY,
            event_date TEXT NOT NULL,
            document_date TEXT NOT NULL,
            evidence_no TEXT NOT NULL,
            title TEXT NOT NULL,
            rel_path TEXT NOT NULL,
            pdf_page_no INTEGER NOT NULL,
            material_page_label TEXT NOT NULL DEFAULT '',
            quote_text TEXT NOT NULL,
            quote_start INTEGER,
            quote_end INTEGER,
            review_status TEXT NOT NULL,
            reviewer TEXT NOT NULL DEFAULT '',
            reviewed_at TEXT NOT NULL DEFAULT '',
            confidence_reason TEXT NOT NULL
        );

        CREATE TABLE build_info (
            key TEXT PRIMARY KEY,
            value TEXT NOT NULL
        );
        """
    )


def page_review_status(page: PageText) -> str:
    if page.method in {"apple-vision", "tesseract"} and page.text:
        return STATUS_OCR_ONLY
    return STATUS_NEEDS_REVIEW


def text_hash(text: str) -> str:
    return hashlib.sha256(text.encode("utf-8", errors="replace")).hexdigest()


def add_timeline_candidates(
    conn: sqlite3.Connection,
    meta: PdfMetadata,
    pages: list[PageText],
    max_events_per_page: int,
) -> None:
    for page in pages:
        seen_dates: set[tuple[str, int]] = set()
        matches = list(DATE_RE.finditer(page.text))
        for match in matches[:max_events_per_page]:
            normalized = normalize_date_match(match)
            if not normalized:
                continue
            key = (normalized, page.page_no)
            if key in seen_dates:
                continue
            seen_dates.add(key)
            quote = excerpt(page.text, match.start(), match.end())
            conn.execute(
                """
                INSERT INTO timeline_events (
                    event_date, document_date, evidence_no, title, rel_path,
                    pdf_page_no, quote_text, quote_start, quote_end,
                    review_status, confidence_reason
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    normalized,
                    meta.document_date,
                    meta.evidence_no,
                    meta.title,
                    meta.rel_path,
                    page.page_no,
                    quote,
                    match.start(),
                    match.end(),
                    STATUS_NEEDS_REVIEW,
                    "本文の日付表現から自動抽出。原PDF目視確認が必要。",
                ),
            )

    if meta.document_date:
        conn.execute(
            """
            INSERT INTO timeline_events (
                event_date, document_date, evidence_no, title, rel_path,
                pdf_page_no, quote_text, review_status, confidence_reason
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                "",
                meta.document_date,
                meta.evidence_no,
                meta.title,
                meta.rel_path,
                0,
                "ファイル名から文書作成日候補を抽出。",
                STATUS_NEEDS_REVIEW,
                "文書作成日候補。出来事日ではない可能性がある。",
            ),
        )


def excerpt(text: str, start: int, end: int, width: int = 90) -> str:
    left = max(0, start - width)
    right = min(len(text), end + width)
    snippet = text[left:right].replace("\r", " ").replace("\n", " ")
    return re.sub(r"\s+", " ", snippet).strip()


def build_database(
    evidence_dir: Path,
    db_path: Path,
    enable_ocr: bool,
    ocr_engine: str,
    force_ocr: bool,
    max_events_per_page: int,
    limit: int | None,
) -> tuple[list[PdfMetadata], int]:
    conn = sqlite3.connect(db_path)
    try:
        create_schema(conn)
        created_at = now_iso()
        conn.execute("INSERT INTO build_info (key, value) VALUES (?, ?)", ("created_at", created_at))
        conn.execute("INSERT INTO build_info (key, value) VALUES (?, ?)", ("evidence_dir", str(evidence_dir)))
        conn.execute("INSERT INTO build_info (key, value) VALUES (?, ?)", ("ocr_enabled", str(enable_ocr)))
        conn.execute("INSERT INTO build_info (key, value) VALUES (?, ?)", ("ocr_engine", ocr_engine))
        conn.execute("INSERT INTO build_info (key, value) VALUES (?, ?)", ("force_ocr", str(force_ocr)))

        manifest: list[PdfMetadata] = []
        total_pages = 0
        for index, pdf_path in enumerate(iter_pdfs(evidence_dir), start=1):
            if limit is not None and index > limit:
                break
            meta = parse_pdf_name(pdf_path, evidence_dir)
            pages, extractor = extract_pdf_pages(pdf_path, enable_ocr, ocr_engine, force_ocr)
            status = "ok" if any(page.text for page in pages) else STATUS_NEEDS_REVIEW
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
                    meta.rel_path,
                    str(meta.path),
                    meta.file_name,
                    meta.evidence_no,
                    meta.title,
                    meta.document_date,
                    meta.person_or_source,
                    meta.sha256,
                    meta.file_size,
                    meta.modified_at,
                    "",
                    len(pages),
                    extractor,
                    status,
                    created_at,
                ),
            )
            doc_id = int(cur.lastrowid)
            for page in pages:
                page_norm = normalize_for_search(page.text)
                cur = conn.execute(
                    """
                    INSERT INTO pages (
                        document_id, page_no, text, text_norm, text_hash, extract_method,
                        ocr_confidence, review_status, error
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        doc_id,
                        page.page_no,
                        page.text,
                        page_norm,
                        text_hash(page.text),
                        page.method,
                        page.ocr_confidence,
                        page_review_status(page),
                        page.error,
                    ),
                )
                page_id = int(cur.lastrowid)
                conn.execute(
                    """
                    INSERT INTO pages_fts (text, evidence_no, title, rel_path, page_id)
                    VALUES (?, ?, ?, ?, ?)
                    """,
                    (page_norm, meta.evidence_no, meta.title, meta.rel_path, page_id),
                )
            add_timeline_candidates(conn, meta, pages, max_events_per_page)
            manifest.append(meta)
            total_pages += len(pages)
            print(f"indexed {index}: {meta.rel_path} ({len(pages)} pages)")
        conn.commit()
        return manifest, total_pages
    finally:
        conn.close()


def acquire_lock(index_dir: Path) -> Path:
    index_dir.mkdir(parents=True, exist_ok=True)
    lock_path = index_dir / LOCK_NAME
    try:
        fd = os.open(lock_path, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
    except FileExistsError as exc:
        raise SystemExit(f"index build lock exists: {lock_path}") from exc
    with os.fdopen(fd, "w", encoding="utf-8") as f:
        f.write(f"pid={os.getpid()}\ncreated_at={now_iso()}\n")
    return lock_path


def atomic_copy(src: Path, dst: Path) -> None:
    dst.parent.mkdir(parents=True, exist_ok=True)
    tmp = dst.with_name(dst.name + ".tmp")
    shutil.copy2(src, tmp)
    os.replace(tmp, dst)


def write_manifest(export_dir: Path, db_path: Path) -> None:
    conn = sqlite3.connect(db_path)
    try:
        rows = conn.execute(
            """
            SELECT rel_path, file_name, evidence_no, title, document_date,
                   person_or_source, sha256, file_size, modified_at,
                   acquired_at, page_count, extract_method, extract_status
            FROM documents
            ORDER BY rel_path
            """
        ).fetchall()
    finally:
        conn.close()
    out = export_dir / "manifest.csv"
    tmp = out.with_suffix(".csv.tmp")
    with tmp.open("w", encoding="utf-8", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(
            [
                "rel_path",
                "file_name",
                "evidence_no",
                "title",
                "document_date",
                "person_or_source",
                "sha256",
                "file_size",
                "modified_at",
                "acquired_at",
                "page_count",
                "extract_method",
                "extract_status",
            ]
        )
        writer.writerows(rows)
    os.replace(tmp, out)


def write_timeline(export_dir: Path, db_path: Path) -> None:
    conn = sqlite3.connect(db_path)
    try:
        rows = conn.execute(
            """
            SELECT event_date, document_date, evidence_no, title, rel_path,
                   pdf_page_no, material_page_label, quote_text, review_status,
                   reviewer, reviewed_at, confidence_reason
            FROM timeline_events
            ORDER BY
                CASE WHEN event_date = '' THEN '9999-99-99' ELSE event_date END,
                document_date,
                evidence_no,
                rel_path,
                pdf_page_no
            """
        ).fetchall()
    finally:
        conn.close()

    headers = [
        "event_date",
        "document_date",
        "evidence_no",
        "title",
        "rel_path",
        "pdf_page_no",
        "material_page_label",
        "quote_text",
        "review_status",
        "reviewer",
        "reviewed_at",
        "confidence_reason",
    ]
    csv_out = export_dir / "timeline.csv"
    tmp_csv = csv_out.with_suffix(".csv.tmp")
    with tmp_csv.open("w", encoding="utf-8", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(headers)
        writer.writerows(rows)
    os.replace(tmp_csv, csv_out)

    md_out = export_dir / "timeline.md"
    tmp_md = md_out.with_suffix(".md.tmp")
    with tmp_md.open("w", encoding="utf-8") as f:
        f.write("# 時系列候補表\n\n")
        f.write(
            "この表は開示証拠由来の二次資料です。書面・尋問・証拠意見に使う前に、"
            "必ず原PDF画像で引用と文脈を確認してください。\n\n"
        )
        f.write(
            "| 出来事日 | 文書作成日 | 証拠番号 | 標目 | PDF頁 | 確認 | 引用 |\n"
            "| --- | --- | --- | --- | ---: | --- | --- |\n"
        )
        for row in rows:
            event_date, document_date, evidence_no, title, _rel_path, page_no, _label, quote, status, *_ = row
            safe_quote = quote.replace("|", "｜")
            f.write(
                f"| {event_date} | {document_date} | {evidence_no} | "
                f"{title.replace('|', '｜')} | {page_no or ''} | {status} | {safe_quote} |\n"
            )
    os.replace(tmp_md, md_out)


def write_llm_log_header(export_dir: Path) -> None:
    out = export_dir / "llm_usage_log.csv"
    if out.exists():
        return
    with out.open("w", encoding="utf-8", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(
            [
                "used_at",
                "tool_or_service",
                "purpose",
                "target_files_or_pages",
                "input_chars",
                "output_location",
                "external_ai_used",
                "confirmed_by",
                "confirmation_note",
            ]
        )


def write_readme(export_dir: Path) -> None:
    out = export_dir / "README.md"
    tmp = out.with_suffix(".md.tmp")
    tmp.write_text(
        """# 検察官証拠インデックス

このフォルダの成果物は、開示PDFを検索・時系列化するための二次資料です。
正本は開示PDFそのものです。

## 運用ルール

- 索引DB、OCRテキスト、時系列表、LLM投入用抽出物も開示証拠由来資料として扱う。
- LLM/NotebookLMの出力は参考メモにとどめ、書面・尋問・証拠意見に使う前に原PDF画像で確認する。
- 外部AI投入前には、対象ファイル、個人情報、利用目的、投入先、削除予定を明示して確認を取る。
- `timeline.*` の `review_status` が `原PDF目視確認済み` でない行は、未確認候補として扱う。
- 弁護人メモ、接見メモ、反対尋問メモ、公開資料は、この検察官証拠インデックスに混在させない。
""",
        encoding="utf-8",
    )
    os.replace(tmp, out)


def build_command(args: argparse.Namespace) -> None:
    evidence_dir = resolve_evidence_dir(args)
    if not evidence_dir.is_dir():
        raise SystemExit(f"not a directory: {evidence_dir}")

    index_dir = evidence_dir / INDEX_DIR_NAME
    export_dir = index_dir / EXPORT_DIR_NAME
    lock_path = acquire_lock(index_dir)
    temp_dir = Path(tempfile.mkdtemp(prefix="evidence-index-"))
    try:
        work_db = temp_dir / "evidence_index.sqlite"
        manifest, total_pages = build_database(
            evidence_dir=evidence_dir,
            db_path=work_db,
            enable_ocr=args.ocr,
            ocr_engine=args.ocr_engine,
            force_ocr=args.force_ocr,
            max_events_per_page=args.max_events_per_page,
            limit=args.limit,
        )
        export_dir.mkdir(parents=True, exist_ok=True)
        atomic_copy(work_db, export_dir / "evidence_index.sqlite")
        write_manifest(export_dir, work_db)
        write_timeline(export_dir, work_db)
        write_llm_log_header(export_dir)
        write_readme(export_dir)
        print(f"wrote export: {export_dir}")
        print(f"indexed PDFs: {len(manifest)}")
        print(f"indexed pages: {total_pages}")
    finally:
        try:
            lock_path.unlink()
        except FileNotFoundError:
            pass
        if args.keep_work:
            print(f"kept work dir: {temp_dir}")
        else:
            shutil.rmtree(temp_dir, ignore_errors=True)


def default_db_path(args: argparse.Namespace) -> Path:
    if args.db:
        return Path(args.db).expanduser().resolve()
    evidence_dir = resolve_evidence_dir(args)
    return evidence_dir / INDEX_DIR_NAME / EXPORT_DIR_NAME / "evidence_index.sqlite"


def escape_fts_query(query: str) -> str:
    """Build an FTS5 MATCH expression from a user query (D1 #2/#4).

    Each whitespace-separated token is normalized with normalize_for_search and
    wrapped in double quotes (escaping inner quotes). Multiple tokens are joined
    with OR so that space-separated input behaves as OR search.
    """
    tokens = [normalize_for_search(t) for t in re.split(r"\s+", query.strip()) if t]
    tokens = [t for t in tokens if t]
    quoted = [f'"{t.replace(chr(34), chr(34) + chr(34))}"' for t in tokens]
    return " OR ".join(quoted)


def query_needs_like_fallback(query: str) -> bool:
    """trigram FTS cannot match tokens shorter than 3 chars; use LIKE then."""
    tokens = [normalize_for_search(t) for t in re.split(r"\s+", query.strip()) if t]
    tokens = [t for t in tokens if t]
    if not tokens:
        return False
    return any(len(t) < 3 for t in tokens)


def search_pages(conn: sqlite3.Connection, query: str, limit: int = 50) -> list[dict]:
    """Search page text and return hit rows (D1).

    Uses the trigram FTS index for queries >= 3 normalized chars, and falls
    back to a normalized LIKE scan over pages.text_norm when any token is
    shorter than 3 chars (trigram cannot index those). Returns dicts with
    sha256, page_no, snippet, evidence_no, title.
    """
    if not query.strip():
        return []
    if query_needs_like_fallback(query):
        tokens = [normalize_for_search(t) for t in re.split(r"\s+", query.strip()) if t]
        tokens = [t for t in tokens if t]
        if not tokens:
            return []
        where = " OR ".join("p.text_norm LIKE ?" for _ in tokens)
        params: list[object] = [f"%{t}%" for t in tokens]
        params.append(limit)
        rows = conn.execute(
            f"""
            SELECT d.sha256, d.evidence_no, d.title, p.page_no, p.text
            FROM pages p
            JOIN documents d ON d.id = p.document_id
            WHERE {where}
            LIMIT ?
            """,
            params,
        ).fetchall()
        out: list[dict] = []
        for sha256, evidence_no, title, page_no, text in rows:
            out.append(
                {
                    "sha256": sha256,
                    "evidence_no": evidence_no,
                    "title": title,
                    "page_no": page_no,
                    "snippet": like_snippet(text, tokens),
                }
            )
        return out

    fts_query = escape_fts_query(query)
    if not fts_query:
        return []
    rows = conn.execute(
        """
        SELECT d.sha256, d.evidence_no, d.title, p.page_no,
               snippet(pages_fts, 0, '[', ']', '...', 20) AS hit
        FROM pages_fts
        JOIN pages p ON p.id = pages_fts.page_id
        JOIN documents d ON d.id = p.document_id
        WHERE pages_fts MATCH ?
        LIMIT ?
        """,
        (fts_query, limit),
    ).fetchall()
    return [
        {
            "sha256": sha256,
            "evidence_no": evidence_no,
            "title": title,
            "page_no": page_no,
            "snippet": hit,
        }
        for sha256, evidence_no, title, page_no, hit in rows
    ]


def like_snippet(text: str, tokens: list[str], width: int = 30) -> str:
    """Build a snippet around the first matching token (normalized LIKE path)."""
    norm = normalize_for_search(text)
    pos = -1
    for t in tokens:
        pos = norm.find(t)
        if pos >= 0:
            break
    if pos < 0:
        return excerpt(text, 0, 0)
    return excerpt(text, pos, pos + len(tokens[0]), width)


def search_command(args: argparse.Namespace) -> None:
    db_path = default_db_path(args)
    if not db_path.exists():
        raise SystemExit(f"database not found: {db_path}")
    conn = sqlite3.connect(db_path)
    try:
        rows = search_pages(conn, args.query, args.limit)
    finally:
        conn.close()
    for row in rows:
        print(
            f"{row['evidence_no']}\t{row['title']}\t"
            f"p.{row['page_no']}\t{row['snippet']}"
        )


def log_llm_command(args: argparse.Namespace) -> None:
    evidence_dir = resolve_evidence_dir(args)
    export_dir = evidence_dir / INDEX_DIR_NAME / EXPORT_DIR_NAME
    export_dir.mkdir(parents=True, exist_ok=True)
    write_llm_log_header(export_dir)
    out = export_dir / "llm_usage_log.csv"
    with out.open("a", encoding="utf-8", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(
            [
                now_iso(),
                args.tool_or_service,
                args.purpose,
                args.target_files_or_pages,
                args.input_chars,
                args.output_location,
                args.external_ai_used,
                args.confirmed_by,
                args.confirmation_note,
            ]
        )
    print(f"appended: {out}")


def make_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    subparsers = parser.add_subparsers(dest="command", required=True)

    common_parent = argparse.ArgumentParser(add_help=False)
    common_parent.add_argument("--evidence-dir", default=None, help="Evidence folder. Prefer --path-file for paths with Unicode whitespace.")
    common_parent.add_argument("--path-file", default=str(DEFAULT_PATH_FILE), help="UTF-8 file whose first line is the evidence folder.")

    build = subparsers.add_parser("build", parents=[common_parent], help="Build index artifacts.")
    build.add_argument("--ocr", action="store_true", help="Run local OCR on pages with no embedded text.")
    build.add_argument(
        "--ocr-engine",
        choices=["auto", "vision", "tesseract"],
        default="auto",
        help="OCR engine. auto uses Apple Vision first, then Tesseract fallback.",
    )
    build.add_argument(
        "--force-ocr",
        action="store_true",
        help="OCR every page even when embedded PDF text exists. Slower but often better for poor embedded OCR.",
    )
    build.add_argument("--max-events-per-page", type=int, default=5)
    build.add_argument("--limit", type=int, default=None, help="Limit PDF count for trial runs.")
    build.add_argument("--keep-work", action="store_true", help="Do not remove the local temporary work directory.")
    build.set_defaults(func=build_command)

    search = subparsers.add_parser("search", parents=[common_parent], help="Search the SQLite FTS index.")
    search.add_argument("query")
    search.add_argument("--db", default=None, help="SQLite snapshot path. Defaults to _index/export/evidence_index.sqlite.")
    search.add_argument("--limit", type=int, default=20)
    search.set_defaults(func=search_command)

    log_llm = subparsers.add_parser("log-llm", parents=[common_parent], help="Append an LLM/NotebookLM usage audit row.")
    log_llm.add_argument("--tool-or-service", required=True)
    log_llm.add_argument("--purpose", required=True)
    log_llm.add_argument("--target-files-or-pages", required=True)
    log_llm.add_argument("--input-chars", default="")
    log_llm.add_argument("--output-location", default="")
    log_llm.add_argument("--external-ai-used", default="yes")
    log_llm.add_argument("--confirmed-by", default="")
    log_llm.add_argument("--confirmation-note", default="")
    log_llm.set_defaults(func=log_llm_command)

    return parser


def main() -> None:
    parser = make_parser()
    args = parser.parse_args()
    args.func(args)


if __name__ == "__main__":
    main()
