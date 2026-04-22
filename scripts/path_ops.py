#!/usr/bin/env python3
"""Perform common filesystem operations on paths that contain Unicode
whitespace (e.g. U+3000 全角スペース), **without** putting those paths
on the shell command line.

All target paths are read from `.cursor/shell_path_utf8.txt` (one path
per line) so the invoking shell command stays ASCII-only, which prevents
Cursor's "Contains Unicode whitespace" confirmation dialog.

Usage (command line is ASCII-only in every case):

    venv/bin/python scripts/path_ops.py ls
        # list the first path in the file

    venv/bin/python scripts/path_ops.py open
        # `open` every path in the file (macOS)

    venv/bin/python scripts/path_ops.py cp
        # copy: line 1 = source, line 2 = destination

    venv/bin/python scripts/path_ops.py mv
        # move: line 1 = source, line 2 = destination

    venv/bin/python scripts/path_ops.py stat
        # print absolute path + exists/size for every line

    venv/bin/python scripts/path_ops.py glob '*.docx'
        # glob inside the first path (pattern is the ONLY arg,
        # so keep the pattern ASCII; no shell expansion is applied)
"""

from __future__ import annotations

import argparse
import glob
import os
import shutil
import subprocess
import sys

_REPO_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
_DEFAULT_PATH_FILE = os.path.join(_REPO_ROOT, ".cursor", "shell_path_utf8.txt")


def _read_lines(path_file: str) -> list[str]:
    with open(path_file, encoding="utf-8") as f:
        lines = [ln.strip() for ln in f if ln.strip()]
    if not lines:
        print(f"Path file is empty: {path_file}", file=sys.stderr)
        sys.exit(1)
    return lines


def _cmd_ls(targets: list[str]) -> None:
    target = targets[0]
    if not os.path.isdir(target):
        print(f"Not a directory: {target!r}", file=sys.stderr)
        sys.exit(1)
    for name in sorted(os.listdir(target)):
        print(name)


def _cmd_open(targets: list[str]) -> None:
    for t in targets:
        if not os.path.exists(t):
            print(f"Not found: {t!r}", file=sys.stderr)
            sys.exit(1)
        subprocess.run(["open", t], check=True)


def _cmd_cp(targets: list[str]) -> None:
    if len(targets) < 2:
        print("cp needs 2 lines (src, dst) in the path file", file=sys.stderr)
        sys.exit(1)
    src, dst = targets[0], targets[1]
    if os.path.isdir(src):
        shutil.copytree(src, dst)
    else:
        shutil.copy2(src, dst)
    print(f"copied: {src!r} -> {dst!r}")


def _cmd_mv(targets: list[str]) -> None:
    if len(targets) < 2:
        print("mv needs 2 lines (src, dst) in the path file", file=sys.stderr)
        sys.exit(1)
    src, dst = targets[0], targets[1]
    shutil.move(src, dst)
    print(f"moved: {src!r} -> {dst!r}")


def _cmd_stat(targets: list[str]) -> None:
    for t in targets:
        exists = os.path.exists(t)
        size = os.path.getsize(t) if exists and os.path.isfile(t) else "-"
        kind = (
            "dir"
            if exists and os.path.isdir(t)
            else "file"
            if exists and os.path.isfile(t)
            else "missing"
        )
        print(f"{kind}\t{size}\t{t}")


def _cmd_glob(targets: list[str], pattern: str) -> None:
    target = targets[0]
    if not os.path.isdir(target):
        print(f"Not a directory: {target!r}", file=sys.stderr)
        sys.exit(1)
    for hit in sorted(glob.glob(os.path.join(target, pattern))):
        print(hit)


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "action",
        choices=["ls", "open", "cp", "mv", "stat", "glob"],
    )
    parser.add_argument(
        "extra",
        nargs="?",
        default=None,
        help="extra argument (currently only used by `glob` as the pattern)",
    )
    parser.add_argument(
        "--path-file",
        default=_DEFAULT_PATH_FILE,
        help=f"UTF-8 file with target paths, one per line (default: {_DEFAULT_PATH_FILE})",
    )
    args = parser.parse_args()

    path_file = args.path_file
    if not os.path.isabs(path_file):
        path_file = os.path.join(_REPO_ROOT, path_file)
    path_file = os.path.abspath(path_file)

    try:
        targets = _read_lines(path_file)
    except OSError as e:
        print(f"Cannot read path file: {e}", file=sys.stderr)
        sys.exit(1)

    if args.action == "ls":
        _cmd_ls(targets)
    elif args.action == "open":
        _cmd_open(targets)
    elif args.action == "cp":
        _cmd_cp(targets)
    elif args.action == "mv":
        _cmd_mv(targets)
    elif args.action == "stat":
        _cmd_stat(targets)
    elif args.action == "glob":
        if not args.extra:
            print("glob needs a pattern argument", file=sys.stderr)
            sys.exit(1)
        _cmd_glob(targets, args.extra)


if __name__ == "__main__":
    main()
