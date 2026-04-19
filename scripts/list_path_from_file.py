#!/usr/bin/env python3
"""List a directory whose path contains Unicode spaces (e.g. U+3000).

The target path is read from a UTF-8 file so the *shell command line* stays
ASCII-only and Cursor does not show "Contains Unicode whitespace".
"""

from __future__ import annotations

import argparse
import os
import sys

_REPO_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
_DEFAULT_PATH_FILE = os.path.join(_REPO_ROOT, ".cursor", "shell_path_utf8.txt")


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--path-file",
        default=_DEFAULT_PATH_FILE,
        help=f"UTF-8 file: one line, absolute directory path (default: {_DEFAULT_PATH_FILE})",
    )
    args = parser.parse_args()
    path_file = args.path_file
    if not os.path.isabs(path_file):
        path_file = os.path.join(_REPO_ROOT, path_file)
    path_file = os.path.abspath(path_file)

    try:
        with open(path_file, encoding="utf-8") as f:
            target = f.read().strip()
    except OSError as e:
        print(f"Cannot read path file: {e}", file=sys.stderr)
        sys.exit(1)

    if not target:
        print("Path file is empty.", file=sys.stderr)
        sys.exit(1)
    if not os.path.isdir(target):
        print(f"Not a directory: {target!r}", file=sys.stderr)
        sys.exit(1)

    for name in sorted(os.listdir(target)):
        print(name)


if __name__ == "__main__":
    main()
