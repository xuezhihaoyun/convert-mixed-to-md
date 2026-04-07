#!/usr/bin/env python3
from __future__ import annotations

import re
import sys


DRIVE_PATH_RE = re.compile(r"[A-Za-z]:\\")
QUOTED_RE = re.compile(r'"([^"]+)"')


def parse_input(raw: str) -> list[str]:
    text = (raw or "").strip()
    if not text:
        return []

    paths: list[str] = []

    # 1) Quoted paths first: "C:\a b.docx" "D:\x.pdf"
    for match in QUOTED_RE.finditer(text):
        candidate = match.group(1).strip()
        if candidate:
            paths.append(candidate)

    # Remove quoted chunks from remaining text.
    remainder = QUOTED_RE.sub(" ", text).strip()

    # 2) Split concatenated drive paths: C:\a.docxC:\b.pdf
    if DRIVE_PATH_RE.search(remainder):
        starts = [m.start() for m in DRIVE_PATH_RE.finditer(remainder)]
        for idx, start in enumerate(starts):
            end = starts[idx + 1] if idx + 1 < len(starts) else len(remainder)
            chunk = remainder[start:end].strip().strip('"').strip()
            if chunk:
                for part in re.split(r"\s*;\s*", chunk):
                    part = part.strip().strip('"').strip()
                    if part:
                        paths.append(part)
    elif remainder:
        for part in re.split(r"\s*;\s*", remainder):
            part = part.strip().strip('"').strip()
            if part:
                paths.append(part)

    # De-duplicate while preserving order.
    seen: set[str] = set()
    result: list[str] = []
    for p in paths:
        if p not in seen:
            seen.add(p)
            result.append(p)
    return result


def main() -> int:
    raw = sys.argv[1] if len(sys.argv) > 1 else ""
    for item in parse_input(raw):
        print(item)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
