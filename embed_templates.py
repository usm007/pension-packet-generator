#!/usr/bin/env python3
"""Bake the two DOCX templates into pension_form.html as base64.

The page works by double-clicking it (file://), where browsers block
fetch() of sibling files — so the templates travel inside the HTML.
Re-run this after changing either template; it only touches the two
placeholder lines and leaves everything else byte-identical.
"""
from __future__ import annotations

import base64
from pathlib import Path

HERE = Path(__file__).resolve().parent
HTML = HERE / "pension_form.html"
TEMPLATES = {
    "RETIREMENT": HERE / "Retirement_Certificate_Formats_Blank_Template.docx",
    "KRITAGYATA": HERE / "Data_Required_for_Online_Pension_File_Kritagyata.docx",
}


def main() -> int:
    html = HTML.read_text(encoding="utf-8")
    for key, path in TEMPLATES.items():
        if not path.exists():
            print(f"Missing template: {path.name}")
            return 2
        b64 = base64.b64encode(path.read_bytes()).decode("ascii")
        marker = f"/*__TEMPLATE_{key}_B64__*/"
        head, sep, tail = html.partition(marker)
        if not sep:
            print(f"Marker not found: {marker}")
            return 2
        # Replace only the quoted payload directly after the marker.
        quote, q, rest = tail.partition('"')
        if not q:
            print(f"No payload string after marker: {marker}")
            return 2
        payload, q2, remainder = rest.partition('"')
        if not q2:
            print(f"Unterminated payload string after marker: {marker}")
            return 2
        html = head + marker + quote + q + b64 + q2 + remainder
        print(f"Embedded {path.name} ({len(b64) // 1024} KB base64).")
    HTML.write_text(html, encoding="utf-8")
    print("pension_form.html updated.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
