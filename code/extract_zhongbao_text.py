# -*- coding: utf-8 -*-
"""Extract text from the interim report PDF for structured review."""

from pathlib import Path as _PublicPath
_PUBLIC_ROOT = _PublicPath(__file__).resolve().parents[1]
from pathlib import Path

import pdfplumber


ROOT = Path(str(_PUBLIC_ROOT))
PDF_PATH = ROOT / "中报.pdf"
OUT_PATH = ROOT / "data" / "中报_2026_提取文本.txt"


def main() -> None:
    lines: list[str] = []
    with pdfplumber.open(PDF_PATH) as pdf:
        for i, page in enumerate(pdf.pages, start=1):
            text = page.extract_text() or ""
            lines.append(f"\n===== PAGE {i} =====\n")
            lines.append(text)
    OUT_PATH.write_text("\n".join(lines), encoding="utf-8")
    print(f"extracted {len(pdf.pages)} pages -> {OUT_PATH}")


if __name__ == "__main__":
    main()
