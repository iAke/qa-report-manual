from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

from .capture import capture_site
from .docx_writer import write_qa_report_docx, write_user_manual_docx
from .generators import render_qa_report, render_user_manual


def _sanitize_filename_part(url: str) -> str:
    host = re.sub(r"^https?://", "", url).split("/")[0]
    host = re.sub(r"[^a-zA-Z0-9._-]+", "_", host)[:80]
    return host or "site"


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Input: web URL → Output: site-wide QA report + user manual (Markdown drafts).",
    )
    parser.add_argument("url", help="Full URL, e.g. https://example.com/path")
    parser.add_argument(
        "-o", "--out", type=Path, default=Path("output"),
        help="Output directory (default: ./output)",
    )
    parser.add_argument(
        "--timeout", type=int, default=45_000,
        help="Navigation timeout in ms (default: 45000)",
    )
    parser.add_argument(
        "--max-pages", type=int, default=15,
        help="Maximum same-site pages to crawl (default: 15)",
    )
    parser.add_argument(
        "--format", choices=["md", "docx", "both"], default="both",
        help="Output format (default: both)",
    )
    parser.add_argument(
        "--no-screenshots", action="store_true",
        help="Disable screenshot capture (faster, smaller output).",
    )
    args = parser.parse_args(argv)

    url = args.url.strip()
    if not url.startswith(("http://", "https://")):
        url = "https://" + url

    out_dir: Path = args.out
    out_dir.mkdir(parents=True, exist_ok=True)
    prefix = _sanitize_filename_part(url)

    screenshots_dir = None if args.no_screenshots else out_dir / f"{prefix}-screenshots"

    try:
        site = capture_site(
            url,
            max_pages=args.max_pages,
            timeout_ms=args.timeout,
            screenshots_dir=screenshots_dir,
        )
    except Exception as exc:
        print(f"Error loading page: {exc}", file=sys.stderr)
        print("Tip: run `playwright install chromium` once after installing dependencies.", file=sys.stderr)
        return 1

    written: list[Path] = []
    if args.format in ("md", "both"):
        qa_md = out_dir / f"{prefix}-qa-report.md"
        man_md = out_dir / f"{prefix}-user-manual.md"
        qa_md.write_text(render_qa_report(site), encoding="utf-8")
        man_md.write_text(render_user_manual(site), encoding="utf-8")
        written += [qa_md, man_md]

    if args.format in ("docx", "both"):
        qa_docx = out_dir / f"{prefix}-qa-report.docx"
        man_docx = out_dir / f"{prefix}-user-manual.docx"
        write_qa_report_docx(site, qa_docx)
        write_user_manual_docx(site, man_docx)
        written += [qa_docx, man_docx]

    lines = "\n".join(f"  {p.resolve()}" for p in written)
    print(f"Crawled {len(site.pages)} page(s). Wrote:\n{lines}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
