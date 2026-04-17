from __future__ import annotations

from datetime import datetime, timezone

from .models import PageSnapshot, SiteSnapshot


def _esc(s: object) -> str:
    return str(s).replace("|", "\\|").replace("\n", " ")


def _md_table(headers: list[str], rows: list[list[str]]) -> str:
    lines = ["| " + " | ".join(_esc(h) for h in headers) + " |"]
    lines.append("| " + " | ".join("---" for _ in headers) + " |")
    for row in rows:
        lines.append("| " + " | ".join(_esc(c) for c in row) + " |")
    return "\n".join(lines)


def slugify(title: str) -> str:
    s = title.lower().strip()
    out = []
    for ch in s:
        if ch.isalnum():
            out.append(ch)
        elif ch in " -_":
            out.append("-")
    slug = "".join(out)
    while "--" in slug:
        slug = slug.replace("--", "-")
    return slug.strip("-") or "section"


# ---------------------------------------------------------------------------
# Site-wide QA report
# ---------------------------------------------------------------------------

def _verdict(site: SiteSnapshot) -> str:
    sev = [f.severity for f in site.findings]
    if any(s == "S1" for s in sev):
        return "🔴 **FAIL** — critical issues block release"
    if sum(1 for s in sev if s == "S2") >= 2:
        return "🟡 **CONDITIONAL** — major issues need remediation"
    if sev:
        return "🟢 **PASS WITH NOTES** — minor/moderate issues"
    return "🟢 **PASS** — no automated red flags (manual QA still required)"


def render_qa_report(site: SiteSnapshot) -> str:
    now = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC")
    pages = site.pages

    total_pages = len(pages)
    failed = [p for p in pages if p.load_error]
    ok_pages = [p for p in pages if not p.load_error]

    total_imgs = sum(p.total_images for p in ok_pages)
    total_missing_alt = sum(p.images_without_alt for p in ok_pages)
    total_forms = sum(len(p.forms) for p in ok_pages)
    total_console_errs = sum(len(p.console_errors) for p in ok_pages)
    pages_multi_h1 = [p for p in ok_pages if sum(1 for lvl, _ in p.headings if lvl == 1) > 1]
    pages_no_h1 = [p for p in ok_pages if not any(lvl == 1 for lvl, _ in p.headings)]
    pages_no_meta = [p for p in ok_pages if not p.meta_description]
    pages_no_lang = [p for p in ok_pages if not p.lang]

    # Per-page summary table
    per_page_rows = []
    for p in pages:
        if p.load_error:
            per_page_rows.append([p.final_url or p.url, "ERROR", "-", "-", "-", "-", p.load_error])
            continue
        h1_count = sum(1 for lvl, _ in p.headings if lvl == 1)
        per_page_rows.append([
            p.final_url or p.url,
            "OK",
            p.title or "(empty)",
            f"{h1_count}",
            f"{p.images_without_alt}/{p.total_images}",
            f"{len(p.forms)}",
            f"{len(p.console_errors)}",
        ])

    per_page_md = _md_table(
        ["URL", "Status", "Title", "H1 count", "Missing alt / total img", "Forms", "Console errs"],
        per_page_rows or [["(no pages captured)"] + ["-"] * 6],
    )

    # Top risks (auto-derived)
    risks: list[str] = []
    if failed:
        risks.append(f"**{len(failed)} page(s) failed to load** — investigate first.")
    if pages_multi_h1:
        risks.append(f"{len(pages_multi_h1)} page(s) have multiple `<h1>` — accessibility / SEO issue.")
    if pages_no_h1:
        risks.append(f"{len(pages_no_h1)} page(s) have no `<h1>` at all.")
    if total_imgs and total_missing_alt / total_imgs > 0.1:
        pct = round(100 * total_missing_alt / total_imgs)
        risks.append(f"{pct}% of images ({total_missing_alt}/{total_imgs}) lack `alt` text.")
    if total_console_errs:
        risks.append(f"{total_console_errs} console error(s) across {len([p for p in ok_pages if p.console_errors])} page(s).")
    if pages_no_meta:
        risks.append(f"{len(pages_no_meta)} page(s) missing meta description.")
    if pages_no_lang:
        risks.append(f"{len(pages_no_lang)} page(s) missing `<html lang>`.")
    if not risks:
        risks.append("No automated red flags detected; manual QA still required.")

    # Console errors aggregated
    console_rows: list[list[str]] = []
    for p in ok_pages:
        for err in p.console_errors[:5]:
            console_rows.append([p.final_url or p.url, err])
    if not console_rows:
        console_rows = [["—", "(none captured)"]]

    # Forms aggregated
    forms_md = ""
    form_idx = 0
    for p in ok_pages:
        for f in p.forms:
            form_idx += 1
            forms_md += f"\n### Form {form_idx} — {p.final_url or p.url} ({f.method} `{f.action or '/'}`)\n\n"
            if not f.fields:
                forms_md += "_No fields detected._\n"
                continue
            forms_md += _md_table(
                ["Label / name", "Type", "Required"],
                [[x.label, x.type, str(x.required)] for x in f.fields],
            )
            forms_md += "\n"
    if not forms_md:
        forms_md = "_No `<form>` elements detected on any crawled page._"

    # Manual checklist (as before)
    checklist = _md_table(
        ["Area", "Test ID (optional)", "Result", "Notes"],
        [
            ["Smoke: every crawled page loads without critical errors", " ", "Pass / Fail / Blocked", " "],
            ["Navigation: primary nav consistent across pages", " ", " ", " "],
            ["Forms: validation, required fields, error messages", " ", " ", " "],
            ["Links: no broken internal links (sample)", " ", " ", " "],
            ["Accessibility: landmarks, heading order, alt text", " ", " ", " "],
            ["Responsive: mobile / tablet / desktop breakpoints", " ", " ", " "],
            ["Cross-browser: Chrome, Firefox, Safari (pick matrix)", " ", " ", " "],
            ["Security: HTTPS, cookies flags (if auth)", " ", " ", " "],
            ["Performance: LCP / CLS spot check (field or lab)", " ", " ", " "],
        ],
    )

    defect_template = _md_table(
        ["ID", "Severity", "Page URL", "Summary", "Steps", "Expected", "Actual", "Evidence"],
        [["DEF-001", "S1–S4", " ", " ", " ", " ", " ", "screenshot / HAR"]],
    )

    risks_md = "\n".join(f"- {r}" for r in risks)

    # Findings table
    findings_rows = [
        [f.severity, f.area, f.message, f.reference or "-", str(len(f.pages))]
        for f in sorted(site.findings, key=lambda x: x.severity)
    ]
    findings_md = _md_table(
        ["Severity", "Area", "Finding", "Reference", "Pages"],
        findings_rows,
    ) if findings_rows else "_No automated red flags detected._"

    # Performance table
    perf_rows = [
        [p.final_url or p.url, f"{p.perf.load_ms} ms",
         f"{p.perf.dom_content_loaded_ms} ms",
         f"{p.perf.transfer_bytes/1024:.0f} KB",
         str(p.perf.resource_count)]
        for p in ok_pages
    ]
    perf_md = _md_table(["URL", "Load", "DCL", "Transfer", "Resources"], perf_rows) \
        if perf_rows else "_No performance data captured._"

    return f"""# QA Report (draft) — Site crawl

**Entry URL:** {site.entry_url}
**Pages crawled:** {total_pages}
**Generated:** {now}
**Standard:** ISO/IEC/IEEE 29119-3 §7 Test Report structure

## Executive summary

**Verdict:** {_verdict(site)}

Automated structural pass across {total_pages} page(s). {len(failed)} failed to load. {len(site.findings)} auto-derived finding(s).

| Severity | Count | Meaning |
| --- | ---: | --- |
| S1 | {sum(1 for f in site.findings if f.severity=='S1')} | Production down / data loss |
| S2 | {sum(1 for f in site.findings if f.severity=='S2')} | Major feature broken |
| S3 | {sum(1 for f in site.findings if f.severity=='S3')} | Moderate, workaround exists |
| S4 | {sum(1 for f in site.findings if f.severity=='S4')} | Minor / cosmetic |

## 1. Purpose and scope

Automated structural pass across same-site pages reachable from the entry URL.
This is **not** a full test run — it produces inputs for manual QA and a defect log skeleton.

## 2. Test environment

| Item | Value |
| --- | --- |
| Browser / version | Chromium (Playwright, headless) |
| Viewport | 1280×720 |
| Network | _fill in_ |
| Test data | _fill in_ |

## 3. Overall summary

| Metric | Value |
| --- | ---: |
| Pages crawled | {total_pages} |
| Pages failed to load | {len(failed)} |
| Total forms detected | {total_forms} |
| Total images | {total_imgs} |
| Images missing alt | {total_missing_alt} |
| Console errors (sum) | {total_console_errs} |
| Pages with multiple H1 | {len(pages_multi_h1)} |
| Pages with no H1 | {len(pages_no_h1)} |
| Pages missing meta description | {len(pages_no_meta)} |
| Pages missing `<html lang>` | {len(pages_no_lang)} |

### Top risks (auto-derived)

{risks_md}

## 4. Auto-derived findings

{findings_md}

### Performance snapshot

{perf_md}

## 5. Per-page overview

{per_page_md}

## 6. Console errors (sample across pages)

{_md_table(["Page", "Message"], console_rows)}

## 7. Forms detected across site

{forms_md}

## 8. Manual test checklist

{checklist}

**Severity guide:** S1 production down / data loss · S2 major feature broken · S3 moderate with workaround · S4 minor / cosmetic.

## 9. Defect log (template)

{defect_template}

## 10. Sign-off

| Role | Name | Date | Comments |
| --- | --- | --- | --- |
| QA |  |  |  |
| Product |  |  |  |
"""


# ---------------------------------------------------------------------------
# Site-wide user manual
# ---------------------------------------------------------------------------

def _page_sections(p: PageSnapshot) -> list[tuple[str, list[str]]]:
    """Build (section_title, bullet_lines) from H2-driven outline, with H1 fallback."""
    sections: list[tuple[str, list[str]]] = []
    current: str | None = None
    buf: list[str] = []
    for level, text in p.headings:
        if level == 2:
            if current is not None:
                sections.append((current, buf))
            current = text
            buf = []
        elif level == 3 and current is not None:
            buf.append(f"- {text}")
    if current is not None:
        sections.append((current, buf))

    # Fallback: if no H2, use H1s as section titles
    if not sections:
        h1s = [t for lvl, t in p.headings if lvl == 1]
        if h1s:
            sections = [(t, []) for t in h1s]
        else:
            sections = [(p.title or "Using this page", [])]
    return sections


def render_user_manual(site: SiteSnapshot) -> str:
    now = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC")
    pages = [p for p in site.pages if not p.load_error]
    entry = site.entry_url
    home = pages[0] if pages else None

    site_title = (home.title if home else "") or entry
    intro = (home.meta_description if home else "") or f"This guide explains how to use **{site_title}**."

    # Top-level TOC by page
    toc_lines = []
    for i, p in enumerate(pages, start=1):
        label = p.title or p.final_url or p.url
        toc_lines.append(f"{i}. [{label}](#{slugify(label)})")
    toc = "\n".join(toc_lines) or "_(no pages captured)_"

    import os
    def _rel(path: str) -> str:
        # markdown is written next to output dir; make relative path portable
        return os.path.basename(os.path.dirname(path)) + "/" + os.path.basename(path) if path else ""

    body = ""
    for p in pages:
        label = p.title or p.final_url or p.url
        body += f"\n## {label}\n\n"
        body += f"**URL:** [{p.final_url or p.url}]({p.final_url or p.url})\n\n"
        if p.meta_description:
            body += f"{p.meta_description}\n\n"

        if p.screenshot_viewport:
            body += f"![Above-the-fold view of {label}]({_rel(p.screenshot_viewport)})\n\n"
            body += f"_Figure: Above-the-fold view of **{label}**._\n\n"

        # Detected controls (buttons)
        if p.buttons:
            body += "**Controls on this page:**\n\n"
            body += _md_table(["Control label"], [[b] for b in p.buttons[:15]]) + "\n\n"

        for idx, sp in enumerate(p.form_screenshots, start=1):
            body += f"![Form {idx} on {label}]({_rel(sp)})\n\n"
            body += f"_Figure: Form {idx} on **{label}**._\n\n"

        if p.screenshot_fullpage:
            body += f"<details><summary>Full-page screenshot</summary>\n\n"
            body += f"![Full page of {label}]({_rel(p.screenshot_fullpage)})\n\n</details>\n\n"

        # Sections derived from headings
        sections = _page_sections(p)
        for title, bullets in sections:
            body += f"### {title}\n\n"
            if bullets:
                body += "\n".join(bullets) + "\n\n"
            body += (
                f"1. Open [{p.final_url or p.url}]({p.final_url or p.url}).\n"
                "2. Locate this section in the interface (labels may differ slightly).\n"
                "3. Complete the primary task for this area.\n"
                "4. Confirm the outcome matches what you expect.\n\n"
                "_Add screenshots and role-specific notes here._\n\n"
            )

    return f"""# User manual (draft): {site_title}

**Entry URL:** [{entry}]({entry})
**Pages covered:** {len(pages)}
**Generated:** {now}

## About this manual

{intro}

This document was **machine-generated** by crawling same-site pages and reading their structure (headings, navigation, buttons). Replace placeholders with accurate product language, screenshots, and policies before publication.

## Audience

- **Primary:** _e.g. new customers, internal staff_
- **Prerequisites:** _accounts, hardware, supported browsers_

## Table of contents

{toc}

## Quick start

1. Visit [{entry}]({entry}).
2. Use the main navigation to reach the area you need.
3. Follow the numbered steps in each section below.
{body}

## Troubleshooting

| Symptom | What to try |
| --- | --- |
| Page does not load | Check network; try another browser; clear cache |
| Cannot sign in | Verify credentials; reset password; check caps lock |
| Missing data | Refresh; verify filters; contact support with URL and time |

## Glossary

| Term | Definition |
| --- | --- |
| _add terms_ |  |

## Support

**Contact:** _support email / portal_
**Privacy & terms:** _link to legal pages_
"""
