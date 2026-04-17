# QA Report & User Manual (from URL)

**Input:** a website URL
**Output:** two drafts — a **QA report** and an **end-user manual** — in Markdown *and* Word (.docx), with screenshots, generated from a real browser pass.

Structure follows **ISO/IEC/IEEE 29119-3 §7 Test Report** (QA report) and **DITA information typing** — Overview/Steps/Reference per page (user manual). Findings are tagged with **WCAG 2.1** / **Core Web Vitals** / **OWASP** references where relevant.

## What the tool does

1. **Crawls same-site pages** via BFS (seeded by the entry URL + `sitemap.xml` if present).
2. **Captures evidence per page**: DOM structure, headings, forms, buttons, console errors, navigation/content links, performance timing (load, DCL, transfer bytes), landmarks, and screenshots (desktop viewport + full-page + mobile iPhone viewport + each form).
3. **Derives findings automatically** with severity S1–S4 (reliability, accessibility, SEO, performance).
4. **Generates**:
   - Site-wide QA report with **Executive summary + Verdict** (PASS / CONDITIONAL / FAIL) and severity-rated findings table.
   - Task-oriented user manual with **Overview → Steps → Reference** per page, numbered figures, and a Word TOC field.

## Setup

```bash
cd qa-report-manual
python3 -m venv .venv
source .venv/bin/activate        # Windows: .venv\Scripts\activate
pip install -r requirements.txt
playwright install chromium
```

## Usage — CLI

```bash
python -m qa_report_manual "https://example.com" -o output
# or, after pip install -e .:
qa-report-manual "https://example.com" -o output --max-pages 10
```

Flags:

- `-o, --out` — output directory (default `./output`)
- `--max-pages` — same-site pages to crawl (default 15)
- `--timeout` — per-page navigation timeout in ms (default 45000)
- `--format` — `md` · `docx` · `both` (default `both`)
- `--no-screenshots` — skip screenshots (faster, small files)

Files written per run (`<host>` from the URL):

- `<host>-qa-report.md` and `.docx`
- `<host>-user-manual.md` and `.docx`
- `<host>-screenshots/` — viewport + fullpage + mobile + form PNGs

## Usage — Web UI

```bash
python -m qa_report_manual.web --port 5055
# or: qa-report-web --port 5055
```

Open http://127.0.0.1:5055, paste a URL, click **Generate reports**. The page polls the job and offers downloads when done.

## QA report structure

1. Executive summary + Verdict + Severity snapshot
2. Purpose & scope
3. Test environment
4. Auto-derived findings (WCAG/OWASP/perf references) + Performance snapshot
5. Per-page overview table
6. Console errors sample
7. Forms detected
8. Manual test checklist
9. Defect log template
10. Sign-off

## User manual structure (per page)

- **Overview** (concept) — meta description + above-the-fold + mobile view
- **Steps** (task) — one numbered task per heading area, with Expected outcome
- **Reference** — controls table, language, heading counts, forms, load time
- **Feedback + Last updated** section at document end

## Best practices reflected

- IEEE 829 / ISO 29119-3 test report layout
- WCAG 2.1 heuristics: alt coverage, heading order, single H1, `<main>` landmark, viewport meta, vague link text, `<html lang>`
- Core Web Vitals thresholds: load > 4s, transfer > 3 MB
- Microsoft Writing Style Guide + Google Developer Docs: task-oriented, imperative, expected outcome stated
- DITA information typing: Concept / Task / Reference per topic
- Numbered figures with captions and alt text
- Word document best practice: heading styles, TOC field, zebra-shaded tables, cover block

## Limits

- Output is a **draft**: human review is required for accuracy, legal/compliance wording, and screenshot redaction.
- Findings are heuristic signals, not a substitute for a proper WCAG audit.
- Respect robots.txt, terms of service, and rate limits when crawling production systems.

## License

MIT (add a `LICENSE` file if you redistribute).
