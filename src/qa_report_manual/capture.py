from __future__ import annotations

import re
import urllib.request
import xml.etree.ElementTree as ET
from pathlib import Path
from urllib.parse import urljoin, urlparse

from playwright.sync_api import Page, sync_playwright

from .models import (
    A11ySignals, Finding, FormField, FormSummary, LinkItem,
    PageSnapshot, PerfTiming, SiteSnapshot,
)

VAGUE_LINKS = {"click here", "here", "read more", "more", "learn more", "this", "link"}
MAX_IMAGE_WIDTH = 1600  # px — downscale to keep .docx slim
JPEG_QUALITY = 80


_SKIP_EXT = (
    ".pdf", ".zip", ".rar", ".7z", ".tar", ".gz",
    ".png", ".jpg", ".jpeg", ".gif", ".svg", ".webp", ".ico",
    ".mp4", ".mp3", ".mov", ".avi", ".wav",
    ".doc", ".docx", ".xls", ".xlsx", ".ppt", ".pptx",
)


def _normalize_url(u: str) -> str:
    try:
        p = urlparse(u)
        scheme = p.scheme or "https"
        netloc = p.netloc.lower()
        path = p.path or "/"
        if path != "/" and path.endswith("/"):
            path = path[:-1]
        return f"{scheme}://{netloc}{path}"
    except Exception:
        return u


def _is_crawlable(href: str) -> bool:
    if not href:
        return False
    if href.startswith(("mailto:", "tel:", "javascript:")):
        return False
    low = href.lower().split("?", 1)[0].split("#", 1)[0]
    return not low.endswith(_SKIP_EXT)


def _same_site(base: str, href: str) -> bool:
    try:
        h = urlparse(href).netloc
        return h == urlparse(base).netloc or not h
    except Exception:
        return False


def _clean_text(s: str) -> str:
    s = re.sub(r"\s+", " ", s or "").strip()
    return s[:500] if len(s) > 500 else s


def _slugify(text: str, fallback: str = "page") -> str:
    s = re.sub(r"[^a-zA-Z0-9]+", "-", text or "").strip("-").lower()
    return (s[:60] or fallback)


def _auto_scroll(page: Page) -> None:
    """Scroll to bottom in steps to trigger lazy-loaded images/components."""
    try:
        page.evaluate(
            """async () => {
                await new Promise((resolve) => {
                    let y = 0;
                    const step = 400;
                    const timer = setInterval(() => {
                        window.scrollBy(0, step);
                        y += step;
                        if (y >= document.body.scrollHeight) {
                            clearInterval(timer);
                            window.scrollTo(0, 0);
                            resolve();
                        }
                    }, 120);
                });
            }"""
        )
        page.wait_for_timeout(400)
    except Exception:
        pass


def _downscale(path: Path, max_w: int = MAX_IMAGE_WIDTH) -> None:
    """Downscale an image in place if wider than max_w (keeps .docx printable)."""
    try:
        from PIL import Image
        with Image.open(path) as im:
            if im.width <= max_w:
                return
            ratio = max_w / im.width
            new_size = (max_w, int(im.height * ratio))
            im = im.resize(new_size, Image.LANCZOS)
            im.save(path, optimize=True)
    except Exception:
        pass


def _capture_screenshots(page: Page, base_dir: Path, index: int, slug: str,
                         mobile_ctx=None) -> tuple[str, str, str, list[str]]:
    """Return (viewport, fullpage, mobile, [form_paths])."""
    base_dir.mkdir(parents=True, exist_ok=True)
    stem = f"{index:02d}-{slug}"
    vp_path = base_dir / f"{stem}-viewport.png"
    fp_path = base_dir / f"{stem}-fullpage.png"
    mb_path = base_dir / f"{stem}-mobile.png"

    try:
        page.screenshot(path=str(vp_path), full_page=False)
        _downscale(vp_path)
    except Exception:
        vp_path = Path("")

    _auto_scroll(page)
    try:
        page.screenshot(path=str(fp_path), full_page=True)
        _downscale(fp_path)
    except Exception:
        fp_path = Path("")

    # Mobile viewport variant
    mb_out = ""
    if mobile_ctx is not None:
        try:
            mpage = mobile_ctx.new_page()
            mpage.goto(page.url, wait_until="load", timeout=30_000)
            mpage.wait_for_timeout(400)
            mpage.screenshot(path=str(mb_path), full_page=False)
            _downscale(mb_path, max_w=900)
            mb_out = str(mb_path)
            mpage.close()
        except Exception:
            mb_out = ""

    form_paths: list[str] = []
    try:
        forms = page.locator("form").all()[:5]
        for i, form in enumerate(forms, start=1):
            try:
                form.scroll_into_view_if_needed(timeout=2000)
                page.wait_for_timeout(120)
                fpath = base_dir / f"{stem}-form{i}.png"
                form.screenshot(path=str(fpath))
                _downscale(fpath)
                form_paths.append(str(fpath))
            except Exception:
                continue
    except Exception:
        pass

    return (
        str(vp_path) if vp_path else "",
        str(fp_path) if fp_path else "",
        mb_out,
        form_paths,
    )


def _collect_perf(page: Page) -> PerfTiming:
    try:
        data = page.evaluate(
            """() => {
                const t = performance.timing || {};
                const nav = performance.getEntriesByType && performance.getEntriesByType('navigation')[0];
                const resources = performance.getEntriesByType ? performance.getEntriesByType('resource') : [];
                const transferBytes = resources.reduce((s, r) => s + (r.transferSize || 0), 0);
                return {
                    dcl: nav ? Math.round(nav.domContentLoadedEventEnd) : Math.max(0, (t.domContentLoadedEventEnd||0) - (t.navigationStart||0)),
                    load: nav ? Math.round(nav.loadEventEnd) : Math.max(0, (t.loadEventEnd||0) - (t.navigationStart||0)),
                    bytes: transferBytes,
                    count: resources.length,
                };
            }"""
        )
        return PerfTiming(
            dom_content_loaded_ms=int(data.get("dcl") or 0),
            load_ms=int(data.get("load") or 0),
            transfer_bytes=int(data.get("bytes") or 0),
            resource_count=int(data.get("count") or 0),
        )
    except Exception:
        return PerfTiming()


def _collect_a11y(page: Page) -> A11ySignals:
    sig = A11ySignals()
    try:
        sig.has_header = page.locator("header, [role='banner']").count() > 0
        sig.has_nav = page.locator("nav, [role='navigation']").count() > 0
        sig.has_main = page.locator("main, [role='main']").count() > 0
        sig.has_footer = page.locator("footer, [role='contentinfo']").count() > 0
        sig.has_viewport_meta = page.locator('meta[name="viewport"]').count() > 0
    except Exception:
        pass
    try:
        texts = page.locator("a[href]").all_inner_texts()
        sig.vague_link_texts = sorted(
            {t.strip() for t in texts if t and t.strip().lower() in VAGUE_LINKS}
        )
    except Exception:
        pass
    return sig


def _fetch_sitemap(entry_url: str, limit: int = 100) -> list[str]:
    """Try /sitemap.xml to seed the crawl queue. Returns same-site page URLs."""
    base = urlparse(entry_url)
    sm_url = f"{base.scheme}://{base.netloc}/sitemap.xml"
    try:
        req = urllib.request.Request(sm_url, headers={"User-Agent": "qa-report-manual/0.2"})
        with urllib.request.urlopen(req, timeout=8) as resp:
            data = resp.read(2_000_000)
    except Exception:
        return []
    urls: list[str] = []
    try:
        root = ET.fromstring(data)
        ns = {"s": "http://www.sitemaps.org/schemas/sitemap/0.9"}
        for loc in root.findall(".//s:loc", ns) or root.findall(".//loc"):
            if loc.text and urlparse(loc.text).netloc == base.netloc:
                urls.append(loc.text.strip())
                if len(urls) >= limit:
                    break
    except Exception:
        return []
    return urls


def capture_page(url: str, timeout_ms: int = 45_000) -> PageSnapshot:
    """Load URL in Chromium and extract structure for report generation."""
    site = capture_site(url, max_pages=1, timeout_ms=timeout_ms)
    return site.pages[0]


def capture_site(
    url: str,
    max_pages: int = 10,
    timeout_ms: int = 45_000,
    screenshots_dir: Path | None = None,
) -> SiteSnapshot:
    """Crawl same-site pages (BFS from entry URL) up to max_pages."""
    start = _normalize_url(url)
    visited: set[str] = set()
    queue: list[str] = [url]
    pages: list[PageSnapshot] = []

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(
            viewport={"width": 1280, "height": 800},
            device_scale_factor=2,
            user_agent=(
                "Mozilla/5.0 (compatible; QAReportManual/0.2; "
                "+https://example.invalid/qa-report-manual)"
            ),
        )
        mobile_ctx = browser.new_context(
            viewport={"width": 390, "height": 844},
            device_scale_factor=2,
            is_mobile=True,
            has_touch=True,
            user_agent=(
                "Mozilla/5.0 (iPhone; CPU iPhone OS 16_0 like Mac OS X) "
                "AppleWebKit/605.1.15 (KHTML, like Gecko) Version/16.0 Mobile/15E148 Safari/604.1"
            ),
        ) if screenshots_dir is not None else None

        # Seed from sitemap if available
        sm_urls = _fetch_sitemap(url, limit=max_pages * 2)
        for sm in sm_urls:
            if sm not in queue:
                queue.append(sm)

        while queue and len(pages) < max_pages:
            current = queue.pop(0)
            norm = _normalize_url(current)
            if norm in visited:
                continue
            visited.add(norm)

            page = context.new_page()
            errs: list[str] = []
            page.on("console", lambda m, errs=errs: errs.append(m.text) if m.type == "error" else None)

            try:
                page.goto(current, wait_until="load", timeout=timeout_ms)
                page.wait_for_timeout(600)
                snap = _extract_snapshot(page)
                snap.console_errors = errs[:50]
                snap.perf = _collect_perf(page)
                snap.a11y = _collect_a11y(page)
                # Heading order: a page should generally not jump e.g. H1 -> H3
                levels = [lvl for lvl, _ in snap.headings]
                snap.a11y.heading_order_ok = all(
                    levels[i] - levels[i - 1] <= 1 for i in range(1, len(levels))
                ) if levels else True

                if screenshots_dir is not None:
                    slug = _slugify(snap.title or urlparse(snap.final_url).path, fallback="page")
                    vp, fp, mb, forms = _capture_screenshots(
                        page, screenshots_dir, len(pages) + 1, slug, mobile_ctx=mobile_ctx
                    )
                    snap.screenshot_viewport = vp
                    snap.screenshot_fullpage = fp
                    snap.screenshot_mobile = mb
                    snap.form_screenshots = forms

                pages.append(snap)

                # enqueue same-site links for BFS
                for li in list(snap.nav_links) + list(snap.content_links):
                    if not _is_crawlable(li.href):
                        continue
                    if not _same_site(start, li.href):
                        continue
                    n = _normalize_url(li.href)
                    if n in visited:
                        continue
                    queue.append(li.href)
            except Exception as exc:
                pages.append(
                    PageSnapshot(
                        url=current, final_url=current, title="", meta_description="",
                        lang="", headings=[], nav_links=[], content_links=[], buttons=[],
                        forms=[], images_without_alt=0, total_images=0,
                        console_errors=errs[:50], load_error=str(exc),
                    )
                )
            finally:
                try:
                    page.close()
                except Exception:
                    pass

        browser.close()

    site = SiteSnapshot(entry_url=url, pages=pages, sitemap_count=len(sm_urls))
    site.findings = derive_findings(site)
    return site


# ---------------------------------------------------------------------------
# Auto-derived findings (WCAG / SEO / performance / reliability heuristics)
# ---------------------------------------------------------------------------

def derive_findings(site: SiteSnapshot) -> list[Finding]:
    out: list[Finding] = []
    ok = [p for p in site.pages if not p.load_error]
    failed = [p for p in site.pages if p.load_error]

    if failed:
        out.append(Finding("S1", "Reliability",
                           f"{len(failed)} page(s) failed to load — blocks users entirely.",
                           reference="OWASP Availability",
                           pages=[p.url for p in failed]))

    # Accessibility — alt text (WCAG 1.1.1)
    total_imgs = sum(p.total_images for p in ok)
    missing_alt = sum(p.images_without_alt for p in ok)
    if total_imgs and missing_alt / total_imgs > 0.25:
        out.append(Finding(
            "S2", "Accessibility",
            f"{round(100 * missing_alt / total_imgs)}% of images lack alt text ({missing_alt}/{total_imgs}).",
            reference="WCAG 1.1.1 Non-text content",
            pages=[p.final_url for p in ok if p.images_without_alt],
        ))
    elif total_imgs and missing_alt:
        out.append(Finding(
            "S3", "Accessibility",
            f"{missing_alt} image(s) lack alt text.",
            reference="WCAG 1.1.1",
            pages=[p.final_url for p in ok if p.images_without_alt],
        ))

    # Headings (WCAG 1.3.1 / 2.4.6)
    multi_h1 = [p for p in ok if sum(1 for lvl, _ in p.headings if lvl == 1) > 1]
    if multi_h1:
        out.append(Finding("S3", "Accessibility",
                           f"{len(multi_h1)} page(s) use multiple <h1>.",
                           reference="WCAG 1.3.1 / 2.4.6",
                           pages=[p.final_url for p in multi_h1]))
    no_h1 = [p for p in ok if not any(lvl == 1 for lvl, _ in p.headings)]
    if no_h1:
        out.append(Finding("S3", "Accessibility",
                           f"{len(no_h1)} page(s) have no <h1>.",
                           reference="WCAG 2.4.6",
                           pages=[p.final_url for p in no_h1]))
    bad_order = [p for p in ok if not p.a11y.heading_order_ok]
    if bad_order:
        out.append(Finding("S4", "Accessibility",
                           f"{len(bad_order)} page(s) skip heading levels.",
                           reference="WCAG 1.3.1",
                           pages=[p.final_url for p in bad_order]))

    # Landmarks
    no_main = [p for p in ok if not p.a11y.has_main]
    if no_main:
        out.append(Finding("S3", "Accessibility",
                           f"{len(no_main)} page(s) missing <main> landmark.",
                           reference="WCAG 1.3.1 / ARIA landmarks",
                           pages=[p.final_url for p in no_main]))

    # Vague link text
    vague = [p for p in ok if p.a11y.vague_link_texts]
    if vague:
        out.append(Finding("S4", "Accessibility",
                           f"{len(vague)} page(s) contain vague link text (e.g. 'click here').",
                           reference="WCAG 2.4.4",
                           pages=[p.final_url for p in vague]))

    # Viewport meta (mobile-friendliness)
    no_vp = [p for p in ok if not p.a11y.has_viewport_meta]
    if no_vp:
        out.append(Finding("S3", "Accessibility",
                           f"{len(no_vp)} page(s) missing <meta name='viewport'>.",
                           reference="WCAG 1.4.10 Reflow",
                           pages=[p.final_url for p in no_vp]))

    # SEO: meta description + lang
    no_meta = [p for p in ok if not p.meta_description]
    if no_meta:
        out.append(Finding("S4", "SEO",
                           f"{len(no_meta)} page(s) missing meta description.",
                           pages=[p.final_url for p in no_meta]))
    no_lang = [p for p in ok if not p.lang]
    if no_lang:
        out.append(Finding("S3", "Accessibility",
                           f"{len(no_lang)} page(s) missing <html lang>.",
                           reference="WCAG 3.1.1",
                           pages=[p.final_url for p in no_lang]))

    # Performance budget (simple thresholds)
    slow = [p for p in ok if p.perf.load_ms > 4000]
    if slow:
        out.append(Finding("S3", "Performance",
                           f"{len(slow)} page(s) have load time > 4s.",
                           reference="Core Web Vitals (LCP budget)",
                           pages=[p.final_url for p in slow]))
    heavy = [p for p in ok if p.perf.transfer_bytes > 3_000_000]
    if heavy:
        out.append(Finding("S3", "Performance",
                           f"{len(heavy)} page(s) transfer > 3 MB.",
                           reference="Web Almanac page-weight budget",
                           pages=[p.final_url for p in heavy]))

    # Console errors (reliability)
    noisy = [p for p in ok if p.console_errors]
    if noisy:
        sev = "S2" if sum(len(p.console_errors) for p in noisy) > 10 else "S3"
        out.append(Finding(sev, "Reliability",
                           f"{len(noisy)} page(s) logged console errors (total {sum(len(p.console_errors) for p in noisy)}).",
                           pages=[p.final_url for p in noisy]))

    return out


def _extract_snapshot(page: Page) -> PageSnapshot:
    base = page.url
    title = _clean_text(page.title() or "")

    meta_description = ""
    lang = ""
    try:
        meta_description = _clean_text(
            page.locator('meta[name="description"]').first.get_attribute("content") or ""
        )
    except Exception:
        pass
    try:
        lang = _clean_text(page.locator("html").first.get_attribute("lang") or "")
    except Exception:
        pass

    headings: list[tuple[int, str]] = []
    for level in range(1, 4):
        for el in page.locator(f"h{level}").all():
            t = _clean_text(el.inner_text())
            if t:
                headings.append((level, t))

    nav_links = _links_in_region(page, "nav, [role='navigation'], header", base)
    content_links = _links_in_region(page, "main, article, [role='main'], body", base)
    buttons = _button_labels(page)

    forms = _summarize_forms(page)
    img_stats = _image_alt_stats(page)

    return PageSnapshot(
        url=base,
        final_url=page.url,
        title=title,
        meta_description=meta_description,
        lang=lang,
        headings=headings,
        nav_links=nav_links,
        content_links=content_links,
        buttons=buttons,
        forms=forms,
        images_without_alt=img_stats[0],
        total_images=img_stats[1],
        console_errors=[],
    )


def _links_in_region(page: Page, selector: str, base: str) -> list[LinkItem]:
    items: list[LinkItem] = []
    seen: set[tuple[str, str]] = set()
    try:
        root = page.locator(selector).first
        if root.count() == 0:
            return items
        for a in root.locator("a[href]").all():
            try:
                href = a.get_attribute("href") or ""
                if href.startswith("#") or href.lower().startswith("javascript:"):
                    continue
                abs_href = urljoin(base, href)
                if not _same_site(base, abs_href):
                    continue
                text = _clean_text(a.inner_text())
                if not text:
                    continue
                key = (text, abs_href)
                if key in seen:
                    continue
                seen.add(key)
                items.append(LinkItem(text=text, href=abs_href))
            except Exception:
                continue
    except Exception:
        pass
    return items[:80]


def _button_labels(page: Page) -> list[str]:
    labels: list[str] = []
    seen: set[str] = set()
    for sel in ("button", "[role='button']"):
        for el in page.locator(sel).all():
            try:
                t = _clean_text(el.inner_text())
                if not t or len(t) > 120:
                    continue
                if t in seen:
                    continue
                seen.add(t)
                labels.append(t)
            except Exception:
                continue
    return labels[:40]


def _summarize_forms(page: Page) -> list[FormSummary]:
    out: list[FormSummary] = []
    for form in page.locator("form").all()[:15]:
        try:
            action = form.get_attribute("action") or ""
            method = (form.get_attribute("method") or "get").upper()
            fields: list[FormField] = []
            for inp in form.locator("input, select, textarea").all()[:40]:
                try:
                    tag = inp.evaluate("e => e.tagName.toLowerCase()")
                    name = inp.get_attribute("name") or ""
                    itype = inp.get_attribute("type") or ("select" if tag == "select" else tag)
                    required = inp.get_attribute("required") is not None
                    label = ""
                    if inp.get_attribute("id"):
                        lid = inp.get_attribute("id")
                        lab = page.locator(f'label[for="{lid}"]').first
                        if lab.count():
                            label = _clean_text(lab.inner_text())
                    if not label:
                        aria = inp.get_attribute("aria-label") or ""
                        label = _clean_text(aria)
                    if not label and name:
                        label = name
                    fields.append(
                        FormField(
                            label=label or "(unlabeled)",
                            name=name,
                            type=itype or "text",
                            required=required,
                        )
                    )
                except Exception:
                    continue
            out.append(FormSummary(action=action, method=method, fields=fields))
        except Exception:
            continue
    return out


def _image_alt_stats(page: Page) -> tuple[int, int]:
    missing = 0
    total = 0
    for img in page.locator("img").all()[:200]:
        try:
            total += 1
            alt = img.get_attribute("alt")
            if alt is None or alt.strip() == "":
                missing += 1
        except Exception:
            continue
    return missing, total
