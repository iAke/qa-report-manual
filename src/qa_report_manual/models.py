from __future__ import annotations

from dataclasses import dataclass, field


@dataclass
class LinkItem:
    text: str
    href: str


@dataclass
class FormField:
    label: str
    name: str
    type: str
    required: bool


@dataclass
class FormSummary:
    action: str
    method: str
    fields: list[FormField] = field(default_factory=list)


@dataclass
class Finding:
    """An auto-derived QA finding with WCAG/OWASP/heuristic reference."""
    severity: str  # "S1" | "S2" | "S3" | "S4"
    area: str     # "Accessibility" | "SEO" | "Performance" | "Reliability" | "Security"
    message: str
    reference: str = ""  # e.g. "WCAG 1.1.1" or "Core Web Vitals"
    pages: list[str] = field(default_factory=list)


@dataclass
class PerfTiming:
    dom_content_loaded_ms: int = 0
    load_ms: int = 0
    transfer_bytes: int = 0
    resource_count: int = 0


@dataclass
class A11ySignals:
    has_header: bool = False
    has_nav: bool = False
    has_main: bool = False
    has_footer: bool = False
    has_viewport_meta: bool = False
    vague_link_texts: list[str] = field(default_factory=list)  # "click here", "read more"
    heading_order_ok: bool = True


@dataclass
class PageSnapshot:
    url: str
    final_url: str
    title: str
    meta_description: str
    lang: str
    headings: list[tuple[int, str]]  # (level, text)
    nav_links: list[LinkItem]
    content_links: list[LinkItem]
    buttons: list[str]
    forms: list[FormSummary]
    images_without_alt: int
    total_images: int
    console_errors: list[str]
    load_error: str = ""
    screenshot_viewport: str = ""
    screenshot_fullpage: str = ""
    screenshot_mobile: str = ""
    form_screenshots: list[str] = field(default_factory=list)
    perf: PerfTiming = field(default_factory=PerfTiming)
    a11y: A11ySignals = field(default_factory=A11ySignals)


@dataclass
class SiteSnapshot:
    entry_url: str
    pages: list[PageSnapshot] = field(default_factory=list)
    findings: list[Finding] = field(default_factory=list)
    sitemap_count: int = 0
