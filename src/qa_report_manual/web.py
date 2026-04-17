"""Simple Flask web UI: paste a URL → get QA report + user manual."""
from __future__ import annotations

import re
import threading
import uuid
from pathlib import Path

from flask import Flask, abort, jsonify, redirect, render_template_string, request, send_from_directory, url_for

from .capture import capture_site
from .docx_writer import write_qa_report_docx, write_user_manual_docx
from .generators import render_qa_report, render_user_manual
from .xlsx_writer import write_qa_report_xlsx

app = Flask(__name__)
OUT_ROOT = Path("output/web")
OUT_ROOT.mkdir(parents=True, exist_ok=True)

# job_id -> {"status": "running|done|error", "url", "files", "error", "prefix", "pages"}
JOBS: dict[str, dict] = {}


def _prefix(url: str) -> str:
    host = re.sub(r"^https?://", "", url).split("/")[0]
    host = re.sub(r"[^a-zA-Z0-9._-]+", "_", host)[:80]
    return host or "site"


def _run_job(job_id: str, url: str, max_pages: int, screenshots: bool) -> None:
    try:
        prefix = _prefix(url)
        job_dir = OUT_ROOT / job_id
        job_dir.mkdir(parents=True, exist_ok=True)
        shots_dir = (job_dir / f"{prefix}-screenshots") if screenshots else None

        JOBS[job_id]["stage"] = "crawling"
        site = capture_site(url, max_pages=max_pages, timeout_ms=45_000, screenshots_dir=shots_dir)

        JOBS[job_id]["stage"] = "rendering"
        files = []
        (job_dir / f"{prefix}-qa-report.md").write_text(render_qa_report(site), encoding="utf-8")
        files.append(f"{prefix}-qa-report.md")
        (job_dir / f"{prefix}-user-manual.md").write_text(render_user_manual(site), encoding="utf-8")
        files.append(f"{prefix}-user-manual.md")
        write_qa_report_docx(site, job_dir / f"{prefix}-qa-report.docx")
        files.append(f"{prefix}-qa-report.docx")
        write_user_manual_docx(site, job_dir / f"{prefix}-user-manual.docx")
        files.append(f"{prefix}-user-manual.docx")
        write_qa_report_xlsx(site, job_dir / f"{prefix}-qa-report.xlsx")
        files.append(f"{prefix}-qa-report.xlsx")

        JOBS[job_id].update(status="done", files=files, prefix=prefix, pages=len(site.pages), stage="done")
    except Exception as exc:
        JOBS[job_id].update(status="error", error=str(exc), stage="error")


INDEX_HTML = """<!doctype html>
<html lang="en"><head>
<meta charset="utf-8"><title>QA Report & User Manual Generator</title>
<meta name="viewport" content="width=device-width,initial-scale=1">
<style>
  :root { --brand:#0b5fa5; --bg:#f5f7fa; --card:#fff; --muted:#666; }
  *{box-sizing:border-box} body{margin:0;font-family:system-ui,Segoe UI,Roboto,sans-serif;background:var(--bg);color:#222}
  .wrap{max-width:720px;margin:5vh auto;padding:0 20px}
  h1{color:var(--brand);margin:0 0 6px;font-size:28px}
  p.lead{color:var(--muted);margin:0 0 24px}
  .card{background:var(--card);border-radius:12px;box-shadow:0 6px 30px rgba(0,0,0,.06);padding:28px}
  label{display:block;font-weight:600;margin-top:16px;margin-bottom:6px;font-size:14px}
  input[type=url],input[type=number]{width:100%;padding:12px 14px;border:1px solid #ccd;border-radius:8px;font-size:15px}
  input[type=url]:focus,input[type=number]:focus{outline:none;border-color:var(--brand);box-shadow:0 0 0 3px rgba(11,95,165,.15)}
  .row{display:flex;gap:14px}
  .row>div{flex:1}
  .check{display:flex;align-items:center;gap:8px;margin-top:18px;font-size:14px}
  button{margin-top:22px;background:var(--brand);color:#fff;border:0;padding:13px 22px;font-size:16px;font-weight:600;border-radius:8px;cursor:pointer;width:100%}
  button:hover{background:#084a84}
  .hint{font-size:12px;color:var(--muted);margin-top:4px}
  .feat{display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-top:22px;font-size:13px;color:var(--muted)}
  .feat div::before{content:"✓ ";color:var(--brand);font-weight:700}
</style></head><body>
<div class="wrap">
  <h1>QA Report & User Manual Generator</h1>
  <p class="lead">Paste any website URL. Get a QA report and an end-user manual in Markdown + Word, with screenshots.</p>
  <div class="card">
    <form method="post" action="/run">
      <label for="url">Website URL</label>
      <input type="url" id="url" name="url" placeholder="https://example.com" required autofocus>
      <div class="hint">Same-site pages will be crawled from this entry.</div>
      <div class="row">
        <div>
          <label for="max_pages">Max pages</label>
          <input type="number" id="max_pages" name="max_pages" min="1" max="50" value="8">
        </div>
        <div>
          <label>&nbsp;</label>
          <label class="check"><input type="checkbox" name="screenshots" checked> Capture screenshots</label>
        </div>
      </div>
      <button type="submit">Generate reports</button>
    </form>
    <div class="feat">
      <div>Site-wide crawl (BFS)</div>
      <div>Auto-derived risks</div>
      <div>Forms & console check</div>
      <div>Viewport + full-page shots</div>
      <div>Markdown + Word export</div>
      <div>Accessibility heuristics</div>
    </div>
  </div>
</div></body></html>
"""


JOB_HTML = """<!doctype html>
<html lang="en"><head>
<meta charset="utf-8"><title>Job {{ job_id }} — {{ job.status }}</title>
<meta name="viewport" content="width=device-width,initial-scale=1">
{% if job.status == 'running' %}<meta http-equiv="refresh" content="3">{% endif %}
<style>
  :root { --brand:#0b5fa5; --bg:#f5f7fa; --card:#fff; --muted:#666; --ok:#16794a; --err:#b3261e; }
  body{margin:0;font-family:system-ui,Segoe UI,Roboto,sans-serif;background:var(--bg);color:#222}
  .wrap{max-width:720px;margin:5vh auto;padding:0 20px}
  h1{color:var(--brand);margin:0 0 6px;font-size:26px}
  a.back{color:var(--muted);text-decoration:none;font-size:14px}
  a.back:hover{color:var(--brand)}
  .card{background:var(--card);border-radius:12px;box-shadow:0 6px 30px rgba(0,0,0,.06);padding:28px;margin-top:12px}
  .state{display:inline-block;padding:4px 12px;border-radius:999px;font-size:13px;font-weight:600}
  .state.running{background:#fff4d6;color:#805500}
  .state.done{background:#e4f6ec;color:var(--ok)}
  .state.error{background:#fde8e6;color:var(--err)}
  .pulse{display:inline-block;width:10px;height:10px;border-radius:50%;background:#e0a800;margin-right:6px;animation:pulse 1s infinite}
  @keyframes pulse {0%,100%{opacity:1}50%{opacity:.3}}
  ul.files{list-style:none;padding:0;margin:18px 0 0}
  ul.files li{padding:10px 12px;border:1px solid #e2e6ec;border-radius:8px;margin-bottom:8px;display:flex;justify-content:space-between;align-items:center}
  ul.files a{color:var(--brand);text-decoration:none;font-weight:600}
  ul.files a:hover{text-decoration:underline}
  .kv{margin-top:10px;font-size:14px;color:var(--muted)}
  .err{background:#fde8e6;color:var(--err);padding:12px;border-radius:8px;font-family:monospace;font-size:13px;white-space:pre-wrap}
</style></head><body>
<div class="wrap">
  <a class="back" href="/">← New job</a>
  <h1>Job <code>{{ job_id }}</code></h1>
  <div class="card">
    {% if job.status == 'running' %}
      <div class="state running"><span class="pulse"></span>Running — {{ job.stage or 'starting' }}</div>
      <div class="kv">URL: <b>{{ job.url }}</b></div>
      <p style="margin-top:18px;color:#666">This page refreshes every 3 seconds. Crawling + screenshots usually takes 30–90 seconds for ≤10 pages.</p>
    {% elif job.status == 'done' %}
      <div class="state done">✓ Done</div>
      <div class="kv">URL: <b>{{ job.url }}</b> · Pages crawled: <b>{{ job.pages }}</b></div>
      <ul class="files">
        {% for f in job.files %}
        <li><span>{{ f }}</span><a href="{{ url_for('download', job_id=job_id, filename=f) }}">Download</a></li>
        {% endfor %}
      </ul>
    {% else %}
      <div class="state error">✗ Error</div>
      <div class="err">{{ job.error }}</div>
    {% endif %}
  </div>
</div></body></html>
"""


@app.get("/")
def index():
    return render_template_string(INDEX_HTML)


@app.post("/run")
def run():
    url = (request.form.get("url") or "").strip()
    if not url:
        return redirect(url_for("index"))
    if not url.startswith(("http://", "https://")):
        url = "https://" + url
    try:
        max_pages = max(1, min(50, int(request.form.get("max_pages", 8))))
    except ValueError:
        max_pages = 8
    screenshots = request.form.get("screenshots") == "on"

    job_id = uuid.uuid4().hex[:12]
    JOBS[job_id] = {"status": "running", "url": url, "stage": "queued"}
    threading.Thread(target=_run_job, args=(job_id, url, max_pages, screenshots), daemon=True).start()
    return redirect(url_for("status", job_id=job_id))


@app.get("/job/<job_id>")
def status(job_id):
    job = JOBS.get(job_id)
    if not job:
        abort(404)
    return render_template_string(JOB_HTML, job=job, job_id=job_id)


@app.get("/job/<job_id>.json")
def status_json(job_id):
    job = JOBS.get(job_id)
    if not job:
        abort(404)
    return jsonify(job)


@app.get("/download/<job_id>/<path:filename>")
def download(job_id, filename):
    folder = (OUT_ROOT / job_id).resolve()
    if not folder.is_dir():
        abort(404)
    return send_from_directory(folder, filename, as_attachment=True)


def main() -> int:
    import argparse
    p = argparse.ArgumentParser(description="Run the QA report & user manual web UI.")
    p.add_argument("--host", default="127.0.0.1")
    p.add_argument("--port", type=int, default=5001)
    p.add_argument("--debug", action="store_true")
    args = p.parse_args()
    print(f"→ open http://{args.host}:{args.port} in your browser")
    app.run(host=args.host, port=args.port, debug=args.debug, threaded=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
