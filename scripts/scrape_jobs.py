"""Daily biostatistics job scraper.

Reads scripts/companies.json (a list of companies + ATS config), queries each
company's public job API, filters to:
  - title matches Biostatistics (or Biostat, Statistical Programming)
  - level matches Senior Manager / Sr. Manager / Associate Director / Assoc. Dir.
  - posted within the last LOOKBACK_HOURS (default 24)

Writes:
  daily-jobs/YYYY-MM-DD.xlsx           — matched jobs
  daily-jobs/latest.xlsx               — same as today, stable filename
  daily-jobs/diagnostics-YYYY-MM-DD.txt — per-company status (so you can fix broken slugs)
"""
import json
import os
import re
import sys
import ssl
import time
import urllib.error
import urllib.parse
import urllib.request
from datetime import datetime, timezone, timedelta
from pathlib import Path

# Allow running as `python scripts/scrape_jobs.py` from repo root
ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "scripts"))
from xlsx_writer import write_xlsx  # noqa: E402

LOOKBACK_HOURS = int(os.environ.get("LOOKBACK_HOURS", "24"))
DEFAULT_UA = "daily-jobs-bot/1.0 (+github.com/Kanaka-Kan/cloude)"

# ── Match rules ──────────────────────────────────────────────────────────────
TITLE_KEYWORDS = re.compile(
    r"\b(biostat|statistical\s+programm|statistician|stat\.?\s+programm)\b",
    re.IGNORECASE,
)
LEVEL_PATTERNS = [
    re.compile(r"\b(sr\.?|senior)\s+manager\b", re.IGNORECASE),
    re.compile(r"\bassoc(iate|\.)?\s+director\b", re.IGNORECASE),
    re.compile(r"\bassoc\.?\s+dir\b", re.IGNORECASE),
    # Some companies use "AD" suffix
    re.compile(r"\bassociate\s+director\b", re.IGNORECASE),
]


def title_matches(title: str) -> bool:
    if not title:
        return False
    if not TITLE_KEYWORDS.search(title):
        return False
    return any(p.search(title) for p in LEVEL_PATTERNS)


# ── HTTP helpers ─────────────────────────────────────────────────────────────
def _http(req, timeout=25):
    ctx = ssl.create_default_context()
    return urllib.request.urlopen(req, timeout=timeout, context=ctx)


def http_get(url, headers=None, timeout=25):
    hdrs = {"User-Agent": DEFAULT_UA, "Accept": "application/json,*/*"}
    if headers:
        hdrs.update(headers)
    req = urllib.request.Request(url, headers=hdrs)
    with _http(req, timeout=timeout) as resp:
        return resp.read().decode("utf-8", errors="replace")


def http_post_json(url, payload, headers=None, timeout=25):
    hdrs = {
        "User-Agent": DEFAULT_UA,
        "Content-Type": "application/json",
        "Accept": "application/json",
    }
    if headers:
        hdrs.update(headers)
    req = urllib.request.Request(
        url, data=json.dumps(payload).encode("utf-8"), headers=hdrs
    )
    with _http(req, timeout=timeout) as resp:
        return resp.read().decode("utf-8", errors="replace")


def parse_iso(s):
    if not s:
        return None
    s = str(s).replace("Z", "+00:00")
    try:
        dt = datetime.fromisoformat(s)
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=timezone.utc)
        return dt
    except Exception:
        return None


# ── ATS adapters ─────────────────────────────────────────────────────────────
def fetch_greenhouse(cfg):
    """Greenhouse public board API."""
    slug = cfg["slug"]
    url = f"https://boards-api.greenhouse.io/v1/boards/{slug}/jobs?content=true"
    data = json.loads(http_get(url))
    out = []
    for j in data.get("jobs", []):
        loc = (j.get("location") or {}).get("name", "")
        out.append({
            "title": j.get("title", ""),
            "url": j.get("absolute_url", ""),
            "location": loc,
            "posted_at": parse_iso(j.get("updated_at") or j.get("first_published")),
            "posted_at_text": j.get("updated_at") or j.get("first_published") or "",
        })
    return out


def fetch_lever(cfg):
    slug = cfg["slug"]
    url = f"https://api.lever.co/v0/postings/{slug}?mode=json"
    data = json.loads(http_get(url))
    out = []
    for j in data:
        ts = j.get("createdAt")
        posted = datetime.fromtimestamp(ts / 1000, tz=timezone.utc) if ts else None
        cats = j.get("categories") or {}
        out.append({
            "title": j.get("text", ""),
            "url": j.get("hostedUrl", ""),
            "location": cats.get("location", ""),
            "posted_at": posted,
            "posted_at_text": str(j.get("createdAtStr", "")) or (posted.isoformat() if posted else ""),
        })
    return out


def fetch_ashby(cfg):
    slug = cfg["slug"]
    url = f"https://api.ashbyhq.com/posting-api/job-board/{slug}?includeCompensation=false"
    data = json.loads(http_get(url))
    out = []
    for j in data.get("jobs", []):
        out.append({
            "title": j.get("title", ""),
            "url": j.get("jobUrl", ""),
            "location": j.get("locationName", ""),
            "posted_at": parse_iso(j.get("publishedAt")),
            "posted_at_text": j.get("publishedAt") or "",
        })
    return out


def fetch_smartrecruiters(cfg):
    slug = cfg["slug"]
    url = f"https://api.smartrecruiters.com/v1/companies/{slug}/postings?limit=100"
    data = json.loads(http_get(url))
    out = []
    for j in data.get("content", []):
        city = (j.get("location") or {}).get("city") or ""
        country = (j.get("location") or {}).get("country") or ""
        loc = ", ".join(p for p in [city, country] if p)
        out.append({
            "title": j.get("name", ""),
            "url": f"https://jobs.smartrecruiters.com/{slug}/{j.get('id', '')}",
            "location": loc,
            "posted_at": parse_iso(j.get("releasedDate") or j.get("createdOn")),
            "posted_at_text": j.get("releasedDate") or j.get("createdOn") or "",
        })
    return out


def _workday_relative_date(text):
    """Workday postedOn fields are like 'Posted Today', 'Posted Yesterday', 'Posted 30+ Days Ago'."""
    if not text:
        return None
    t = str(text).lower()
    now = datetime.now(timezone.utc)
    if "today" in t:
        return now
    if "yesterday" in t:
        return now - timedelta(days=1)
    m = re.search(r"(\d+)\+?\s*day", t)
    if m:
        return now - timedelta(days=int(m.group(1)))
    m = re.search(r"(\d+)\+?\s*hour", t)
    if m:
        return now - timedelta(hours=int(m.group(1)))
    m = re.search(r"(\d+)\+?\s*month", t)
    if m:
        return now - timedelta(days=int(m.group(1)) * 30)
    return None


def fetch_workday(cfg):
    tenant = cfg["tenant"]
    site = cfg["site"]
    host = cfg.get("host", "wd1")
    base = f"https://{tenant}.{host}.myworkdayjobs.com"
    api = f"{base}/wday/cxs/{tenant}/{site}/jobs"
    out = []
    # Use searchText to narrow at the API level — saves a lot of pagination.
    for search_term in ("biostatistics", "statistical programming"):
        offset = 0
        seen = set()
        while True:
            payload = {
                "appliedFacets": {},
                "limit": 20,
                "offset": offset,
                "searchText": search_term,
            }
            try:
                text = http_post_json(api, payload, headers={"Origin": base, "Referer": f"{base}/{site}"})
                data = json.loads(text)
            except Exception:
                break
            postings = data.get("jobPostings") or []
            if not postings:
                break
            for j in postings:
                key = j.get("externalPath", "")
                if key in seen:
                    continue
                seen.add(key)
                full_url = f"{base}{key}" if key.startswith("/") else f"{base}/{site}{key}"
                posted_text = j.get("postedOn", "")
                out.append({
                    "title": j.get("title", ""),
                    "url": full_url,
                    "location": j.get("locationsText", ""),
                    "posted_at": _workday_relative_date(posted_text),
                    "posted_at_text": posted_text,
                })
            offset += len(postings)
            if len(postings) < 20 or offset > 200:
                break
    return out


ADAPTERS = {
    "greenhouse": fetch_greenhouse,
    "lever": fetch_lever,
    "ashby": fetch_ashby,
    "smartrecruiters": fetch_smartrecruiters,
    "workday": fetch_workday,
}


# ── Main scrape loop ─────────────────────────────────────────────────────────
def scrape(companies, lookback_hours=LOOKBACK_HOURS):
    cutoff = datetime.now(timezone.utc) - timedelta(hours=lookback_hours)
    rows = []
    diagnostics = []
    for c in companies:
        name = c.get("name", "<unknown>")
        ats = c.get("ats")
        if ats not in ADAPTERS:
            diagnostics.append(f"SKIP {name}: no adapter for ats={ats!r}")
            continue
        t0 = time.time()
        try:
            jobs = ADAPTERS[ats](c)
        except urllib.error.HTTPError as e:
            diagnostics.append(f"ERR  {name} [{ats}]: HTTP {e.code} {e.reason}")
            continue
        except Exception as e:
            diagnostics.append(f"ERR  {name} [{ats}]: {type(e).__name__}: {e}")
            continue
        elapsed = time.time() - t0
        n_total = len(jobs)
        # filter by title
        title_hits = [j for j in jobs if title_matches(j["title"])]
        # filter by recency: if posted_at is known, must be >= cutoff
        # if posted_at is None (e.g. some Workday entries), keep it (will note)
        recent = []
        for j in title_hits:
            pa = j.get("posted_at")
            if pa is None:
                recent.append(j)  # unknown date — keep, flagged
            elif pa >= cutoff:
                recent.append(j)
        diagnostics.append(
            f"OK   {name} [{ats}]: scanned={n_total} title_hits={len(title_hits)} "
            f"in_window={len(recent)} ({elapsed:.1f}s)"
        )
        for j in recent:
            rows.append({
                "company": name,
                "title": j["title"],
                "location": j.get("location", ""),
                "posted": j.get("posted_at_text", "") or (
                    j["posted_at"].isoformat() if j.get("posted_at") else "unknown"
                ),
                "url": j["url"],
                "ats": ats,
            })
    return rows, diagnostics


def main():
    companies = json.loads((ROOT / "scripts" / "companies.json").read_text())
    out_dir = ROOT / "daily-jobs"
    out_dir.mkdir(exist_ok=True)
    today = datetime.now(timezone.utc).strftime("%Y-%m-%d")

    rows, diagnostics = scrape(companies)

    headers = ["公司", "职位", "地点", "发布时间", "申请链接", "ATS"]
    widths = [32, 60, 36, 26, 70, 14]
    table = [[r["company"], r["title"], r["location"], r["posted"], r["url"], r["ats"]] for r in rows]

    daily_xlsx = out_dir / f"{today}.xlsx"
    latest_xlsx = out_dir / "latest.xlsx"
    write_xlsx(str(daily_xlsx), headers, table,
               sheet_name=f"Biostat jobs {today}", widths=widths)
    write_xlsx(str(latest_xlsx), headers, table,
               sheet_name=f"Biostat jobs {today}", widths=widths)

    diag_path = out_dir / f"diagnostics-{today}.txt"
    diag_path.write_text(
        f"Run: {datetime.now(timezone.utc).isoformat()}\n"
        f"Lookback: {LOOKBACK_HOURS}h\n"
        f"Companies: {len(companies)}\n"
        f"Matches: {len(rows)}\n\n"
        + "\n".join(diagnostics) + "\n"
    )

    print(f"Wrote {daily_xlsx} ({len(rows)} matches)")
    print(f"Wrote {diag_path}")


if __name__ == "__main__":
    main()
