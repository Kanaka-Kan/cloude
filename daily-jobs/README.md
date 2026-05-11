# daily-jobs/

GitHub Actions runs `scripts/scrape_jobs.py` every day at **13:00 UTC** (09:00 ET / 06:00 PT)
and commits the results here.

## Files

- `YYYY-MM-DD.xlsx` — Biostat **Senior Manager / Associate Director** jobs posted
  in the **last 24 hours** at the companies in `scripts/companies.json`.
- `latest.xlsx` — Same as the most recent daily file (stable filename for bookmarking).
- `diagnostics-YYYY-MM-DD.txt` — Per-company scrape status (HTTP errors, hit counts).
  Use this to find companies whose ATS slug is wrong — fix them in `scripts/companies.json`.

## Columns

| 公司 | 职位 | 地点 | 发布时间 | 申请链接 | ATS |
|---|---|---|---|---|---|

## Run locally

```bash
python scripts/scrape_jobs.py
# Or with a different window:
LOOKBACK_HOURS=72 python scripts/scrape_jobs.py
```

## Adding / fixing a company

Edit `scripts/companies.json`. Each entry needs `name`, `careers_url`, and ATS config.

### Workday
```json
{ "name": "ACME Pharma",
  "careers_url": "https://acme.wd1.myworkdayjobs.com/External",
  "ats": "workday", "tenant": "acme", "host": "wd1", "site": "External" }
```
Get `tenant.hostN.myworkdayjobs.com/<site>` by clicking "Search Jobs" on the
company's careers landing page and reading the URL.

### Greenhouse
```json
{ "name": "ACME Biotech",
  "careers_url": "https://boards.greenhouse.io/acmebio",
  "ats": "greenhouse", "slug": "acmebio" }
```

### Lever
```json
{ "ats": "lever", "slug": "acmebio" }
```

### Ashby
```json
{ "ats": "ashby", "slug": "acmebio" }
```

### SmartRecruiters
```json
{ "ats": "smartrecruiters", "slug": "acmebio" }
```

## Match rules (edit in `scripts/scrape_jobs.py`)

- Title regex: `\b(biostat|statistical programm|statistician|stat programm)\b`
- Level regex: `(Sr|Senior) Manager` **or** `Associate Director` / `Assoc. Dir.`
- Posting window: 24h (override with `LOOKBACK_HOURS` env var)
