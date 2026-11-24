# Pharma RFP / Grant Monitor (Windows)

Checks the company URLs listed in the Excel workbook, follows a bounded set of relevant same-site links, and records page-level funding signals. When enabled, Gemini reviews all scraped records after collection to fill missed titles, status, deadlines, eligibility, evidence, and direct application links.

## What it does

- Reads the hyperlink target in each Excel `URL` cell.
- Fetches each listed page with `requests` and BeautifulSoup, with an optional Playwright fallback for JavaScript-heavy pages.
- Applies keyword rules to the page text to identify oncology/neurology RFP signals and changes since the previous run.
- Requires both an oncology/neurology topic and explicit RFP/proposal language before a record can become an RFP alert.
- Searches only the configured announcement terms: request for proposal, RFP, call for proposal, CCP, call for concept proposal, independent medical education/IME, call for grant applications/CGA, funding opportunity, apply for funding, submit a proposal, external funding, call(s) for grants, and CFG/CFGs.
- Verifies that those terms occur in an actual announcement/application context instead of treating a footer, policy, form, or unrelated mention as an opportunity.
- Checks robots.txt/XML sitemaps and follows relevant official links to external grant portals.
- Supports an authorized Playwright `storage_state` session for portals requiring login; CAPTCHA or access challenges are recorded in `access_issue` for manual review and are not bypassed.
- Resolves direct RFP/grant/application document links from landing pages and uses that link in email alerts when available.
- Follows one level of likely RFP/grant/application links from each defined URL (`MAX_LINKS_PER_SEED`, `MAX_DISCOVERED_URLS`).
- Uses Gemini Search grounding in the scheduled workflow to discover additional public oncology/neurology RFP pages outside the workbook and to review the completed scrape.
- Uses Gemini for discovery and post-scrape content review only. Email handling does not use Gemini: URLs come from links found in the scraped page, and email evidence must be present in the source text.
- When a page is blocked by login/CAPTCHA/access controls, Gemini may search for a public official alternative; that candidate must be fetched and verified by the monitor before it can be used.
- Suppresses email alerts when a stated deadline is yesterday or earlier; today and future deadlines remain eligible.
- Suppresses alerts whose deadline is expired or within `ALERT_MIN_LEAD_DAYS` calendar days (default: 7). Malformed dates such as `January 18,; 2026` are normalized before filtering, including after Gemini enrichment.
- Optionally persists each run and the latest record for every monitored opportunity in PostgreSQL.
- Sends an optional email when the existing change-detection rules identify a positive change. Treat alerts as leads to review on the source page.
- Verifies document content before email: unrelated policies/forms/templates are excluded from alerts.
- Writes reports with a stable set of columns. Fields requiring per-opportunity extraction (such as validated dates, eligibility, and application URL) remain blank in this simple mode.

## Windows setup

1. Put the Excel workbook in the same folder as `main.py` and set its name in `.env` (`INPUT_XLSX`).
2. Install Python 3.11 or 3.12, then open PowerShell in this folder and run:

   ```powershell
   py -3.11 -m venv .venv
   .\.venv\Scripts\Activate.ps1
   python -m pip install --upgrade pip
   pip install -r requirements.txt
   python -m playwright install chromium
   Copy-Item .env.example .env
   ```

3. Edit `.env` with your SMTP settings and comma- or semicolon-separated `ALERT_TO` recipients. Use an app password if your mail provider requires one. Do not commit or share `.env`.
4. Run a small check, then a normal run. Dependencies are pinned in
   `requirements.lock`; update that file deliberately when upgrading packages:

   ```powershell
   python main.py --max-urls 2
   python main.py
   python main.py --discover-gemini
   ```

   To use the expanded seed workbook supplied for the enhanced model, run:

   ```powershell
   python main.py --input "D:\Prudhvi\RFP Scrape\new model\RFP Scrape Pharma URLs.xlsx" --discover-gemini
   ```

   The workbook may contain multiple sheets. Each sheet must have `Company` and
   `URL` columns; URL cells may contain friendly labels because the hyperlink
   target is read from the workbook. Set `GEMINI_API_KEY` and enable
   `GEMINI_REVIEW=true` for structured extraction of deadlines, funding,
   geography, eligibility, references, and contact details.

   Gemini discovery/review requires `google-genai` and `GEMINI_API_KEY` in the environment (or `.env`). Discovery fails fast when enabled without an API key.
   `--discover-gemini` enables both external discovery and post-scrape review. For normal
   runs, set `GEMINI_REVIEW=true` to enable only the review stage. Set `GEMINI_MAX_RESULTS`
   to limit discovery, and `MAX_LINKS_PER_SEED` / `MAX_DISCOVERED_URLS` to bound page-link
   expansion. Discovered URLs are saved to `output\gemini_discovered_urls.json`. Gemini does
   not solve CAPTCHAs or bypass authentication. Set `PLAYWRIGHT_STORAGE_STATE` to an
   authorized Playwright state JSON when a permitted portal login is required. `GEMINI_THINKING_BUDGET` controls the
   optional Gemini reasoning budget (default `2048`); the installed SDK/model must support
   thinking configuration, otherwise the call safely falls back to normal generation.

   Gemini discovery uses the Google Search grounding tool (`google_search`) to look for
   public RFP/grant pages that are absent from the workbook. It is opt-in because it
   makes external API calls. Set `GEMINI_API_KEY`, then use `--discover-gemini` or set
   `GEMINI_DISCOVERY=true`. Set `GEMINI_REVIEW=true` to review pages already scraped.
   Gemini may identify missed fields and provide an evidence quote, but the quote is
   accepted for email only when it exactly appears in the scraped page text. Gemini
   does not select application URLs for email. For blocked pages, Gemini may suggest
   a public alternative URL, but the monitor fetches and verifies it first; unverified
   search-result URLs are discarded.

   PostgreSQL is also opt-in. Set `POSTGRES_ENABLED=true` and `POSTGRES_DSN` (or
   `DATABASE_URL`) to a PostgreSQL connection string. The monitor creates
   `rfp_monitor_runs` and `rfp_opportunities` automatically, upserting the latest
   record by company/source URL. Database errors are logged without suppressing local
   CSV, Markdown, Sheets, or email output.

5. Use `python main.py --baseline` to store the current state without sending initial change alerts.

## Report columns

Both CSV files use these columns in order:

`company`, `source_url`, `final_url`, `discovered_from`, `http_status`, `extraction_method`, `page_title`, `opportunity_type`, `opportunity_title`, `opportunity_status`, `status_reason`, `opportunity_id`, `published_date`, `published_date_text`, `deadline_text`, `deadline_date`, `eligibility`, `application_url`, `evidence`, `evidence_quote`, `field_evidence`, `funding_terms`, `funding_amount`, `sponsor`, `sponsor_type`, `therapeutic_areas`, `target_topics`, `focus`, `geography`, `reference_id`, `contact_email`, `source_type`, `source_authority`, `priority`, `relevance_score`, `change_type`, `changed_fields`, `flag`, `gemini_relevance`, `gemini_confidence`, `gemini_summary`, `document_status`, `confidence_score`, `verification_reason`, `email_eligible`, `email_block_reason`, `access_issue`, `checked_at`, `error`.

The tool follows only a bounded, one-level set of same-site links from each workbook URL. It does not recursively crawl a whole site, bypass access controls, or solve CAPTCHAs. Gemini fields are review/enrichment signals and should be verified against the cited public page.

## Output files

- `output/last_run_report.csv` — results for the current run.
- `output/last_alert_candidates.csv` — the exact records that passed the final email-alert filters for the current run.
- `output/rfp_grant_matches.csv` — current page-level matches.
- `output/monitor_state.json` — saved fingerprints used by change detection.
- `output/gemini_discovered_urls.json` — URLs returned by the optional Gemini discovery stage.
- `logs/monitor.log` — run log.

If Google Sheets sync is configured, the cumulative `rfp_grant_matches.csv` contents (including its header row) replace the configured report worksheet on each run. The `Email Alerts` worksheet is replaced on each run with only alerts successfully sent for that scrape. Email alert tabs and messages include company, RFP URL, evidence, RFP/CGA/IME keywords, RFP title, and detected specialty (from the monitor's configured topic terms; `Not identified` if none match). `Email Alerts History` retains sent alerts for a rolling 30 days and adds an `Alert Sent At (UTC)` timestamp; older rows are removed on each run. Older timestamped rows from the previous combined tab are migrated into the history tab, and undated legacy rows are archived separately. Set `GOOGLE_SHEETS_SPREADSHEET_ID`, `GOOGLE_SHEETS_WORKSHEET`, `GOOGLE_SHEETS_EMAIL_WORKSHEET`, `GOOGLE_SHEETS_EMAIL_HISTORY_WORKSHEET`, and `GOOGLE_SERVICE_ACCOUNT_JSON` in `.env`, share the spreadsheet with the service account email from the JSON credentials, and install dependencies with `pip install -r requirements.txt`. A sync error is logged and does not stop local report or email processing.

The relevance score and status signals are heuristic indicators, not confirmation that an RFP or grant is open. Check the cited source page before acting.

`output\opportunities.md` is also generated as a grouped, human-readable opportunity brief with source links and evidence.

## Run automatically on weekdays at 8 PM IST

After a baseline run and a successful email check, create a Weekly task in Windows Task Scheduler for Monday through Friday at 8:00 PM. Set the program to this project's `.venv\Scripts\python.exe`, arguments to `main.py`, and **Start in** to the project folder. The computer must be powered on and connected to the internet. Check `logs\monitor.log` and the output CSVs after the first scheduled run. The scheduled GitHub workflow stores monitor state in the Actions cache and uploads reports as artifacts; it does not commit generated files to `main`.

## Operations and security

- Copy `.env.example` to `.env`; never commit `.env` or files under `credentials\`.
- Rotate SMTP and Google service-account credentials immediately if they have been exposed or shared.
- `logs\monitor.log` is newline-delimited JSON and includes a `run_id` for correlating one execution.
- `output\health.json` reports run status, record counts, failures, and alert counts. A run fails when every monitored URL errors, preventing bad state from being advanced.
- GitHub Actions runs tests, compilation, Ruff linting, and `pip-audit` on pushes and pull requests.
- For scheduled email alerts, configure the repository secrets `SMTP_HOST`, `SMTP_PORT`, `SMTP_USE_TLS`, `SMTP_USERNAME`, `SMTP_PASSWORD`, `ALERT_FROM`, and `ALERT_TO`. The workflow fails fast if SMTP is partially configured.

## Limitations

- Workbook URLs and bounded same-site candidate links are checked; unrelated or deeply nested opportunity pages may still be missed.
- Keyword matching can miss unfamiliar phrasing or flag generic funding language. Page-level matches may not represent a newly announced opportunity.
- Sites requiring login, blocking automated access, or relying on unsupported JavaScript layouts may not be read successfully. Scanned PDFs are not processed in this simple mode.
