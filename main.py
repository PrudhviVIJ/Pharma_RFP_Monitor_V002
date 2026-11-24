from __future__ import annotations
import argparse
import csv
import hashlib
import html
import ipaddress
import json
import logging
import os
import re
import smtplib
import socket
import tempfile
import time
import uuid
from datetime import datetime, timedelta, timezone
from email.message import EmailMessage
from pathlib import Path
from urllib.parse import urljoin, urlparse
from zoneinfo import ZoneInfo
import requests
from bs4 import BeautifulSoup
from dotenv import load_dotenv
from openpyxl import load_workbook
from io import BytesIO
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
from gemini_enrichment import (
    GeminiEnricher,
    extract_candidate_links,
    merge_enrichment,
    resolve_application_url,
    save_discovered,
    discover_sitemap_links,
)
from report_csv import read_records, write_records
from opportunity_report import write_markdown
from quality_gate import (
    ANNOUNCEMENT_TERMS,
    detect_access_issue,
    detect_login_only_page,
    email_url_block_reason,
    verify_opportunity,
)
from postgres_store import persist_output_files, persist_run

try:
    from playwright.sync_api import sync_playwright
except ImportError:
    sync_playwright = None

ROOT = Path(__file__).resolve().parent
load_dotenv(ROOT / ".env")
UTC = timezone.utc
INPUT_XLSX = Path(os.getenv("INPUT_XLSX", str(ROOT / "RFP Scrape Pharma URLs.xlsx")))
OUTPUT_DIR = ROOT / "output"
LOG_DIR = ROOT / "logs"
STATE_FILE = OUTPUT_DIR / "monitor_state.json"
MATCHES_FILE = OUTPUT_DIR / "rfp_grant_matches.csv"
RUN_REPORT = OUTPUT_DIR / "last_run_report.csv"
ALERT_REPORT = OUTPUT_DIR / "last_alert_candidates.csv"
DAILY_RUN_LOG_DIR = OUTPUT_DIR / "daily_run_logs"
MARKDOWN_REPORT = OUTPUT_DIR / "opportunities.md"
REPORT_COLUMNS = [
    "company", "source_url", "final_url", "discovered_from", "http_status",
    "extraction_method", "page_title", "opportunity_type", "opportunity_title",
    "opportunity_status", "status_reason", "opportunity_id", "published_date",
    "published_date_text", "deadline_text", "deadline_date", "eligibility",
    "application_url", "evidence", "evidence_quote", "field_evidence", "funding_terms", "funding_amount",
    "sponsor", "sponsor_type", "therapeutic_areas", "target_topics", "focus", "geography", "reference_id",
    "contact_email", "source_type", "source_authority", "priority", "relevance_score", "change_type",
    "changed_fields", "flag", "gemini_relevance", "gemini_confidence",
    "gemini_summary", "document_status", "confidence_score", "verification_reason",
    "email_eligible", "email_block_reason", "access_issue", "checked_at", "error",
]
DAILY_RUN_LOG_COLUMNS = [
    "run_date", "run_id", "checked_at", "company", "source_url", "final_url",
    "page_title", "opportunity_type", "opportunity_title", "opportunity_status",
    "target_topics", "deadline_date", "deadline_text", "change_type", "finding_status",
    "finding_reason", "alert_status", "alert_reason", "alert_sent_at", "email_configured",
    "email_attempted", "http_status", "extraction_method", "evidence", "error",
]

# Tune these lists to reduce or increase alert sensitivity.
FUNDING_TERMS = list(ANNOUNCEMENT_TERMS)
TOPIC_TERMS = [
    "oncology", "cancer", "tumor", "tumour", "hematology", "haematology",
    "pulmonary fibrosis", "hemophilia", "haemophilia", "gene therapy",
    "mucopolysaccharidosis", "fatty acid oxidation", "glycogen storage disease",
    "neuroscience", "neurology", "neurologic", "neurological", "parkinson",
    "migraine", "epilepsy", "multiple sclerosis", "alzheimer", "dementia",
    "brain", "neurodegenerative", "immuno-oncology", "precision oncology",
    "breast cancer", "lung cancer", "colorectal cancer", "prostate cancer",
    "pancreatic cancer", "ovarian cancer", "lymphoma", "leukemia", "leukaemia",
    "multiple myeloma", "solid tumor", "solid tumour", "tumor microenvironment",
    "car t", "cell therapy", "radiotherapy", "stroke", "cerebrovascular",
    "movement disorder", "alzheimer's disease", "seizure", "headache",
    "neuromuscular", "neuroimmunology", "motor neuron disease", "amyotrophic lateral sclerosis",
    "spinal muscular atrophy", "huntington's disease", "neuropathy", "cognitive impairment",
]
TARGET_TOPIC_TERMS = [
    "oncology", "cancer", "tumor", "tumour", "hematology", "haematology",
    "immuno-oncology", "precision oncology", "breast cancer", "lung cancer",
    "colorectal cancer", "prostate cancer", "pancreatic cancer", "ovarian cancer",
    "lymphoma", "leukemia", "leukaemia", "multiple myeloma", "solid tumor",
    "solid tumour", "tumor microenvironment", "car t", "neuroscience", "neurology",
    "neurologic", "neurological", "parkinson", "migraine", "epilepsy",
    "multiple sclerosis", "alzheimer", "dementia", "neurodegenerative", "brain",
    "stroke", "cerebrovascular", "movement disorder", "seizure", "headache",
    "neuromuscular", "neuroimmunology", "motor neuron disease",
    "amyotrophic lateral sclerosis", "spinal muscular atrophy", "huntington's disease",
    "neuropathy", "cognitive impairment",
]
OPEN_TERMS = [
    "now accepting", "accepting applications", "currently accepting",
    "application deadline", "due date", "submit by", "deadline",
    "open for applications", "applications are open", "proposal deadline",
]
CLOSED_TERMS = [
    "closed", "expired", "deadline has passed", "no longer accepting",
    "applications are no longer being accepted",
]
RFP_TERMS = ["request for proposal", "requests for proposals", "rfp", "call for proposals", "calls for proposals"]
EMAIL_CATEGORY_PATTERNS = {
    "RFP": re.compile(r"\b(?:RFP|requests? for proposals?|calls? for proposals?)\b", re.I),
    "CGA": re.compile(r"\bCGA\b|\bcall for grant applications?\b", re.I),
    "IME": re.compile(r"\bIME\b|\bindependent medical education\b", re.I),
}
MAX_TEXT_CHARS = int(os.getenv("MAX_TEXT_CHARS", "1500000"))
MAX_RESPONSE_BYTES = int(os.getenv("MAX_RESPONSE_BYTES", "15000000"))
MAX_REDIRECTS = int(os.getenv("MAX_REDIRECTS", "5"))
REQUEST_TIMEOUT = int(os.getenv("REQUEST_TIMEOUT", "25"))
DELAY_SECONDS = float(os.getenv("DELAY_SECONDS", "1.0"))
USE_PLAYWRIGHT_FALLBACK = os.getenv("USE_PLAYWRIGHT_FALLBACK", "true").lower() in {"1", "true", "yes"}
PLAYWRIGHT_STORAGE_STATE = os.getenv("PLAYWRIGHT_STORAGE_STATE", "").strip()
USER_AGENT = os.getenv("USER_AGENT", "PharmaRFPMonitor/1.0 (polite public-page monitoring)")
GEMINI_DISCOVERY = os.getenv("GEMINI_DISCOVERY", "false").lower() in {"1", "true", "yes"}
GEMINI_MAX_RESULTS = int(os.getenv("GEMINI_MAX_RESULTS", "50"))
GEMINI_DISCOVERED_FILE = OUTPUT_DIR / "gemini_discovered_urls.json"
GEMINI_REVIEW = os.getenv("GEMINI_REVIEW", "false").lower() in {"1", "true", "yes"}
MAX_LINKS_PER_SEED = int(os.getenv("MAX_LINKS_PER_SEED", "20"))
MAX_DISCOVERED_URLS = int(os.getenv("MAX_DISCOVERED_URLS", "200"))
MAX_FAILURE_RATIO = float(os.getenv("MAX_FAILURE_RATIO", "0.50"))
ALERT_MIN_LEAD_DAYS = int(os.getenv("ALERT_MIN_LEAD_DAYS", "7"))
MONITOR_TIMEZONE = ZoneInfo("Asia/Kolkata")
HEALTH_FILE = OUTPUT_DIR / "health.json"
RUN_ID = uuid.uuid4().hex[:12]

class JsonFormatter(logging.Formatter):
    """Emit machine-readable logs while retaining the human message."""

    def format(self, record):
        payload = {
            "timestamp": datetime.fromtimestamp(record.created, UTC).isoformat(),
            "level": record.levelname,
            "logger": record.name,
            "run_id": RUN_ID,
            "message": record.getMessage(),
        }
        if record.exc_info:
            payload["exception"] = self.formatException(record.exc_info)
        return json.dumps(payload, ensure_ascii=False)

def setup_logging():
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    formatter = JsonFormatter()
    file_handler = logging.FileHandler(LOG_DIR / "monitor.log", encoding="utf-8")
    stream_handler = logging.StreamHandler()
    file_handler.setFormatter(formatter)
    stream_handler.setFormatter(formatter)
    logging.basicConfig(
        level=logging.INFO,
        handlers=[file_handler, stream_handler],
        force=True,
    )

def atomic_write_json(path: Path, value: object) -> None:
    """Write JSON durably so an interrupted run cannot corrupt the target."""
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, temporary_name = tempfile.mkstemp(prefix=f".{path.name}.", suffix=".tmp", dir=path.parent)
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as handle:
            json.dump(value, handle, indent=2, ensure_ascii=False)
            handle.write("\n")
            handle.flush()
            os.fsync(handle.fileno())
        os.replace(temporary_name, path)
    except Exception:
        try:
            os.unlink(temporary_name)
        except FileNotFoundError:
            pass
        raise

def normalize_url(url):
    if not isinstance(url, str):
        return ""
    url = url.strip()
    if not url:
        return ""
    if url.startswith("//"):
        url = "https:" + url
    if not re.match(r"^https?://", url, re.I):
        return ""
    return url


def is_safe_public_url(url: str) -> bool:
    """Allow only public HTTP(S) URLs; prevent SSRF through redirects and inputs."""
    try:
        parsed = urlparse(url)
        if parsed.scheme.lower() not in {"http", "https"} or not parsed.hostname:
            return False
        if parsed.username or parsed.password:
            return False
        hostname = parsed.hostname.rstrip(".").casefold()
        if hostname in {"localhost", "localhost.localdomain", "ip6-localhost"} or hostname.endswith(".local"):
            return False
        addresses = socket.getaddrinfo(hostname, parsed.port or (443 if parsed.scheme.lower() == "https" else 80), type=socket.SOCK_STREAM)
        if not addresses:
            return False
        for address in addresses:
            ip = ipaddress.ip_address(address[4][0])
            if not ip.is_global:
                return False
        return True
    except (ValueError, socket.gaierror, OSError):
        return False


def safe_get(session, url: str, *, timeout: int = REQUEST_TIMEOUT):
    """Fetch with bounded, validated redirects instead of following arbitrary URLs."""
    current = normalize_url(url)
    for _ in range(MAX_REDIRECTS + 1):
        if not is_safe_public_url(current):
            raise ValueError(f"Blocked non-public or invalid URL: {current}")
        response = session.get(current, timeout=timeout, allow_redirects=False, stream=True)
        if response.is_redirect or response.is_permanent_redirect:
            location = response.headers.get("Location", "")
            response.close()
            current = normalize_url(urljoin(current, location))
            if not current:
                raise ValueError("Redirect did not provide a valid HTTP(S) location")
            continue
        response.url = current
        return response
    raise ValueError(f"Too many redirects for URL: {url}")


def read_response_bytes(response) -> bytes:
    """Read at most MAX_RESPONSE_BYTES to prevent oversized-response exhaustion."""
    try:
        declared = int(response.headers.get("Content-Length", "0") or 0)
    except ValueError:
        declared = 0
    if declared > MAX_RESPONSE_BYTES:
        response.close()
        raise ValueError(f"Response exceeds MAX_RESPONSE_BYTES ({MAX_RESPONSE_BYTES})")
    data = bytearray()
    try:
        for chunk in response.iter_content(chunk_size=65536):
            if not chunk:
                continue
            data.extend(chunk)
            if len(data) > MAX_RESPONSE_BYTES:
                raise ValueError(f"Response exceeds MAX_RESPONSE_BYTES ({MAX_RESPONSE_BYTES})")
        return bytes(data)
    finally:
        response.close()

def load_input(path: Path):
    if not path.exists():
        raise FileNotFoundError(
            f"Input workbook not found: {path}\n"
            "Place the workbook beside main.py or set INPUT_XLSX in .env."
        )
    # Read actual cell hyperlinks, because the uploaded workbook's URL cells display
    # friendly labels while the real URLs are stored as hyperlink targets.
    wb = load_workbook(path, data_only=True)
    records = []
    for ws in wb.worksheets:
        headers = [str(c.value).strip() if c.value is not None else "" for c in ws[1]]
        header_map = {h.lower(): i for i, h in enumerate(headers)}
        company_idx = header_map.get("company")
        url_idx = header_map.get("url")
        desc_idx = header_map.get("description")
        if company_idx is None or url_idx is None:
            logging.warning("Skipping worksheet %s: missing Company or URL columns", ws.title)
            continue
        for row_idx in range(2, ws.max_row + 1):
            company_cell = ws.cell(row_idx, company_idx + 1)
            url_cell = ws.cell(row_idx, url_idx + 1)
            desc_cell = ws.cell(row_idx, desc_idx + 1) if desc_idx is not None else None
            company = str(company_cell.value or "").strip()
            # URL cells contain friendly labels; use the hyperlink target.
            url = normalize_url(url_cell.hyperlink.target if url_cell.hyperlink else url_cell.value)
            description = str(desc_cell.value or "").strip() if desc_cell else ""
            if company and url and is_safe_public_url(url):
                records.append({"company": company, "url": url, "description": description})
            elif company or url:
                logging.warning("Skipping %s row %s: missing company or usable URL", ws.title, row_idx)
    # Deduplicate by company+URL
    unique, seen = [], set()
    for r in records:
        key = (r["company"].casefold(), r["url"].rstrip("/").casefold())
        if key not in seen:
            seen.add(key)
            unique.append(r)
    return unique

def make_session():
    session = requests.Session()
    session.headers.update({"User-Agent": USER_AGENT, "Accept": "text/html,application/xhtml+xml"})
    retry = Retry(total=2, connect=2, read=2, backoff_factor=0.7,
                  status_forcelist=[429, 500, 502, 503, 504],
                  allowed_methods=["GET"], raise_on_status=False)
    adapter = HTTPAdapter(max_retries=retry)
    session.mount("http://", adapter)
    session.mount("https://", adapter)
    return session

def discover_page_links(session, records):
    """Add one bounded layer of relevant same-site links from each seed page."""
    seeds = list(records)
    existing = {(r["company"].casefold(), r["url"].casefold()) for r in records}
    added = 0
    for seed in seeds:
        if added >= MAX_DISCOVERED_URLS:
            break
        try:
            response = safe_get(session, seed["url"])
            content_type = response.headers.get("Content-Type", "").lower()
            body = read_response_bytes(response)
            if response.status_code >= 400 or "html" not in content_type:
                continue
            candidate_links = discover_sitemap_links(session, response.url, MAX_LINKS_PER_SEED, REQUEST_TIMEOUT)
            page_html = body.decode(response.encoding or "utf-8", errors="replace")
            candidate_links += extract_candidate_links(response.url, page_html, MAX_LINKS_PER_SEED, allow_external=True)
            for link in dict.fromkeys(candidate_links):
                if not is_safe_public_url(link):
                    continue
                key = (seed["company"].casefold(), link.casefold())
                if key in existing:
                    continue
                records.append({
                    "company": seed["company"], "url": link,
                    "description": seed.get("description", ""),
                    "discovered_from": "page_link",
                })
                existing.add(key)
                added += 1
                if added >= MAX_DISCOVERED_URLS:
                    break
        except Exception as exc:
            logging.warning("Link discovery failed for %s: %s", seed["url"], exc)
    logging.info("Page-link discovery added %d URL(s).", added)
    return records

def looks_like_pdf(url, content_type, content):
    return (
        "application/pdf" in (content_type or "").lower()
        or urlparse(url).path.lower().endswith(".pdf")
        or (content or b"").lstrip().startswith(b"%PDF-")
    )

def extract_pdf_text(pdf_bytes):
    """
    Extract text from both ordinary PDFs and many scanned PDFs.
    PyMuPDF is the primary extractor. If the PDF has no usable text,
    optional OCR can be attempted with pytesseract + a local Tesseract install.
    """
    try:
        import fitz  # PyMuPDF
    except ImportError as exc:
        raise RuntimeError("PDF support requires PyMuPDF. Install with: pip install pymupdf") from exc

    pages_text = []
    with fitz.open(stream=pdf_bytes, filetype="pdf") as doc:
        for page_number, page in enumerate(doc, start=1):
            try:
                page_text = page.get_text("text", sort=True) or ""
            except Exception:
                page_text = ""
            if page_text.strip():
                pages_text.append(f"\n[PDF page {page_number}]\n{page_text.strip()}")

    text = "\n".join(pages_text).strip()
    if len(text) >= 80:
        return text[:MAX_TEXT_CHARS], "pymupdf"

    # OCR is optional because it requires a separate Tesseract installation.
    try:
        import pytesseract
        from PIL import Image
        import fitz
        ocr_pages = []
        with fitz.open(stream=pdf_bytes, filetype="pdf") as doc:
            for page_number, page in enumerate(doc, start=1):
                pix = page.get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False)
                image = Image.open(BytesIO(pix.tobytes("png")))
                page_text = pytesseract.image_to_string(image) or ""
                if page_text.strip():
                    ocr_pages.append(f"\n[PDF page {page_number} - OCR]\n{page_text.strip()}")
        ocr_text = "\n".join(ocr_pages).strip()
        if len(ocr_text) > len(text):
            return ocr_text[:MAX_TEXT_CHARS], "ocr"
    except Exception:
        pass

    if text:
        return text[:MAX_TEXT_CHARS], "pymupdf_low_text"
    return "", "no_extractable_text; scanned PDF may require OCR"

def extract_with_requests(session, url):
    response = safe_get(session, url)
    body = read_response_bytes(response)
    if response.status_code >= 400:
        raise requests.HTTPError(f"HTTP {response.status_code} for {response.url}")
    content_type = response.headers.get("Content-Type", "").lower()
    if looks_like_pdf(response.url, content_type, body):
        pdf_text, pdf_method = extract_pdf_text(body)
        pdf_note = "" if pdf_text.strip() else "No text extracted from PDF; scanned PDF may need OCR or manual review"
        return response.url, Path(urlparse(response.url).path).name or "PDF document", pdf_text, response.status_code, pdf_note, pdf_method
    page_html = body.decode(response.encoding or "utf-8", errors="replace")
    soup = BeautifulSoup(page_html, "html.parser")
    for tag in soup(["script", "style", "noscript", "svg", "nav", "footer", "header"]):
        tag.decompose()
    title = soup.title.get_text(" ", strip=True) if soup.title else ""
    text = re.sub(r"\s+", " ", soup.get_text(" ", strip=True))[:MAX_TEXT_CHARS]
    return response.url, title, text, response.status_code, "", "requests"

def extract_with_playwright(url):
    if sync_playwright is None:
        raise RuntimeError("Playwright is not installed.")
    if not is_safe_public_url(url):
        raise ValueError(f"Blocked non-public or invalid URL: {url}")
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context_options = {"user_agent": USER_AGENT}
        if PLAYWRIGHT_STORAGE_STATE and Path(PLAYWRIGHT_STORAGE_STATE).exists():
            context_options["storage_state"] = PLAYWRIGHT_STORAGE_STATE
        context = browser.new_context(**context_options)
        def route_handler(route):
            if is_safe_public_url(route.request.url):
                route.continue_()
            else:
                route.abort()
        context.route("**/*", route_handler)
        page = context.new_page()
        try:
            page.goto(url, wait_until="domcontentloaded", timeout=45000)
            page.wait_for_timeout(1800)
            final_url = page.url
            if not is_safe_public_url(final_url):
                raise ValueError(f"Blocked redirect to non-public URL: {final_url}")
            title = page.title()
            text = page.locator("body").inner_text(timeout=10000)
            return final_url, title, re.sub(r"\s+", " ", text)[:MAX_TEXT_CHARS], 200, "", "playwright"
        finally:
            context.close()
            browser.close()

def resolve_application_url_from_page(session, source_url, final_url, page_text=""):
    """Resolve the direct RFP/grant document link from a public landing page."""
    if is_safe_public_url(final_url) and urlparse(final_url).path.casefold().endswith((".pdf", ".doc", ".docx", ".xls", ".xlsx")):
        return final_url
    response = None
    try:
        response = safe_get(session, final_url or source_url)
        body = read_response_bytes(response)
        if response.status_code >= 400:
            raise requests.HTTPError(f"HTTP {response.status_code} for {response.url}")
        deterministic = resolve_application_url(response.url, body.decode(response.encoding or "utf-8", errors="replace"))
        if deterministic and is_safe_public_url(deterministic):
            return deterministic
    except Exception as exc:
        logging.warning("Could not resolve application link from %s: %s", source_url, exc)

    # A number of company grant pages block requests but expose the RFP link in
    # the rendered DOM. Use the existing Playwright fallback when available.
    if sync_playwright is not None:
        try:
            with sync_playwright() as p:
                browser = p.chromium.launch(headless=True)
                context_options = {"user_agent": USER_AGENT}
                if PLAYWRIGHT_STORAGE_STATE and Path(PLAYWRIGHT_STORAGE_STATE).exists():
                    context_options["storage_state"] = PLAYWRIGHT_STORAGE_STATE
                context = browser.new_context(**context_options)
                def route_handler(route):
                    if is_safe_public_url(route.request.url):
                        route.continue_()
                    else:
                        route.abort()
                page = context.new_page()
                try:
                    if not is_safe_public_url(source_url):
                        raise ValueError(f"Blocked non-public or invalid URL: {source_url}")
                    context.route("**/*", route_handler)
                    page.goto(source_url, wait_until="domcontentloaded", timeout=45000)
                    page.wait_for_timeout(1800)
                    rendered_url = page.url
                    if not is_safe_public_url(rendered_url):
                        raise ValueError(f"Blocked redirect to non-public URL: {rendered_url}")
                    rendered_html = page.content()
                    deterministic = resolve_application_url(rendered_url, rendered_html)
                    if deterministic and is_safe_public_url(deterministic):
                        return deterministic
                finally:
                    context.close()
                    browser.close()
        except Exception as exc:
            logging.warning("Rendered link resolution failed for %s: %s", source_url, exc)

    return ""

def find_terms(text, terms):
    low = text.casefold()
    return [term for term in terms if re.search(r"(?<!\w)" + re.escape(term.casefold()) + r"(?!\w)", low)]

def classify(company, url, description, text):
    combined = " ".join([company, url, description, text])
    funding = find_terms(combined, FUNDING_TERMS)
    topics = find_terms(combined, TOPIC_TERMS)
    target_topics = find_terms(combined, TARGET_TOPIC_TERMS)
    rfp_terms = find_terms(combined, RFP_TERMS)
    open_signals = find_terms(combined, OPEN_TERMS)
    closed_signals = find_terms(combined, CLOSED_TERMS)
    target_rfp = bool(funding and target_topics and rfp_terms)
    return {
        "funding_terms": funding,
        "topic_terms": topics,
        "target_topics": target_topics,
        "rfp_terms": rfp_terms,
        "open_signals": open_signals,
        "closed_signals": closed_signals,
        "is_target_rfp": target_rfp,
        "flag": target_rfp,
        "priority": "HIGH" if target_rfp and open_signals and not closed_signals
                    else "MEDIUM" if target_rfp
                    else "NONE",
    }

def opportunity_fingerprint(final_url, application_url, title, deadline_text, evidence_quote, target_topics):
    """Fingerprint opportunity fields, excluding volatile page-wide navigation text."""
    payload = json.dumps({
        "final_url": final_url or "",
        "application_url": application_url or "",
        "title": re.sub(r"\s+", " ", title or "").strip(),
        "deadline": re.sub(r"\s+", " ", deadline_text or "").strip(),
        "evidence": re.sub(r"\s+", " ", evidence_quote or "").strip(),
        "target_topics": sorted(target_topics or []),
    }, sort_keys=True, ensure_ascii=False)
    return hashlib.sha256(payload.encode("utf-8")).hexdigest()

def is_target_rfp_record(record, update=None):
    """Require both target-topic evidence and explicit RFP/proposal language."""
    update = update or {}
    text = " ".join(str(record.get(field, "")) for field in (
        "company", "page_title", "opportunity_title", "evidence", "evidence_quote",
        "focus", "therapeutic_areas", "funding_terms",
    )) + " " + " ".join(str(update.get(field, "")) for field in (
        "opportunity_title", "evidence", "evidence_quote", "focus", "therapeutic_areas", "funding_terms",
    ))
    return bool(find_terms(text, TARGET_TOPIC_TERMS) and find_terms(text, RFP_TERMS))

def opportunity_evidence(text):
    """Return concise funding-related sentences used for positive-change detection."""
    sentences = re.split(r"(?<=[.!?])\s+|\s*[•▪\n]+\s*", text)
    evidence = []
    seen = set()
    for sentence in sentences:
        clean = re.sub(r"\s+", " ", sentence).strip(" -:;•")
        low = clean.casefold()
        has_funding = any(term.casefold() in low for term in FUNDING_TERMS)
        has_positive = any(term.casefold() in low for term in OPEN_TERMS)
        has_rfp = any(term.casefold() in low for term in RFP_TERMS)
        has_opportunity = any(term in low for term in ("opportunity", "accepting", "applications are open", "apply now", "submit by", "deadline"))
        if clean and has_funding and (has_positive or (has_rfp and has_opportunity)):
            key = low
            if key not in seen:
                seen.add(key)
                evidence.append(clean[:700])
    return evidence[:100]


def extract_deadline_text(text: str) -> str:
    """Extract a source sentence containing a deadline or due date."""
    normalized_text = re.sub(
        r"\b(January|February|March|April|May|June|July|August|September|October|November|December|"
        r"Jan|Feb|Mar|Apr|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec)\s+(\d{1,2})\s*[,;:/-]+\s*(\d{4})\b",
        r"\1 \2, \3",
        str(text or ""),
        flags=re.I,
    )
    sentences = re.split(r"(?<=[.!?])\s+|\s*[\n]+\s*", normalized_text)
    date_pattern = re.compile(
        r"\b(?:\d{1,2}[/-]\d{1,2}[/-]\d{2,4}|\d{4}[/-]\d{1,2}[/-]\d{1,2}|"
        r"(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\.?\s+\d{1,2}(?:st|nd|rd|th)?\s*[,;]?\s+\d{4}|"
        r"\d{1,2}(?:st|nd|rd|th)?\s+(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\.?\s+\d{4})\b",
        re.I,
    )
    for sentence in sentences:
        clean = re.sub(r"\s+", " ", sentence).strip(" -:;")
        if re.search(r"\b(?:deadline|due date|application deadline|submit by|submission deadline)\b", clean, re.I) and date_pattern.search(clean):
            return clean[:700]
    return ""

def email_categories(text):
    """Return only the supported RFP, CGA, and IME categories for email content."""
    return [category for category, pattern in EMAIL_CATEGORY_PATTERNS.items() if pattern.search(text or "")]

def email_evidence_snippets(evidence_items):
    """Split page evidence into shorter excerpts and retain only RFP/CGA/IME text."""
    snippets = []
    seen = set()
    for item in evidence_items:
        # The source page is flattened before extraction, so split on common
        # sentence, list, and table separators to avoid emailing whole sections.
        normalized = re.sub(r"\\\\(?=[*_])", "", str(item or ""))
        # Gemini and browser extraction can flatten HTML tables into one line.
        # Split common table labels so one alert does not become one unreadable
        # paragraph containing several unrelated fields.
        parts = re.split(
            r"(?<=[.!?])\s+|\s*[;|\u2022\u25aa]\s*|\s+(?=\d{4}\b)|"
            r"\s+(?=(?:RFP Code|RFP Issue Date|Area of Interest|Type of Request|"
            r"RFP Title|LOR Due Date|Due Date|Deadline|Important Notice)\b)",
            normalized,
            flags=re.I,
        )
        for part in parts:
            clean = re.sub(r"\s+", " ", part).strip(" -:;\u2022\u25aa")
            key = clean.casefold()
            table_context = re.match(
                r"(?:RFP Code|RFP Issue Date|Area of Interest|Type of Request|RFP Title|"
                r"LOR Due Date|Due Date|Deadline|Important Notice)\b",
                clean,
                flags=re.I,
            )
            if clean and (email_categories(clean) or table_context) and key not in seen:
                seen.add(key)
                snippets.append(clean[:500])
    return snippets

def _email_specialty(alert, evidence):
    """Choose specialty from verified fields, then alert-specific evidence."""
    values = alert.get("therapeutic_areas") or alert.get("focus")
    if isinstance(values, str):
        values = re.split(r"\s*;\s*|\s*,\s*", values)
    values = [re.sub(r"\s+", " ", str(value)).strip() for value in (values or [])]
    values = [value for value in values if value]
    if not values:
        values = find_terms(" ".join(evidence), TOPIC_TERMS)
    if not values:
        values = alert.get("topic_terms", [])
    return "; ".join(dict.fromkeys(values)) or "Not identified"

def _parse_deadline_date(value: str):
    """Parse common deadline formats without treating ambiguous text as a date."""
    value = re.sub(r"\s+", " ", str(value or "")).strip()
    value = re.sub(
        r"\b(January|February|March|April|May|June|July|August|September|October|November|December|"
        r"Jan|Feb|Mar|Apr|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec)\s+(\d{1,2})\s*[,;:/-]+\s*(\d{4})\b",
        r"\1 \2, \3",
        value,
        flags=re.I,
    )
    value = re.sub(r"(\d{1,2})(st|nd|rd|th)\b", r"\1", value, flags=re.I)
    if not value:
        return None
    candidates = [value]
    candidates.extend(re.findall(
        r"\b(?:\d{1,2}[/-]\d{1,2}[/-]\d{2,4}|\d{4}[/-]\d{1,2}[/-]\d{1,2}|"
        r"(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec)[a-z]*\.?\s+\d{1,2}\s*[,;]?\s+\d{4}|"
        r"\d{1,2}\s+(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\.?\s+\d{4})\b",
        value, flags=re.I,
    ))
    formats = (
        "%Y-%m-%d", "%Y/%m/%d", "%d/%m/%Y", "%m/%d/%Y", "%d-%m-%Y", "%m-%d-%Y",
        "%B %d, %Y", "%B %d %Y", "%b %d, %Y", "%b %d %Y", "%d %B %Y", "%d %b %Y",
    )
    from datetime import datetime
    for candidate in candidates:
        candidate = candidate.strip().replace(".", "")
        for fmt in formats:
            try:
                return datetime.strptime(candidate, fmt).date()
            except ValueError:
                continue
    return None

def deadline_is_expired(alert, today=None):
    """Return True when a stated deadline was yesterday or earlier."""
    today = today or datetime.now(MONITOR_TIMEZONE).date()
    deadline = _parse_deadline_date(alert.get("deadline_date", ""))
    if deadline is None:
        deadline = _parse_deadline_date(alert.get("deadline_text", ""))
    return deadline is not None and deadline <= today - timedelta(days=1)

def deadline_days_remaining(alert, today=None):
    """Return calendar days remaining, or None when no reliable deadline exists."""
    today = today or datetime.now(MONITOR_TIMEZONE).date()
    deadline = _parse_deadline_date(alert.get("deadline_date", ""))
    if deadline is None:
        deadline = _parse_deadline_date(alert.get("deadline_text", ""))
    return (deadline - today).days if deadline else None

def deadline_is_too_soon(alert, today=None):
    """Block alerts expired or within the configured minimum lead time."""
    days = deadline_days_remaining(alert, today)
    return days is not None and days <= ALERT_MIN_LEAD_DAYS

def first_email(text: str) -> str:
    match = re.search(r"\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b", text or "", re.I)
    return match.group(0) if match else ""

EMAIL_ALERT_COLUMNS = [
    "Pharma Company Name", "RFP URL", "Evidence Text", "Keywords", "Title of the RFP", "Specialty",
]

def build_email_alert_rows(alerts):
    """Build the alert records shown in emails and the alert spreadsheet tabs."""
    rows = []
    ordered_alerts = sorted(
        alerts,
        key=lambda alert: (0 if alert["opportunity_type"] == "RFP" else 1, alert["company"].casefold()),
    )
    for alert in ordered_alerts:
        email_url = alert.get("application_url") or alert.get("final_url") or alert.get("source_url")
        if (
            alert.get("email_eligible") != "Yes"
            or deadline_is_too_soon(alert)
            or alert.get("email_block_reason")
            or email_url_block_reason(email_url)
            or alert.get("access_issue")
        ):
            continue
        evidence_source = [alert["evidence_quote"]] if alert.get("evidence_quote") else alert.get("new_evidence", [])
        evidence = email_evidence_snippets(evidence_source)
        if not evidence:
            continue
        keywords = list(dict.fromkeys(category for item in evidence for category in email_categories(item)))
        rows.append([
            alert["company"],
            email_url,
            "; ".join(evidence),
            "; ".join(keywords),
            alert["opportunity_title"],
            _email_specialty(alert, evidence),
        ])
    return rows

def get_state():
    if STATE_FILE.exists():
        try:
            return json.loads(STATE_FILE.read_text(encoding="utf-8"))
        except Exception:
            logging.warning("State file is invalid; starting with empty state.")
    return {}

def save_state(state):
    atomic_write_json(STATE_FILE, state)

def append_daily_run_log(rows):
    """Append findings to the CSV for the monitor's local calendar day."""
    if not rows:
        return
    run_date = str(rows[0].get("run_date", ""))
    try:
        day = datetime.strptime(run_date, "%d/%m/%y")
    except ValueError:
        day = datetime.now(MONITOR_TIMEZONE)
    # Slashes are invalid in Windows filenames; the CSV field remains DD/MM/YY.
    daily_log = DAILY_RUN_LOG_DIR / f"{day.strftime('%d-%m-%y')}.csv"
    DAILY_RUN_LOG_DIR.mkdir(parents=True, exist_ok=True)
    exists = daily_log.exists() and daily_log.stat().st_size > 0
    with daily_log.open("a", newline="", encoding="utf-8-sig") as handle:
        writer = csv.DictWriter(handle, fieldnames=DAILY_RUN_LOG_COLUMNS, extrasaction="ignore")
        if not exists:
            writer.writeheader()
        for row in rows:
            writer.writerow({field: row.get(field, "") for field in DAILY_RUN_LOG_COLUMNS})
        handle.flush()
        os.fsync(handle.fileno())
    return daily_log

def build_daily_run_log_rows(results, alerts, *, args, email_result, email_error, email_sent):
    """Explain the finding and final alert decision for every checked record."""
    local_now = datetime.now(MONITOR_TIMEZONE)
    run_date = local_now.strftime("%d/%m/%y")
    checked_at = local_now.strftime("%d/%m/%y %H:%M:%S")
    sent_at = checked_at if email_sent else ""
    alert_keys = {
        (str(alert.get("company", "")).casefold(), str(alert.get("source_url", "")).casefold())
        for alert in alerts
    }
    email_configured = bool(
        os.getenv("SMTP_HOST", "").strip()
        and os.getenv("ALERT_FROM", "").strip()
        and os.getenv("ALERT_TO", "").strip()
    )
    rows = []
    for result in results:
        key = (str(result.get("company", "")).casefold(), str(result.get("source_url", "")).casefold())
        if result.get("error") or int(result.get("http_status") or 0) >= 400:
            finding_status = "FAILED"
            finding_reason = result.get("error") or f"HTTP status {result.get('http_status')}"
        elif result.get("is_target_rfp"):
            finding_status = "TARGET_RFP"
            finding_reason = "Verified oncology/neurology RFP or proposal finding"
        elif result.get("flag"):
            finding_status = "FUNDING_FINDING"
            finding_reason = "Funding-related finding did not meet target RFP criteria"
        else:
            finding_status = "NO_TARGET_FINDING"
            finding_reason = "No verified target oncology/neurology RFP finding"

        if key in alert_keys and email_sent:
            alert_status, alert_reason = "SENT", "Email delivered successfully"
        elif args.baseline:
            alert_status, alert_reason = "NOT_SENT", "Baseline mode"
        elif result.get("error") or int(result.get("http_status") or 0) >= 400:
            alert_status, alert_reason = "NOT_SENT", "Page check failed"
        elif not result.get("is_target_rfp"):
            alert_status, alert_reason = "NOT_SENT", "Not a verified target RFP"
        elif result.get("email_eligible") != "Yes":
            alert_status, alert_reason = "NOT_SENT", result.get("email_block_reason") or "Email eligibility failed"
        elif deadline_is_too_soon(result):
            alert_status, alert_reason = "NOT_SENT", "Expired or within minimum lead time"
        elif email_url_block_reason(result.get("application_url") or result.get("final_url") or result.get("source_url")):
            alert_status, alert_reason = "NOT_SENT", "Application URL blocked"
        elif result.get("access_issue"):
            alert_status, alert_reason = "NOT_SENT", "Source access issue"
        elif result.get("change_type") == "UNCHANGED" and not args.force_alert:
            alert_status, alert_reason = "NOT_SENT", "No new or changed evidence"
        elif key in alert_keys and email_error:
            alert_status, alert_reason = "NOT_SENT", f"Email delivery failed: {email_error}"
        elif key in alert_keys and email_result is None:
            alert_status, alert_reason = "NOT_SENT", "Email not configured"
        else:
            alert_status, alert_reason = "NOT_SENT", "Did not pass final alert filters"

        rows.append({
            "run_date": run_date, "run_id": RUN_ID, "checked_at": checked_at,
            "company": result.get("company", ""), "source_url": result.get("source_url", ""),
            "final_url": result.get("final_url", ""), "page_title": result.get("page_title", ""),
            "opportunity_type": result.get("opportunity_type", ""),
            "opportunity_title": result.get("opportunity_title", ""),
            "opportunity_status": result.get("opportunity_status", ""),
            "target_topics": result.get("target_topics", ""),
            "deadline_date": result.get("deadline_date", ""),
            "deadline_text": result.get("deadline_text", ""),
            "change_type": result.get("change_type", ""),
            "finding_status": finding_status, "finding_reason": finding_reason,
            "alert_status": alert_status, "alert_reason": alert_reason,
            "alert_sent_at": sent_at if key in alert_keys and email_sent else "",
            "email_configured": "Yes" if email_configured else "No",
            "email_attempted": "Yes" if email_result is not None else "No",
            "http_status": result.get("http_status", ""),
            "extraction_method": result.get("extraction_method", ""),
            "evidence": result.get("evidence", ""), "error": result.get("error", ""),
        })
    return rows

def sync_google_sheet(csv_path):
    """Replace the configured worksheet contents with the cumulative matches CSV."""
    spreadsheet_id = os.getenv("GOOGLE_SHEETS_SPREADSHEET_ID", "").strip()
    worksheet_name = os.getenv("GOOGLE_SHEETS_WORKSHEET", "").strip()
    credentials_setting = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON", "").strip()
    if not (spreadsheet_id and worksheet_name and credentials_setting):
        logging.info("Google Sheets sync is not configured; skipping spreadsheet update.")
        return False
    # Accept either the spreadsheet ID or its full Google Sheets URL in .env.
    spreadsheet_match = re.search(r"/spreadsheets/d/([A-Za-z0-9_-]+)", spreadsheet_id)
    spreadsheet_key = spreadsheet_match.group(1) if spreadsheet_match else spreadsheet_id

    try:
        import gspread
        from google.oauth2.service_account import Credentials
    except ImportError as exc:
        raise RuntimeError("Google Sheets sync requires gspread and google-auth; install requirements.txt.") from exc

    credentials_path = Path(credentials_setting)
    if not credentials_path.is_absolute():
        credentials_path = ROOT / credentials_path
    credentials = Credentials.from_service_account_file(
        str(credentials_path),
        scopes=["https://www.googleapis.com/auth/spreadsheets"],
    )
    with open(csv_path, "r", encoding="utf-8-sig", newline="") as csv_file:
        values = list(csv.reader(csv_file))
    client = gspread.authorize(credentials)
    spreadsheet = client.open_by_key(spreadsheet_key)
    logging.info(
        "Google Sheets target: spreadsheet='%s' | URL=%s | configured worksheet='%s' | CSV=%s",
        spreadsheet.title, spreadsheet.url, worksheet_name, csv_path.resolve(),
    )
    try:
        worksheet = spreadsheet.worksheet(worksheet_name)
    except gspread.WorksheetNotFound:
        row_count = max(100, len(values))
        column_count = max(1, max((len(row) for row in values), default=1))
        worksheet = spreadsheet.add_worksheet(
            title=worksheet_name,
            rows=str(row_count),
            cols=str(column_count),
        )
        logging.info("Created missing Google Sheets worksheet '%s'.", worksheet_name)

    # Replace the tab contents with the CSV, then read the tab back to verify.
    worksheet.clear()
    if values:
        worksheet.update(range_name="A1", values=values)
    expected_rows = len(values)
    expected_cols = max((len(row) for row in values), default=0)
    written_values = worksheet.get_all_values()
    written_rows = len(written_values)
    logging.info(
        "Google Sheets write verification: tab_url=%s | expected_rows_including_header=%d | readback_rows=%d | expected_columns=%d",
        worksheet.url, expected_rows, written_rows, expected_cols,
    )
    # Sheets omits trailing blank cells when reading rows back, so normalize row
    # widths before comparing the data to avoid false mismatches.
    def pad_rows(rows):
        return [list(row) + [""] * max(0, expected_cols - len(row)) for row in rows]

    if pad_rows(written_values) != pad_rows(values):
        logging.warning(
            "Google Sheets read-back differs from CSV. CSV header=%r; sheet header=%r",
            values[0] if values else [], written_values[0] if written_values else [],
        )
        raise RuntimeError("Google Sheets write completed but read-back did not match the CSV.")
    logging.info(
        "Google Sheet updated and verified: spreadsheet='%s' | worksheet='%s' | %d data row(s).",
        spreadsheet.title, worksheet_name, max(0, len(values) - 1),
    )
    return True

def sync_email_alerts_to_google_sheet(alerts):
    """Replace the current-alert tab and maintain a separate 30-day history tab."""
    current_rows = build_email_alert_rows(alerts)
    sent_at = datetime.now(UTC).isoformat(timespec="seconds")
    history_headers = EMAIL_ALERT_COLUMNS + ["Alert Sent At (UTC)"]
    old_email_headers = EMAIL_ALERT_COLUMNS[:-1]
    old_url_email_headers = ["Pharma Company Name", "URL", "Evidence Text", "Keywords", "Title of the RFP"]
    old_history_headers = [old_email_headers + ["Alert Sent At (IST)"], old_url_email_headers + ["Alert Sent At (IST)"],
                           old_email_headers + ["Alert Sent At (UTC)"], old_url_email_headers + ["Alert Sent At (UTC)"]]
    new_history_rows = [row + [sent_at] for row in current_rows]

    spreadsheet_id = os.getenv("GOOGLE_SHEETS_SPREADSHEET_ID", "").strip()
    credentials_setting = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON", "").strip()
    current_name = os.getenv("GOOGLE_SHEETS_EMAIL_WORKSHEET", "Email Alerts").strip() or "Email Alerts"
    history_name = os.getenv("GOOGLE_SHEETS_EMAIL_HISTORY_WORKSHEET", "Email Alerts History").strip() or "Email Alerts History"
    if not (spreadsheet_id and credentials_setting):
        logging.info("Google Sheets email-alert sync is not configured; skipping.")
        return False
    if current_name == history_name:
        raise ValueError("Current email alerts and email history must use different worksheet names.")
    spreadsheet_match = re.search(r"/spreadsheets/d/([A-Za-z0-9_-]+)", spreadsheet_id)
    spreadsheet_key = spreadsheet_match.group(1) if spreadsheet_match else spreadsheet_id

    try:
        import gspread
        from google.oauth2.service_account import Credentials
    except ImportError as exc:
        raise RuntimeError("Google Sheets sync requires gspread and google-auth; install requirements.txt.") from exc

    credentials_path = Path(credentials_setting)
    if not credentials_path.is_absolute():
        credentials_path = ROOT / credentials_path
    credentials = Credentials.from_service_account_file(
        str(credentials_path), scopes=["https://www.googleapis.com/auth/spreadsheets"],
    )
    spreadsheet = gspread.authorize(credentials).open_by_key(spreadsheet_key)

    def get_or_create_tab(name, columns):
        try:
            return spreadsheet.worksheet(name)
        except gspread.WorksheetNotFound:
            tab = spreadsheet.add_worksheet(title=name, rows="1000", cols=str(columns))
            logging.info("Created Google Sheets worksheet '%s'.", name)
            return tab

    current_tab = get_or_create_tab(current_name, len(EMAIL_ALERT_COLUMNS))
    history_tab = get_or_create_tab(history_name, len(history_headers))
    current_values = current_tab.get_all_values()
    history_values = history_tab.get_all_values()

    def records_for_tab(values, name):
        values = [row for row in values if any(str(cell).strip() for cell in row)]
        if not values:
            return [], []
        if values[0] == history_headers:
            rows = [row[:len(history_headers)] + [""] * max(0, len(history_headers) - len(row)) for row in values[1:]]
            return [normalize_history_timestamp(row) for row in rows], []
        if values[0] in old_history_headers:
            rows = []
            for old_row in values[1:]:
                old_row = old_row + [""] * max(0, 6 - len(old_row))
                row = old_row[:5] + ["Not identified"] + [old_row[5]]
                rows.append(normalize_history_timestamp(row))
            return rows, []
        if values[0] == EMAIL_ALERT_COLUMNS:
            return [], [row[:len(EMAIL_ALERT_COLUMNS)] + [""] * max(0, len(EMAIL_ALERT_COLUMNS) - len(row)) for row in values[1:]]
        if values[0] in (old_email_headers, old_url_email_headers):
            return [], [row[:5] + ["Not identified"] for row in values[1:]]
        raise RuntimeError(f"Worksheet '{name}' has unexpected headers; leaving it unchanged.")

    def normalize_history_timestamp(row):
        """Convert stored timestamp instants to a consistent UTC ISO string."""
        try:
            timestamp = datetime.fromisoformat(row[-1])
            if timestamp.tzinfo is None:
                timestamp = timestamp.replace(tzinfo=UTC)
            row[-1] = timestamp.astimezone(UTC).isoformat(timespec="seconds")
        except ValueError:
            pass  # The retention pass below will discard malformed timestamps.
        return row

    history_rows, legacy_history_rows = records_for_tab(history_values, history_name)
    current_history_rows, legacy_current_rows = records_for_tab(current_values, current_name)
    known_timestamped_payloads = {
        tuple(row[:len(EMAIL_ALERT_COLUMNS)]) for row in history_rows + current_history_rows
    }
    legacy_rows = legacy_history_rows + [
        row for row in legacy_current_rows
        if tuple(row) not in known_timestamped_payloads
    ]

    # Preserve pre-timestamp data in a clearly labelled archive before migrating tabs.
    if legacy_rows:
        archive_headers = EMAIL_ALERT_COLUMNS + ["Timestamp Status"]
        archive = None
        archive_values = []
        for suffix in range(1, 100):
            archive_name = "Email Alerts Legacy Archive" if suffix == 1 else f"Email Alerts Legacy Archive {suffix}"
            try:
                candidate = spreadsheet.worksheet(archive_name)
                candidate_values = candidate.get_all_values()
            except gspread.WorksheetNotFound:
                candidate = spreadsheet.add_worksheet(title=archive_name, rows="1000", cols=str(len(archive_headers)))
                candidate_values = []
            if not candidate_values or candidate_values[0] == archive_headers:
                archive, archive_values = candidate, candidate_values
                break
            logging.warning("Leaving archive worksheet '%s' unchanged because its headers differ.", archive_name)
        if archive is None:
            raise RuntimeError("Could not find or create a compatible legacy email-alert archive worksheet.")
        if not archive_values:
            archive.update(range_name="A1", values=[archive_headers], value_input_option="RAW")
            archive.format("A1:G1", {"textFormat": {"bold": True}})
            archive_values = [archive_headers]
        known_legacy = {tuple(row[:len(EMAIL_ALERT_COLUMNS)]) for row in archive_values[1:]}
        archive_rows = [
            row + ["Timestamp unavailable (pre-change)"]
            for row in legacy_rows
            if tuple(row) not in known_legacy
        ]
        if archive_rows:
            archive.append_rows(archive_rows, value_input_option="RAW")
        logging.info("Archived %d pre-timestamp alert row(s).", len(archive_rows))

    # Existing current-tab history is migrated to the history tab. Identical
    # rows are deduplicated because it may already have been copied on a prior run.
    candidates = history_rows + current_history_rows + new_history_rows
    cutoff = datetime.now(UTC) - timedelta(days=30)
    retained_rows, expired_rows, seen = [], 0, set()
    for row in candidates:
        key = tuple(row)
        if key in seen:
            continue
        seen.add(key)
        try:
            timestamp = datetime.fromisoformat(row[-1])
            if timestamp.tzinfo is None:
                timestamp = timestamp.replace(tzinfo=UTC)
        except ValueError:
            expired_rows += 1
            continue
        if timestamp.astimezone(UTC) >= cutoff:
            retained_rows.append(row)
        else:
            expired_rows += 1

    updated_history = [history_headers] + retained_rows
    history_tab.clear()
    history_tab.update(range_name="A1", values=updated_history, value_input_option="RAW")
    history_tab.format("A1:G1", {"textFormat": {"bold": True}})
    history_tab.freeze(rows=1)
    if history_tab.get_all_values() != updated_history:
        raise RuntimeError(f"Google Sheets read-back did not match '{history_name}'.")

    # This tab represents only emails successfully sent in the current scrape.
    updated_current = [EMAIL_ALERT_COLUMNS] + (current_rows if alerts else [])
    current_tab.clear()
    current_tab.update(range_name="A1", values=updated_current, value_input_option="RAW")
    current_tab.format("A1:F1", {"textFormat": {"bold": True}})
    current_tab.freeze(rows=1)
    if current_tab.get_all_values() != updated_current:
        raise RuntimeError(f"Google Sheets read-back did not match current alerts in '{current_name}'.")

    logging.info(
        "Google Sheets email tabs updated: current='%s' (%d alert(s)); history='%s' (%d retained, %d expired); window=30 days.",
        current_name, len(current_rows) if alerts else 0, history_name, len(retained_rows), expired_rows,
    )
    return True

def send_email_alert(alerts):
    host = os.getenv("SMTP_HOST", "").strip()
    recipients = [r.strip() for r in re.split(r"[,;]", os.getenv("ALERT_TO", "")) if r.strip()]
    sender = os.getenv("ALERT_FROM", "").strip()
    if not (host and recipients and sender):
        logging.info("Email not configured; alerts will be written to CSV/log only.")
        return None
    port = int(os.getenv("SMTP_PORT", "587"))
    username = os.getenv("SMTP_USERNAME", "")
    password = os.getenv("SMTP_PASSWORD", "")
    use_tls = os.getenv("SMTP_USE_TLS", "true").lower() in {"1", "true", "yes"}
    if host.casefold() not in {"localhost", "127.0.0.1", "::1"} and not use_tls:
        raise RuntimeError("SMTP_USE_TLS must be enabled for non-local SMTP servers")

    msg = EmailMessage()
    msg["Subject"] = "Pharma RFP Monitor Alert"
    msg["From"] = sender
    msg["To"] = ", ".join(recipients)
    text_blocks, html_blocks = [], []
    for row in build_email_alert_rows(alerts):
        fields = list(zip(EMAIL_ALERT_COLUMNS, row))
        text_lines = [f"{label}: {value}" for label, value in fields]
        text_blocks.append("\n".join(text_lines))
        html_rows = []
        for label, value in fields:
            if label == "RFP URL":
                rendered_value = f'<a href="{html.escape(value, quote=True)}">{html.escape(value)}</a>'
            elif label == "Evidence Text":
                evidence_lines = [html.escape(part.strip()) for part in value.split("; ") if part.strip()]
                rendered_value = "<ul>" + "".join(f"<li>{part}</li>" for part in evidence_lines) + "</ul>"
            else:
                rendered_value = html.escape(value)
            html_rows.append(
                f"<tr><th align=\"left\" valign=\"top\">{html.escape(label)}</th>"
                f"<td valign=\"top\">{rendered_value}</td></tr>"
            )
        html_blocks.append(
            '<table cellpadding="6" cellspacing="0" border="1" style="border-collapse:collapse; margin:0 0 18px 0;">'
            + "".join(html_rows)
            + "</table>"
        )
    if not html_blocks:
        logging.info("No RFP, CGA, or IME evidence remains after email filtering.")
        return None
    spreadsheet_setting = os.getenv("GOOGLE_SHEETS_SPREADSHEET_ID", "").strip()
    spreadsheet_match = re.search(r"/spreadsheets/d/([A-Za-z0-9_-]+)", spreadsheet_setting)
    spreadsheet_id = spreadsheet_match.group(1) if spreadsheet_match else spreadsheet_setting
    spreadsheet_url = f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/edit" if spreadsheet_id else ""
    text_body = "\n\n".join(text_blocks)
    html_body = "<hr>".join(html_blocks)
    if spreadsheet_url:
        text_body += f"\n\nGoogle Spreadsheet: {spreadsheet_url}"
        html_body += (
            "<hr><p><strong>Google Spreadsheet:</strong> "
            f'<a href="{html.escape(spreadsheet_url, quote=True)}">Open spreadsheet</a></p>'
        )
    msg.set_content(text_body)
    msg.add_alternative("<html><body>" + html_body + "</body></html>", subtype="html")
    with smtplib.SMTP(host, port, timeout=25) as server:
        if use_tls:
            server.starttls()
        if username:
            server.login(username, password)
        server.send_message(msg, to_addrs=recipients)
    logging.info("Email alert sent to %s recipient(s)", len(recipients))
    return True

def main():
    parser = argparse.ArgumentParser(description="Monitor pharma company URLs for grant/RFP-related content.")
    parser.add_argument("--input", default=str(INPUT_XLSX), help="Excel workbook path")
    parser.add_argument("--max-urls", type=int, default=0, help="Test with only the first N URLs")
    parser.add_argument("--force-alert", action="store_true", help="Email every current match, even if unchanged")
    parser.add_argument("--baseline", action="store_true", help="Save current state without sending alerts (recommended before enabling notifications)")
    parser.add_argument("--discover-gemini", action="store_true", help="Discover additional public pharma/biotech funding pages with Gemini Search grounding")
    args = parser.parse_args()

    setup_logging()
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    records = load_input(Path(args.input))
    gemini = None
    discovery_requested = args.discover_gemini or GEMINI_DISCOVERY
    review_requested = GEMINI_REVIEW
    if discovery_requested and not os.getenv("GEMINI_API_KEY", "").strip():
        raise SystemExit("GEMINI_DISCOVERY requires GEMINI_API_KEY to be configured.")
    if discovery_requested or review_requested:
        try:
            gemini = GeminiEnricher()
            logging.info("Gemini API key loaded; existing URLs will be reviewed before grounded discovery.")
        except Exception:
            if discovery_requested:
                raise
            logging.exception("Gemini setup failed; continuing without Gemini review or grounded discovery.")
            gemini = None
    if args.max_urls:
        records = records[:args.max_urls]
    if not records:
        raise SystemExit("No usable URL hyperlinks found in the workbook.")
    logging.info("Loaded %d company URL(s)", len(records))

    session = make_session()
    records = discover_page_links(session, records)
    state = get_state()
    now = datetime.now(timezone.utc).astimezone().isoformat(timespec="seconds")
    results, alerts = [], []
    text_by_key = {}
    for i, rec in enumerate(records, start=1):
        logging.info("[%d/%d] Checking %s — %s", i, len(records), rec["company"], rec["url"])
        status, error, final_url, title, text = 0, "", rec["url"], "", ""
        retrieval_url = rec["url"]
        recovered_with_public_alternative = False
        application_url = ""
        method = "requests"
        try:
            final_url, title, text, status, note, method = extract_with_requests(session, rec["url"])
            if note:
                error = note
            if USE_PLAYWRIGHT_FALLBACK and not method.startswith(("pymupdf", "ocr", "no_extractable_text")) and (len(text) < 250 or status in (401, 403, 429)):
                try:
                    final_url, title, text, status, note, method = extract_with_playwright(rec["url"])
                    error = note or error
                except Exception as browser_exc:
                    error = (error + "; " if error else "") + f"Playwright fallback: {browser_exc}"
        except Exception as exc:
            error = str(exc)
            if USE_PLAYWRIGHT_FALLBACK:
                try:
                    final_url, title, text, status, note, method = extract_with_playwright(rec["url"])
                    error = note
                except Exception as browser_exc:
                    error += f"; Playwright fallback: {browser_exc}"

        original_access_issue = detect_access_issue(text, status) or detect_login_only_page(rec["url"], title, text)
        if gemini and original_access_issue:
            try:
                alternative_url = gemini.find_public_alternative(
                    rec["company"], rec["url"], original_access_issue
                )
                if alternative_url and alternative_url.casefold() != rec["url"].casefold():
                    try:
                        alt_final_url, alt_title, alt_text, alt_status, alt_note, alt_method = extract_with_requests(
                            session, alternative_url
                        )
                        alt_error = alt_note
                        if USE_PLAYWRIGHT_FALLBACK and (len(alt_text) < 250 or alt_status in (401, 403, 429)):
                            try:
                                alt_final_url, alt_title, alt_text, alt_status, alt_note, alt_method = extract_with_playwright(
                                    alternative_url
                                )
                                alt_error = alt_note or alt_error
                            except Exception as browser_exc:
                                alt_error = (alt_error + "; " if alt_error else "") + f"Playwright fallback: {browser_exc}"
                        alt_access_issue = detect_access_issue(alt_text, alt_status) or detect_login_only_page(
                            alternative_url, alt_title, alt_text
                        )
                        if alt_status < 400 and not alt_access_issue and alt_text.strip():
                            retrieval_url = alternative_url
                            final_url, title, text, status, method = (
                                alt_final_url, alt_title, alt_text, alt_status,
                                f"{alt_method}; grounded_alternative",
                            )
                            recovered_with_public_alternative = True
                            error = f"Original page inaccessible ({original_access_issue}); verified public alternative used."
                            logging.info("Verified public alternative for %s: %s", rec["url"], final_url)
                    except Exception as alternative_exc:
                        logging.warning("Could not verify grounded alternative for %s: %s", rec["url"], alternative_exc)
            except Exception:
                logging.exception("Grounded public-alternative lookup failed for %s", rec["url"])

        application_url = resolve_application_url_from_page(session, retrieval_url, final_url, page_text=text)

        cls = classify(rec["company"], rec["url"], rec["description"], text)
        key = rec["company"].casefold() + "|" + rec["url"].casefold()
        previous = state.get(key, {})
        evidence = opportunity_evidence(text)
        source_deadline_text = extract_deadline_text(text)
        source_deadline_date = _parse_deadline_date(source_deadline_text)
        prior_evidence = set(previous.get("opportunity_evidence", []))
        new_evidence = [item for item in evidence if item not in prior_evidence]
        is_rfp = bool(cls["rfp_terms"] and cls["target_topics"])
        positive_change = bool(new_evidence) and cls["flag"]
        opportunity_type = "RFP" if is_rfp else "Grant"
        if cls["closed_signals"]:
            opportunity_status = "CLOSED"
            status_reason = "Closed wording detected: " + "; ".join(cls["closed_signals"])
        elif cls["open_signals"]:
            opportunity_status = "OPEN"
            status_reason = "Open wording detected: " + "; ".join(cls["open_signals"])
        else:
            opportunity_status = "UNKNOWN"
            status_reason = "No open/closed status terms detected"
        relevance_score = (2 if cls["flag"] else 0) + (1 if cls["topic_terms"] else 0) + (1 if cls["open_signals"] and not cls["closed_signals"] else 0)
        result = {
            "company": rec["company"], "source_url": rec["url"], "final_url": final_url,
            "discovered_from": rec.get("discovered_from", ""), "http_status": status, "extraction_method": method,
            "page_title": title, "opportunity_type": opportunity_type,
            "opportunity_title": title, "opportunity_status": opportunity_status,
            "status_reason": status_reason, "opportunity_id": "", "published_date": "",
            "published_date_text": "", "deadline_text": source_deadline_text,
            "deadline_date": source_deadline_date.isoformat() if source_deadline_date else "",
            "eligibility": "", "application_url": application_url,
            "evidence": "; ".join(new_evidence or evidence),
            "evidence_quote": "",
            "field_evidence": "",
            "funding_amount": "",
            "sponsor": rec["company"],
            "sponsor_type": "pharmaceutical" if not rec.get("discovered_from") else "unknown",
            "therapeutic_areas": "; ".join(cls["topic_terms"]),
            "target_topics": "; ".join(cls["target_topics"]),
            "focus": rec.get("description", ""),
            "geography": "",
            "reference_id": "",
            "contact_email": first_email(text),
            "source_type": "seed_url" if not rec.get("discovered_from") else "search_discovered",
            "source_authority": "unverified",
            "priority": cls["priority"],
            "funding_terms": "; ".join(cls["funding_terms"]), "relevance_score": relevance_score,
            "change_type": "",
            "changed_fields": "",
            "flag": cls["flag"], "gemini_relevance": "", "gemini_confidence": "",
            "gemini_summary": "", "checked_at": now, "error": error,
        }
        result.update(verify_opportunity(rec["url"], final_url, title, text))
        result["access_issue"] = detect_access_issue(text, status)
        result["is_target_rfp"] = cls["is_target_rfp"]
        fingerprint = opportunity_fingerprint(
            final_url, application_url, result.get("opportunity_title", ""),
            result.get("deadline_text", ""), result.get("evidence_quote", "") or result.get("evidence", ""),
            cls["target_topics"],
        )
        is_new_or_changed = previous.get("fingerprint") != fingerprint
        result["change_type"] = "NEW_OR_CHANGED" if is_new_or_changed else "UNCHANGED"
        result["changed_fields"] = "opportunity_fields" if is_new_or_changed else ""
        results.append(result)
        text_by_key[key] = text
        email_evidence = email_evidence_snippets(new_evidence)
        if cls["is_target_rfp"] and email_evidence and (positive_change or args.force_alert) and not args.baseline:
            alert = dict(result)
            alert.update({
                "priority": cls["priority"],
                "change_summary": "New RFP related evidence" if is_rfp else "New grant opportunity evidence",
                "new_evidence": email_evidence,
                "topic_terms": cls["target_topics"],
            })
            alerts.append(alert)
        if (not error or recovered_with_public_alternative) and status < 400:
            state[key] = {
                "fingerprint": fingerprint,
                "checked_at": now,
                "flag": cls["flag"],
                "final_url": final_url,
                "application_url": application_url,
                "opportunity_evidence": evidence,
            }
        else:
            logging.warning("State not advanced for failed check: %s", rec["url"])
        if error:
            logging.warning("Issue for %s: %s", rec["company"], error)
        elif cls["flag"]:
            logging.info("FLAG [%s] %s | funding=%s | topics=%s",
                         cls["priority"], rec["company"], cls["funding_terms"], cls["topic_terms"])
        time.sleep(DELAY_SECONDS)

    if gemini and (args.discover_gemini or GEMINI_DISCOVERY or GEMINI_REVIEW):
        try:
            updates = gemini.enrich_records(results, text_by_key)
            by_url = {str(update.get("source_url", "")).casefold(): update for update in updates}
            alert_keys = {(a["company"].casefold(), a["source_url"].casefold()) for a in alerts}
            for result in results:
                update = by_url.get(result["source_url"].casefold())
                if not update:
                    continue
                merge_enrichment(result, update)
                if update.get("target_topics"):
                    result["target_topics"] = "; ".join(
                        dict.fromkeys(str(value).strip() for value in update["target_topics"] if str(value).strip())
                    )
                result["is_target_rfp"] = is_target_rfp_record(result, update)
                result.update(verify_opportunity(
                    result["source_url"], result.get("final_url", ""),
                    result.get("opportunity_title", ""),
                    text_by_key.get(result["company"].casefold() + "|" + result["source_url"].casefold(), ""),
                ))
                reviewed_deadline = _parse_deadline_date(str(result.get("deadline_text", "")))
                result["deadline_date"] = reviewed_deadline.isoformat() if reviewed_deadline else ""
                result["flag"] = bool(result.get("is_target_rfp"))
                source_text = text_by_key.get(
                    result["company"].casefold() + "|" + result["source_url"].casefold(), ""
                )
                ai_quote = str(update.get("evidence_quote", "")).strip()
                if ai_quote and ai_quote.casefold() in source_text.casefold():
                    result["evidence_quote"] = ai_quote[:700]
                for alert in alerts:
                    if (alert["company"].casefold(), alert["source_url"].casefold()) == (result["company"].casefold(), result["source_url"].casefold()):
                        merge_enrichment(alert, update)
                        alert["target_topics"] = result.get("target_topics", "")
                        alert["is_target_rfp"] = result.get("is_target_rfp", False)
                        alert.update({k: result.get(k, "") for k in ("document_status", "confidence_score", "verification_reason", "evidence_quote", "email_eligible", "email_block_reason", "deadline_text", "deadline_date")})
                        if result.get("evidence_quote"):
                            alert["new_evidence"] = [result["evidence_quote"]]
                state_key = result["company"].casefold() + "|" + result["source_url"].casefold()
                if state_key in state:
                    state[state_key].update({
                        "flag": result["flag"],
                        "application_url": result.get("application_url", ""),
                        "gemini_summary": result.get("gemini_summary", ""),
                    })
                if update.get("relevant") and is_target_rfp_record(result, update) and result["change_type"] == "NEW_OR_CHANGED" and not args.baseline:
                    key = (result["company"].casefold(), result["source_url"].casefold())
                    if key not in alert_keys:
                        alert = dict(result)
                        alert.update({
                            "priority": "HIGH" if update.get("confidence") == "high" else "MEDIUM",
                            "change_summary": "Gemini-reviewed funding opportunity",
                            "new_evidence": [result.get("evidence_quote", "")] if result.get("evidence_quote") else [],
                            "topic_terms": result.get("target_topics", "").split("; ") if result.get("target_topics") else [],
                        })
                        alerts.append(alert)
                        alert_keys.add(key)
            alerts = [alert for alert in alerts if alert.get("email_eligible") == "Yes"]
            logging.info("Gemini post-scrape review processed %d record(s).", len(updates))
        except Exception:
            logging.exception("Gemini post-scrape review failed; deterministic results will be used.")

    # Grounded discovery intentionally runs after the existing pharmaceutical
    # URLs have been scraped and reviewed. Newly discovered URLs are then
    # fetched and independently verified before they can become alerts.
    if gemini and (args.discover_gemini or GEMINI_DISCOVERY):
        try:
            excluded = [record["company"] for record in records]
            discovered = gemini.discover(excluded, max_results=GEMINI_MAX_RESULTS)
            save_discovered(discovered, GEMINI_DISCOVERED_FILE)
            existing_keys = {
                (result["company"].casefold(), result["source_url"].casefold())
                for result in results
            }
            discovered_results = []
            discovered_text = {}
            for candidate in discovered:
                key = (candidate.company.casefold(), candidate.url.casefold())
                if key in existing_keys:
                    continue
                logging.info("Checking grounded discovery: %s — %s", candidate.company, candidate.url)
                status, error, final_url, title, text = 0, "", candidate.url, "", ""
                method = "requests"
                recovered_with_public_alternative = False
                try:
                    final_url, title, text, status, note, method = extract_with_requests(session, candidate.url)
                    error = note or ""
                    if USE_PLAYWRIGHT_FALLBACK and (len(text) < 250 or status in (401, 403, 429)):
                        try:
                            final_url, title, text, status, note, method = extract_with_playwright(candidate.url)
                            error = note or error
                        except Exception as browser_exc:
                            error = (error + "; " if error else "") + f"Playwright fallback: {browser_exc}"
                except Exception as exc:
                    error = str(exc)

                discovered_access_issue = detect_access_issue(text, status) or detect_login_only_page(
                    candidate.url, title, text
                )
                if gemini and discovered_access_issue:
                    try:
                        alternative_url = gemini.find_public_alternative(
                            candidate.company, candidate.url, discovered_access_issue
                        )
                        if alternative_url and alternative_url.casefold() != candidate.url.casefold():
                            alt_final_url, alt_title, alt_text, alt_status, alt_note, alt_method = extract_with_requests(
                                session, alternative_url
                            )
                            alt_access_issue = detect_access_issue(alt_text, alt_status) or detect_login_only_page(
                                alternative_url, alt_title, alt_text
                            )
                            if alt_status < 400 and not alt_access_issue and alt_text.strip():
                                final_url, title, text, status, method = (
                                    alt_final_url, alt_title, alt_text, alt_status,
                                    f"{alt_method}; grounded_alternative",
                                )
                                recovered_with_public_alternative = True
                                error = f"Original page inaccessible ({discovered_access_issue}); verified public alternative used."
                                logging.info("Verified grounded alternative for %s: %s", candidate.url, final_url)
                    except Exception:
                        logging.exception("Grounded alternative lookup failed for discovered URL %s", candidate.url)

                application_url = resolve_application_url_from_page(
                    session, candidate.url, final_url, page_text=text
                )
                cls = classify(candidate.company, candidate.url, candidate.description, text)
                evidence = opportunity_evidence(text)
                deadline_text = extract_deadline_text(text)
                deadline_date = _parse_deadline_date(deadline_text)
                source_key = candidate.company.casefold() + "|" + candidate.url.casefold()
                result = {
                    "company": candidate.company, "source_url": candidate.url, "final_url": final_url,
                    "discovered_from": candidate.discovered_from, "http_status": status,
                    "extraction_method": method, "page_title": title,
                    "opportunity_type": "RFP" if any(term in " ".join(evidence).casefold() for term in RFP_TERMS) else "Grant",
                    "opportunity_title": title, "opportunity_status": "UNKNOWN",
                    "status_reason": "Grounded discovery; status requires verified review",
                    "opportunity_id": "", "published_date": "", "published_date_text": "",
                    "deadline_text": deadline_text,
                    "deadline_date": deadline_date.isoformat() if deadline_date else "",
                    "eligibility": "", "application_url": application_url,
                    "evidence": "; ".join(evidence), "evidence_quote": "", "field_evidence": "",
                    "funding_amount": "", "sponsor": candidate.company, "sponsor_type": "unknown",
                    "therapeutic_areas": "; ".join(cls["topic_terms"]),
                    "target_topics": "; ".join(cls["target_topics"]), "focus": candidate.description,
                    "geography": "", "reference_id": "", "contact_email": first_email(text),
                    "source_type": "search_discovered", "source_authority": "unverified",
                    "priority": cls["priority"], "funding_terms": "; ".join(cls["funding_terms"]),
                    "relevance_score": (2 if cls["flag"] else 0) + (1 if cls["topic_terms"] else 0),
                    "change_type": "NEW_OR_CHANGED", "changed_fields": "page_content",
                    "flag": cls["flag"], "is_target_rfp": cls["is_target_rfp"],
                    "gemini_relevance": "", "gemini_confidence": "",
                    "gemini_summary": "", "checked_at": now, "error": error,
                }
                result.update(verify_opportunity(candidate.url, final_url, title, text))
                result["access_issue"] = detect_access_issue(text, status)
                results.append(result)
                discovered_results.append(result)
                discovered_text[source_key] = text
                existing_keys.add(key)
                if (not error or recovered_with_public_alternative) and status < 400:
                    state[source_key] = {
                        "fingerprint": opportunity_fingerprint(
                            final_url, application_url, result.get("opportunity_title", ""),
                            result.get("deadline_text", ""), result.get("evidence", ""),
                            cls["target_topics"],
                        ),
                        "checked_at": now, "flag": cls["flag"], "final_url": final_url,
                        "application_url": application_url, "opportunity_evidence": evidence,
                    }

            if discovered_results:
                updates = gemini.enrich_records(discovered_results, discovered_text)
                by_url = {str(update.get("source_url", "")).casefold(): update for update in updates}
                for result in discovered_results:
                    update = by_url.get(result["source_url"].casefold())
                    if not update:
                        continue
                    merge_enrichment(result, update)
                    if update.get("target_topics"):
                        result["target_topics"] = "; ".join(
                            dict.fromkeys(str(value).strip() for value in update["target_topics"] if str(value).strip())
                        )
                    result["is_target_rfp"] = is_target_rfp_record(result, update)
                    source_text = discovered_text.get(
                        result["company"].casefold() + "|" + result["source_url"].casefold(), ""
                    )
                    ai_quote = str(update.get("evidence_quote", "")).strip()
                    if ai_quote and ai_quote.casefold() in source_text.casefold():
                        result["evidence_quote"] = ai_quote[:700]
                    result.update(verify_opportunity(
                        result["source_url"], result.get("final_url", ""),
                        result.get("opportunity_title", ""), source_text,
                    ))
                    result["flag"] = bool(result.get("is_target_rfp"))
                    reviewed_deadline = _parse_deadline_date(str(result.get("deadline_text", "")))
                    result["deadline_date"] = reviewed_deadline.isoformat() if reviewed_deadline else ""
                    if (
                        update.get("relevant")
                        and is_target_rfp_record(result, update)
                        and result.get("email_eligible") == "Yes"
                        and result.get("evidence_quote")
                        and not deadline_is_too_soon(result)
                        and not args.baseline
                    ):
                        alert = dict(result)
                        alert.update({
                            "priority": "HIGH" if update.get("confidence") == "high" else "MEDIUM",
                            "change_summary": "Gemini-grounded RFP discovery",
                            "new_evidence": [result["evidence_quote"]],
                            "topic_terms": result.get("therapeutic_areas", "").split("; ") if result.get("therapeutic_areas") else [],
                        })
                        alerts.append(alert)
            logging.info(
                "Grounded discovery ran after existing-URL review: %d candidate(s), %d verified record(s).",
                len(discovered), len(discovered_results),
            )
        except Exception:
            logging.exception("Post-review Gemini grounded discovery failed; existing URL results retained.")

    alerts = [
        alert for alert in alerts
        if (
            alert.get("email_eligible") == "Yes"
            and not deadline_is_too_soon(alert)
            and not alert.get("email_block_reason")
            and not email_url_block_reason(alert.get("application_url") or alert.get("final_url") or alert.get("source_url"))
            and not alert.get("access_issue")
        )
    ]
    logging.info("Filtered expired alerts; remaining eligible alerts: %d", len(alerts))
    # Write full run report and cumulative current flagged results.
    write_records(RUN_REPORT, results, REPORT_COLUMNS)
    write_records(ALERT_REPORT, alerts, REPORT_COLUMNS)
    old = read_records(MATCHES_FILE) if MATCHES_FILE.exists() else []
    by_key = {(str(r.get("company", "")).casefold(), str(r.get("source_url", "")).casefold()): r for r in old}
    for r in results:
        k = (r["company"].casefold(), r["source_url"].casefold())
        if r["flag"]:
            by_key[k] = r
        elif k in by_key:
            del by_key[k]
    write_records(MATCHES_FILE, list(by_key.values()), REPORT_COLUMNS)
    write_markdown(MARKDOWN_REPORT, results, now)

    failed_results = [
        result for result in results
        if result.get("error") or int(result.get("http_status") or 0) >= 500
    ]
    health = {
        "run_id": RUN_ID,
        "checked_at": now,
        "status": "degraded" if failed_results else "ok",
        "records": len(results),
        "failed_records": len(failed_results),
        "flagged_records": sum(bool(result.get("flag")) for result in results),
        "alerts": len(alerts),
    }
    atomic_write_json(HEALTH_FILE, health)
    failure_ratio = len(failed_results) / len(results) if results else 1.0
    if results and failure_ratio >= MAX_FAILURE_RATIO:
        logging.error("Monitor health check failed: %.0f%% of URLs failed.", failure_ratio * 100)
        raise RuntimeError("Monitor health is too degraded to send or advance alerts")

    try:
        # Record candidates before delivery, but never mark them as sent yet.
        persist_run(results, alerts, health, delivered_alerts=[])
    except Exception:
        logging.exception("PostgreSQL update failed; local reports and notifications will continue.")

    # Sync the cumulative matches after writing the local CSV. A Sheets failure
    # must not prevent local reports or email alerts from being produced.
    try:
        sync_google_sheet(MATCHES_FILE)
    except Exception:
        logging.exception("Google Sheets update failed; local CSV and email-alert processing will continue.")

    state_saved = not alerts or args.baseline
    email_sent = False
    email_result = None
    email_error = ""
    if alerts and not args.baseline:
        try:
            email_result = send_email_alert(alerts)
            email_sent = email_result is True
            # Unconfigured email is an intentional optional mode and should
            # not cause the same alert to be retried forever. A real delivery
            # failure remains False and is retried on the next run.
            state_saved = email_result is not False
        except Exception:
            email_error = "SMTP delivery failed"
            logging.exception("Email notification failed; see CSV output for matches.")
            state_saved = False
    if email_sent:
        try:
            # Update delivery state only after SMTP accepted the complete message.
            persist_run(results, alerts, health, delivered_alerts=alerts)
        except Exception:
            logging.exception("PostgreSQL delivery-status update failed; email was sent but DB status may lag.")
    try:
        sync_email_alerts_to_google_sheet(alerts if email_sent else [])
    except Exception:
        logging.exception("Google Sheets email-alert history update failed.")
    if state_saved:
        save_state(state)
    else:
        logging.warning("State was not advanced so these alerts can be retried on the next run.")
    try:
        persist_output_files(OUTPUT_DIR, RUN_ID)
    except Exception:
        logging.exception("PostgreSQL output-file persistence failed; local output files remain available.")
    daily_log = append_daily_run_log(build_daily_run_log_rows(
        results, alerts, args=args, email_result=email_result,
        email_error=email_error, email_sent=email_sent,
    ))
    logging.info("Daily run log: %s", daily_log)
    logging.info("Finished. Flagged this run: %d; positive-change alerts: %d", sum(r["flag"] for r in results), len(alerts))
    logging.info("Run report: %s", RUN_REPORT)
    logging.info("Current flagged matches: %s", MATCHES_FILE)


if __name__ == "__main__":
    main()
