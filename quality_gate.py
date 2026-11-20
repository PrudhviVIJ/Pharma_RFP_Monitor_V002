"""Content-quality checks used before a funding opportunity reaches email."""

from __future__ import annotations

import re
from datetime import date
from urllib.parse import urlparse


ANNOUNCEMENT_TERMS = (
    "request for proposal", "rfp", "call for proposal", "ccp",
    "call for concept proposal", "independent medical education", "ime",
    "call for grant applications", "cga", "funding opportunity",
    "apply for funding", "submit a proposal", "external funding",
    "call for grants", "calls for grants", "cfg", "cfgs",
)
PROPOSAL_TERMS = ANNOUNCEMENT_TERMS
ACTIONABLE_TERMS = (
    "deadline", "due date", "submit", "application", "eligibility", "eligible",
    "funding", "award", "upload proposal", "apply now", "portal",
)
GENERIC_DOCUMENT_TERMS = (
    "cookie policy", "privacy policy", "terms of use", "tax strategy", "form w-9",
    "w-9", "budget template", "budget instructions", "early access policy",
    "investigational medicines prior to regulatory approval", "annual report",
    "modern slavery statement", "modern slavery", "privacy notice", "consent statement",
    "tip sheet", "training guide", "frequently asked questions", "faq",
    "compliance declaration", "compliance program", "human trafficking",
    "policy statement", "full proposal review", "review summary", "annual declaration",
)
FORM_TERMS = ("template", "application form", "budget form", "w-9", "checklist")
PORTAL_TERMS = ("login", "sign in", "create your password", "portal")
CONTEXT_TERMS = (
    "announcement", "accepting", "deadline", "due date", "eligible", "eligibility",
    "submit", "apply", "award", "applications are", "open", "call for",
)
EXCLUDED_EVIDENCE_TERMS = (
    "privacy", "consent", "cookie", "modern slavery", "terms of use", "annual report",
    "tip sheet", "training guide", "frequently asked questions", "faq", "template",
)
DOCUMENT_EXTENSIONS = (".pdf", ".doc", ".docx", ".xls", ".xlsx")
ACCESS_CHALLENGE_TERMS = ("captcha", "verify you are human", "access denied", "checking your browser", "security challenge")
LOGIN_TERMS = ("login", "log in", "sign in", "sign-in", "username", "password")
CURRENT_STATUS_TERMS = (
    "currently accepting", "current opportunities", "current / open", "open opportunities",
    "now accepting", "applications are open", "open for applications", "active rfp",
    "submission deadline", "application deadline", "due date", "submit by", "deadline",
)
BLOCKED_EMAIL_URL_TERMS = (
    "privacy", "cookie", "terms", "modern-slavery", "modern_slavery", "human-trafficking",
    "annual-report", "tax-strategy", "compliance", "policy", "review-summary", "proposal-review",
    "tip-sheet", "training-guide", "faq",
)


def _contains(text: str, terms: tuple[str, ...]) -> list[str]:
    value = (text or "").casefold()
    matches = []
    for term in terms:
        pattern = re.escape(term.casefold())
        if term.casefold() in {"request for proposal", "call for proposal"}:
            pattern += "s?"
        elif term.casefold() in {"call for grant applications", "cga", "cfg"}:
            pattern += "s?"
        if re.search(r"(?<!\w)" + pattern + r"(?!\w)", value):
            matches.append(term)
    return matches


def email_url_block_reason(url: str) -> str:
    """Return a reason when a URL is clearly not a safe alert destination."""
    parsed = urlparse(url or "")
    value = f"{parsed.path} {parsed.query}".casefold()
    if any(term in value for term in BLOCKED_EMAIL_URL_TERMS):
        return "URL points to a policy, compliance, administrative, or review document"
    if any(term in value for term in PORTAL_TERMS) and any(term in value for term in ("login", "sign-in", "signin")):
        return "URL is a login-only portal"
    return ""


def verify_opportunity(source_url: str, final_url: str, title: str, text: str) -> dict:
    """Classify scraped content and decide whether it is safe for email."""
    source_text = " ".join((title, text or ""))
    url_text = " ".join((source_url, final_url))
    proposal_hits = _contains(source_text, PROPOSAL_TERMS)
    actionable_hits = _contains(source_text, ACTIONABLE_TERMS)
    generic_hits = _contains(source_text, GENERIC_DOCUMENT_TERMS) + _contains(url_text, GENERIC_DOCUMENT_TERMS)
    form_hits = _contains(source_text, FORM_TERMS)
    portal_hits = _contains(source_text, PORTAL_TERMS) + _contains(url_text, PORTAL_TERMS)
    current_hits = _contains(source_text, CURRENT_STATUS_TERMS)
    path = urlparse(final_url or source_url).path.casefold()
    is_document = path.endswith(DOCUMENT_EXTENSIONS)
    blocked_url_reason = email_url_block_reason(final_url or source_url)
    years = [int(value) for value in re.findall(r"\b(?:19|20)\d{2}\b", f"{path} {title}")]
    stale_document = is_document and bool(years) and max(years) < date.today().year and str(date.today().year) not in source_text
    context_sentences = [
        sentence.strip() for sentence in re.split(r"(?<=[.!?])\s+|\s*[•▪]\s*", text or "")
        if _contains(sentence, PROPOSAL_TERMS)
        and _contains(sentence, CONTEXT_TERMS)
        and not _contains(sentence, EXCLUDED_EVIDENCE_TERMS)
        and not re.search(r"\b(?:no|not|without|never|does not|do not)\b", sentence.casefold())
    ]
    contextual_hits = _contains(" ".join(context_sentences), PROPOSAL_TERMS)

    if blocked_url_reason:
        status = "UNRELATED_DOCUMENT"
        reason = blocked_url_reason
    elif stale_document:
        status = "STALE_DOCUMENT"
        reason = f"Document appears to be from {max(years)} and does not contain current-year opportunity evidence"
    elif generic_hits and is_document:
        status = "UNRELATED_DOCUMENT"
        reason = "Generic administrative/policy document: " + ", ".join(generic_hits)
    elif form_hits and not actionable_hits[0:]:
        status = "RELATED_APPLICATION_FORM"
        reason = "Form/template detected without proposal-specific instructions"
    elif contextual_hits and current_hits and len(actionable_hits) >= 2 and len((text or "").strip()) >= 160:
        status = "VERIFIED_PROPOSAL"
        reason = "Current proposal language and actionable application details found"
    elif portal_hits and contextual_hits:
        status = "PORTAL_ONLY"
        reason = "Funding portal found, but no announcement document was verified"
    elif proposal_hits:
        status = "GENERAL_PROGRAM_INFORMATION"
        reason = "Approved term found, but not in an announcement/application context"
    else:
        status = "REVIEW_REQUIRED"
        reason = "No sufficiently strong proposal evidence found"

    score = 0.0
    score += min(0.45, 0.22 * len(set(contextual_hits)))
    score += min(0.35, 0.12 * len(set(actionable_hits)))
    if len((text or "").strip()) >= 160:
        score += 0.10
    if is_document:
        score += 0.05
    if generic_hits or blocked_url_reason or stale_document:
        score -= 0.65
    score = max(0.0, min(1.0, score))
    eligible = status == "VERIFIED_PROPOSAL" and score >= 0.70 and bool(current_hits)
    evidence = ""
    if eligible:
        sentences = re.split(r"(?<=[.!?])\s+|\s*[•▪]\s*", text or "")
        evidence = next((s.strip()[:700] for s in sentences
                         if _contains(s, PROPOSAL_TERMS)
                         and _contains(s, CONTEXT_TERMS)
                         and _contains(s, ACTIONABLE_TERMS)
                         and not _contains(s, EXCLUDED_EVIDENCE_TERMS)
                         and not re.search(r"\b(?:no|not|without|never|does not|do not)\b", s.casefold())), "")
        eligible = bool(evidence)
        if not eligible:
            reason = "Proposal terms found, but no concise evidence sentence was extracted"
    return {
        "document_status": status,
        "confidence_score": f"{score:.2f}",
        "verification_reason": reason,
        "evidence_quote": evidence,
        "email_eligible": "Yes" if eligible else "No",
        "email_block_reason": blocked_url_reason,
    }


def detect_access_issue(text: str, status: int = 0) -> str:
    """Identify blocked/challenge pages for manual review; never bypass them."""
    hits = _contains(text, ACCESS_CHALLENGE_TERMS)
    if hits:
        return ", ".join(hits)
    if status in {401, 403, 429}:
        return f"HTTP {status}"
    return ""


def detect_login_only_page(url: str, title: str, text: str) -> str:
    """Detect an authenticated portal that does not expose public RFP content."""
    value = " ".join((url or "", title or "", text or "")).casefold()
    path = urlparse(url or "").path.casefold()
    login_hits = _contains(value, LOGIN_TERMS)
    portal_path = any(term in path for term in ("login", "signin", "sign-in", "authenticate", "account"))
    credential_form = any(term in value for term in ("enter your password", "forgot password", "create an account", "username and password"))
    if (portal_path or credential_form) and login_hits:
        return "login-only portal"
    return ""
