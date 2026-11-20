"""Render verified monitor records as a concise opportunity brief."""

from __future__ import annotations

from collections import defaultdict
from datetime import datetime
import os
import tempfile
from pathlib import Path
from typing import Iterable, Mapping


def _display(value: object, fallback: str = "Not identified") -> str:
    text = str(value or "").strip()
    return text or fallback


def _link(value: object) -> str:
    url = str(value or "").strip()
    if url.startswith(("http://", "https://")):
        return f"[{url}]({url})"
    return "Not available"


def _areas(record: Mapping[str, object]) -> list[str]:
    raw = record.get("therapeutic_areas") or record.get("topic_terms") or ""
    if isinstance(raw, (list, tuple)):
        return [str(item).strip() for item in raw if str(item).strip()]
    return [item.strip() for item in str(raw).replace(",", ";").split(";") if item.strip()]


def _sort_key(record: Mapping[str, object]) -> tuple[int, str, str]:
    priority = {"HIGH": 0, "MEDIUM": 1, "REVIEW": 2}.get(str(record.get("priority", "REVIEW")), 3)
    deadline = str(record.get("deadline_date") or record.get("deadline_text") or "9999-12-31")
    return priority, deadline, str(record.get("opportunity_title") or record.get("page_title") or "")


def render_markdown(records: Iterable[Mapping[str, object]], checked_at: str | None = None) -> str:
    """Create a review-friendly Markdown report from normalized records."""
    eligible = [
        record for record in records
        if str(record.get("document_status", "")) in {"VERIFIED_PROPOSAL", "PORTAL_ONLY", "GENERAL_PROGRAM_INFORMATION"}
        and bool(record.get("flag"))
    ]
    grouped: dict[str, list[Mapping[str, object]]] = defaultdict(list)
    for record in eligible:
        areas = _areas(record) or ["Other funding opportunities"]
        grouped[areas[0].title()].append(record)

    if checked_at:
        try:
            checked = datetime.fromisoformat(checked_at).astimezone().strftime("%B %d, %Y").replace(" 0", " ")
        except ValueError:
            checked = checked_at
    else:
        checked = datetime.now().astimezone().strftime("%B %d, %Y").replace(" 0", " ")
    lines = [
        "# RFP and funding opportunities found",
        "",
        f"**{len(eligible)} relevant opportunities**",
        "",
        f"Research checked: {checked}",
        "",
    ]
    if not eligible:
        lines.append("No sufficiently verified opportunities were found in this run.")
        return "\n".join(lines) + "\n"

    for area in sorted(grouped):
        lines.extend([f"## {area} opportunities", ""])
        for index, record in enumerate(sorted(grouped[area], key=_sort_key), start=1):
            title = _display(record.get("opportunity_title") or record.get("page_title"))
            sponsor = _display(record.get("sponsor") or record.get("company"))
            status = _display(record.get("opportunity_status"), "REVIEW").title()
            official_url = record.get("application_url") or record.get("final_url") or record.get("source_url")
            source_url = record.get("source_url")
            lines.extend([
                f"{index}. **{sponsor} – {title}**",
                "",
                f"**Status:** {status}",
                f"**Sponsor:** {sponsor}",
                f"**Type:** {_display(record.get('opportunity_type'))}",
                f"**Focus:** {_display(record.get('focus') or record.get('gemini_summary') or record.get('evidence'))}",
                f"**Funding:** {_display(record.get('funding_amount') or record.get('funding_terms'))}",
                f"**Geography:** {_display(record.get('geography'))}",
                f"**Deadline:** {_display(record.get('deadline_text') or record.get('deadline_date'))}",
                f"**Eligible applicants:** {_display(record.get('eligibility'))}",
                f"**Priority:** {_display(record.get('priority'))}",
                f"**Official source:** {_link(official_url)}",
                f"**Source checked:** {_link(source_url)}",
                f"**Evidence:** {_display(record.get('evidence_quote') or record.get('evidence'))}",
                "",
            ])
    return "\n".join(lines)


def write_markdown(path: str | Path, records: Iterable[Mapping[str, object]], checked_at: str | None = None) -> None:
    destination = Path(path)
    destination.parent.mkdir(parents=True, exist_ok=True)
    fd, temporary_name = tempfile.mkstemp(prefix=f".{destination.name}.", suffix=".tmp", dir=destination.parent)
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as handle:
            handle.write(render_markdown(records, checked_at))
            handle.flush()
            os.fsync(handle.fileno())
        os.replace(temporary_name, destination)
    except Exception:
        try:
            os.unlink(temporary_name)
        except FileNotFoundError:
            pass
        raise
