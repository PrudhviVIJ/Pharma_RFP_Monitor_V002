"""Optional PostgreSQL persistence for monitor runs and opportunity records."""

from __future__ import annotations

from datetime import datetime, timezone
import hashlib
import json
import logging
import mimetypes
import os
from pathlib import Path
from typing import Any


LOGGER = logging.getLogger(__name__)

SCHEMA = """
CREATE TABLE IF NOT EXISTS rfp_monitor_runs (
    run_id TEXT PRIMARY KEY,
    checked_at TIMESTAMPTZ NOT NULL,
    status TEXT NOT NULL,
    records INTEGER NOT NULL DEFAULT 0,
    failed_records INTEGER NOT NULL DEFAULT 0,
    flagged_records INTEGER NOT NULL DEFAULT 0,
    alerts INTEGER NOT NULL DEFAULT 0
);

CREATE TABLE IF NOT EXISTS rfp_opportunities (
    opportunity_key TEXT PRIMARY KEY,
    company TEXT NOT NULL,
    source_url TEXT NOT NULL,
    last_run_id TEXT NOT NULL REFERENCES rfp_monitor_runs(run_id),
    checked_at TIMESTAMPTZ NOT NULL,
    is_flagged BOOLEAN NOT NULL DEFAULT FALSE,
    alert_sent BOOLEAN NOT NULL DEFAULT FALSE,
    data JSONB NOT NULL
);

CREATE INDEX IF NOT EXISTS idx_rfp_opportunities_checked_at
    ON rfp_opportunities (checked_at);
CREATE INDEX IF NOT EXISTS idx_rfp_opportunities_company
    ON rfp_opportunities (company);

CREATE TABLE IF NOT EXISTS rfp_monitor_output_files (
    run_id TEXT NOT NULL REFERENCES rfp_monitor_runs(run_id) ON DELETE CASCADE,
    relative_path TEXT NOT NULL,
    mime_type TEXT NOT NULL DEFAULT 'application/octet-stream',
    size_bytes BIGINT NOT NULL,
    sha256 TEXT NOT NULL,
    content BYTEA NOT NULL,
    stored_at TIMESTAMPTZ NOT NULL,
    PRIMARY KEY (run_id, relative_path)
);

CREATE INDEX IF NOT EXISTS idx_rfp_monitor_output_files_path
    ON rfp_monitor_output_files (relative_path);
"""


def _configured() -> bool:
    enabled = os.getenv("POSTGRES_ENABLED", "false").strip().lower() in {"1", "true", "yes"}
    return enabled and bool(os.getenv("POSTGRES_DSN", os.getenv("DATABASE_URL", "")).strip())


def _key(company: str, source_url: str) -> str:
    value = f"{company.casefold().strip()}|{source_url.casefold().strip()}"
    return hashlib.sha256(value.encode("utf-8")).hexdigest()


def persist_run(
    results: list[dict[str, Any]],
    alerts: list[dict[str, Any]],
    health: dict[str, Any],
    *,
    delivered_alerts: list[dict[str, Any]] | None = None,
) -> bool:
    """Persist the latest run and upsert its records. Returns False when disabled."""
    if not _configured():
        LOGGER.info("PostgreSQL persistence is not configured; skipping database update.")
        return False

    try:
        import psycopg
    except ImportError as exc:
        raise RuntimeError("PostgreSQL persistence requires psycopg; install requirements.txt.") from exc

    dsn = os.getenv("POSTGRES_DSN", os.getenv("DATABASE_URL", "")).strip()
    connect_options = {}
    if not any(token in dsn.casefold() for token in ("localhost", "127.0.0.1", "::1")):
        connect_options["sslmode"] = "require"
    checked_at = datetime.now(timezone.utc)
    run_id = str(health.get("run_id", ""))
    delivered_keys = {
        _key(str(alert.get("company", "")), str(alert.get("source_url", "")))
        for alert in (delivered_alerts or [])
    }

    with psycopg.connect(dsn, **connect_options) as connection:
        with connection.cursor() as cursor:
            cursor.execute(SCHEMA)
            cursor.execute(
                """
                INSERT INTO rfp_monitor_runs
                    (run_id, checked_at, status, records, failed_records, flagged_records, alerts)
                VALUES (%s, %s, %s, %s, %s, %s, %s)
                ON CONFLICT (run_id) DO UPDATE SET
                    checked_at = EXCLUDED.checked_at,
                    status = EXCLUDED.status,
                    records = EXCLUDED.records,
                    failed_records = EXCLUDED.failed_records,
                    flagged_records = EXCLUDED.flagged_records,
                    alerts = EXCLUDED.alerts
                """,
                (
                    run_id, checked_at, health.get("status", "unknown"), health.get("records", 0),
                    health.get("failed_records", 0), health.get("flagged_records", 0), health.get("alerts", 0),
                ),
            )
            for result in results:
                company = str(result.get("company", ""))
                source_url = str(result.get("source_url", ""))
                cursor.execute(
                    """
                    INSERT INTO rfp_opportunities
                        (opportunity_key, company, source_url, last_run_id, checked_at,
                         is_flagged, alert_sent, data)
                    VALUES (%s, %s, %s, %s, %s, %s, %s, %s::jsonb)
                    ON CONFLICT (opportunity_key) DO UPDATE SET
                        company = EXCLUDED.company,
                        source_url = EXCLUDED.source_url,
                        last_run_id = EXCLUDED.last_run_id,
                        checked_at = EXCLUDED.checked_at,
                        is_flagged = EXCLUDED.is_flagged,
                        alert_sent = EXCLUDED.alert_sent,
                        data = EXCLUDED.data
                    """,
                    (
                        _key(company, source_url), company, source_url, run_id, checked_at,
                         bool(result.get("flag")), _key(company, source_url) in delivered_keys,
                        json.dumps(result, ensure_ascii=False, default=str),
                    ),
                )
    LOGGER.info("PostgreSQL persistence complete: run=%s records=%d", run_id, len(results))
    return True


def persist_output_files(output_dir: str | Path, run_id: str) -> bool:
    """Store a complete snapshot of regular files in output_dir for one run."""
    if not _configured():
        LOGGER.info("PostgreSQL persistence is not configured; skipping output-file update.")
        return False

    try:
        import psycopg
    except ImportError as exc:
        raise RuntimeError("PostgreSQL persistence requires psycopg; install requirements.txt.") from exc

    root = Path(output_dir).resolve()
    if not root.is_dir():
        LOGGER.warning("Output directory does not exist; no output files stored: %s", root)
        return False

    dsn = os.getenv("POSTGRES_DSN", os.getenv("DATABASE_URL", "")).strip()
    connect_options = {}
    if not any(token in dsn.casefold() for token in ("localhost", "127.0.0.1", "::1")):
        connect_options["sslmode"] = "require"

    files = [path for path in root.rglob("*") if path.is_file() and not path.is_symlink()]
    stored_at = datetime.now(timezone.utc)
    with psycopg.connect(dsn, **connect_options) as connection:
        with connection.cursor() as cursor:
            cursor.execute(SCHEMA)
            for path in files:
                relative_path = path.relative_to(root).as_posix()
                content = path.read_bytes()
                cursor.execute(
                    """
                    INSERT INTO rfp_monitor_output_files
                        (run_id, relative_path, mime_type, size_bytes, sha256, content, stored_at)
                    VALUES (%s, %s, %s, %s, %s, %s, %s)
                    ON CONFLICT (run_id, relative_path) DO UPDATE SET
                        mime_type = EXCLUDED.mime_type,
                        size_bytes = EXCLUDED.size_bytes,
                        sha256 = EXCLUDED.sha256,
                        content = EXCLUDED.content,
                        stored_at = EXCLUDED.stored_at
                    """,
                    (
                        run_id,
                        relative_path,
                        mimetypes.guess_type(path.name)[0] or "application/octet-stream",
                        len(content),
                        hashlib.sha256(content).hexdigest(),
                        content,
                        stored_at,
                    ),
                )
    LOGGER.info("PostgreSQL output-file persistence complete: run=%s files=%d", run_id, len(files))
    return True
