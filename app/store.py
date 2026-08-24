from __future__ import annotations

from datetime import datetime, timezone

from app.models import PEAReport

_reports: dict[str, PEAReport] = {}
_audit_log: list[dict] = []


def save_report(report: PEAReport) -> None:
    _reports[report.report_id] = report


def get_report(report_id: str) -> PEAReport | None:
    return _reports.get(report_id)


def add_audit_entry(report_id: str, action: str, detail: str = "") -> None:
    _audit_log.append(
        {
            "report_id": report_id,
            "action": action,
            "detail": detail,
            "timestamp": datetime.now(timezone.utc).isoformat(),
        }
    )


def get_audit_log() -> list[dict]:
    return list(_audit_log)
