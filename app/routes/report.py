from __future__ import annotations

import os

from fastapi import APIRouter, HTTPException
from fastapi.responses import FileResponse

from app.store import get_report

router = APIRouter()

DOCX_MEDIA_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"


@router.get("/report/{report_id}/download")
async def download_report(report_id: str):
    report = get_report(report_id)
    if report is None or not report.docx_path or not os.path.exists(report.docx_path):
        raise HTTPException(status_code=404, detail="Report not found")

    site_name = report.site_info.site_name.replace(" ", "_")
    filename = f"PEA_{site_name}.docx"

    return FileResponse(
        report.docx_path,
        media_type=DOCX_MEDIA_TYPE,
        filename=filename,
    )


@router.get("/report/{report_id}")
async def get_report_json(report_id: str) -> dict:
    report = get_report(report_id)
    if report is None:
        raise HTTPException(status_code=404, detail="Report not found")

    data = report.model_dump()
    data.pop("docx_path", None)
    return data
