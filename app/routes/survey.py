from __future__ import annotations

import os
import uuid

from fastapi import APIRouter, HTTPException

from app.document import build_document
from app.generator import generate_pea_report
from app.models import SurveyRequest, SurveyResponse
from app.store import add_audit_entry, save_report

router = APIRouter()

LOGO_PATH = os.environ.get("LOGO_PATH")


@router.post("/survey", response_model=SurveyResponse)
async def submit_survey(survey: SurveyRequest) -> SurveyResponse:
    survey_id = str(uuid.uuid4())

    try:
        report = await generate_pea_report(survey, survey_id=survey_id)
    except RuntimeError as exc:
        raise HTTPException(status_code=503, detail=str(exc))

    docx_path = build_document(report, logo_path=LOGO_PATH)
    report.docx_path = docx_path

    save_report(report)
    add_audit_entry(
        report.report_id,
        "survey_submitted",
        detail=f"Report generated for {survey.site_info.site_name}",
    )

    return SurveyResponse(
        report_id=report.report_id,
        generated_at=report.generated_at,
        download_url=f"/report/{report.report_id}/download",
    )
