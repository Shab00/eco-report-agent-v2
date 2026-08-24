from __future__ import annotations

import json
import os
import tempfile
import uuid

from fastapi import APIRouter, File, Form, HTTPException, UploadFile
from pydantic import ValidationError

from app.document import build_document
from app.generator import generate_pea_report
from app.models import SurveyRequest, SurveyResponse
from app.store import add_audit_entry, save_report

router = APIRouter()

LOGO_PATH = os.environ.get("LOGO_PATH")


async def _save_upload(upload: UploadFile | None) -> str | None:
    if upload is None or not upload.filename:
        return None
    suffix = os.path.splitext(upload.filename)[1]
    fd, path = tempfile.mkstemp(suffix=suffix)
    try:
        with os.fdopen(fd, "wb") as f:
            f.write(await upload.read())
    finally:
        await upload.close()
    return path


async def _save_uploads(uploads: list[UploadFile]) -> list[str]:
    paths = []
    for upload in uploads:
        path = await _save_upload(upload)
        if path:
            paths.append(path)
    return paths


def _cleanup(*paths: str | None | list[str]) -> None:
    for path in paths:
        if path is None:
            continue
        for p in (path if isinstance(path, list) else [path]):
            try:
                os.remove(p)
            except OSError:
                pass


@router.post("/survey", response_model=SurveyResponse)
async def submit_survey(
    survey: str = Form(..., description="JSON-encoded SurveyRequest payload"),
    photo_descriptions: str = Form(
        "[]", description="JSON-encoded list of photo caption strings, aligned by index with photos"
    ),
    habitat_map: UploadFile | None = File(None),
    location_map: UploadFile | None = File(None),
    proposed_plan: UploadFile | None = File(None),
    photos: list[UploadFile] = File(default=[]),
) -> SurveyResponse:
    try:
        survey_data = SurveyRequest.model_validate_json(survey)
    except ValidationError as exc:
        raise HTTPException(status_code=422, detail=exc.errors())

    try:
        photo_descriptions_list = json.loads(photo_descriptions)
        if not isinstance(photo_descriptions_list, list) or not all(
            isinstance(item, str) for item in photo_descriptions_list
        ):
            raise ValueError("photo_descriptions must be a JSON array of strings")
    except (json.JSONDecodeError, ValueError) as exc:
        raise HTTPException(status_code=422, detail=f"Invalid photo_descriptions: {exc}")

    survey_id = str(uuid.uuid4())

    habitat_map_path = await _save_upload(habitat_map)
    location_map_path = await _save_upload(location_map)
    proposed_plan_path = await _save_upload(proposed_plan)
    photo_paths = await _save_uploads(photos)

    try:
        try:
            report = await generate_pea_report(survey_data, survey_id=survey_id)
        except RuntimeError as exc:
            raise HTTPException(status_code=503, detail=str(exc))

        docx_path = build_document(
            report,
            logo_path=LOGO_PATH,
            habitat_map_path=habitat_map_path,
            location_map_path=location_map_path,
            proposed_plan_path=proposed_plan_path,
            photo_paths=photo_paths,
            photo_descriptions=photo_descriptions_list,
        )
        report.docx_path = docx_path

        save_report(report)
        add_audit_entry(
            report.report_id,
            "survey_submitted",
            detail=f"Report generated for {survey_data.site_info.site_name}",
        )

        return SurveyResponse(
            report_id=report.report_id,
            generated_at=report.generated_at,
            download_url=f"/report/{report.report_id}/download",
        )
    finally:
        _cleanup(habitat_map_path, location_map_path, proposed_plan_path, photo_paths)
