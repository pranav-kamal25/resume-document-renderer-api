from __future__ import annotations

import hmac
import math
import os
import re
import shutil
import uuid
from pathlib import Path

import boto3
from botocore.config import Config
from fastapi import Depends, FastAPI, Header, HTTPException

from .models import RenderRequest, RenderResponse
from .renderer import build_docx, convert_to_pdf, inspect_pdf

app = FastAPI(
    title="Resume Document Renderer API",
    version="1.0.0",
    description=(
        "An authenticated API that renders structured resume content to DOCX and PDF, "
        "checks basic PDF content fit, and returns short-lived storage URLs."
    ),
)

ROOT = Path(os.getenv("RENDER_ROOT", "/tmp/resume-renderer"))
ROOT.mkdir(parents=True, exist_ok=True)
API_KEY = os.getenv("RENDERER_API_KEY", "")

S3_ACCESS_KEY_ID = os.getenv("S3_ACCESS_KEY_ID", "")
S3_SECRET_ACCESS_KEY = os.getenv("S3_SECRET_ACCESS_KEY", "")
S3_BUCKET_NAME = os.getenv("S3_BUCKET_NAME", "")
S3_ENDPOINT_URL = os.getenv("S3_ENDPOINT_URL", "")
S3_REGION = os.getenv("S3_REGION", "auto")
SIGNED_URL_TTL_SECONDS = max(60, min(int(os.getenv("SIGNED_URL_TTL_SECONDS", "900")), 604800))


def require_api_key(authorization: str | None = Header(default=None)):
    if not API_KEY:
        raise HTTPException(status_code=503, detail="Service authentication is not configured")
    if not authorization or not hmac.compare_digest(authorization, f"Bearer {API_KEY}"):
        raise HTTPException(status_code=401, detail="Unauthorized")


def _safe_stem(value: str) -> str:
    value = re.sub(r"[^A-Za-z0-9._-]+", "_", value).strip("._-")
    return value[:140] or "resume"


def _storage_configured() -> bool:
    return all([S3_ACCESS_KEY_ID, S3_SECRET_ACCESS_KEY, S3_BUCKET_NAME, S3_ENDPOINT_URL])


def _storage_client():
    if not _storage_configured():
        raise HTTPException(status_code=503, detail="Artifact storage is not fully configured")
    return boto3.client(
        "s3",
        endpoint_url=S3_ENDPOINT_URL,
        aws_access_key_id=S3_ACCESS_KEY_ID,
        aws_secret_access_key=S3_SECRET_ACCESS_KEY,
        region_name=S3_REGION,
        config=Config(signature_version="s3v4", retries={"max_attempts": 3, "mode": "standard"}),
    )


def _upload_file(client, local_path: Path, object_key: str, content_type: str) -> None:
    try:
        client.upload_file(
            str(local_path),
            S3_BUCKET_NAME,
            object_key,
            ExtraArgs={
                "ContentType": content_type,
                "ContentDisposition": f'attachment; filename="{local_path.name}"',
            },
        )
    except Exception as exc:
        raise HTTPException(status_code=502, detail="Unable to store generated artifact") from exc


def _signed_download_url(client, object_key: str, filename: str) -> str:
    try:
        return client.generate_presigned_url(
            ClientMethod="get_object",
            Params={
                "Bucket": S3_BUCKET_NAME,
                "Key": object_key,
                "ResponseContentDisposition": f'attachment; filename="{filename}"',
            },
            ExpiresIn=SIGNED_URL_TTL_SECONDS,
        )
    except Exception as exc:
        raise HTTPException(status_code=502, detail="Unable to create artifact download URL") from exc


@app.get("/health")
def health():
    return {
        "ok": True,
        "service": "resume-document-renderer-api",
        "version": app.version,
        "storage_configured": _storage_configured(),
        "signed_url_ttl_seconds": SIGNED_URL_TTL_SECONDS,
        "qa_statuses": ["PASS", "UNDERFILL", "OVERDENSE", "OVERFLOW", "FAIL"],
    }


@app.post("/v1/render", response_model=RenderResponse, dependencies=[Depends(require_api_key)])
def render_resume(payload: RenderRequest):
    render_id = uuid.uuid4().hex
    work = ROOT / render_id
    work.mkdir(parents=True, exist_ok=False)
    stem = _safe_stem(payload.filename_stem)
    docx_path = work / f"{stem}.docx"

    try:
        build_docx(payload.content, docx_path)
        pdf_path = convert_to_pdf(docx_path, work)
        qa = inspect_pdf(pdf_path)
        if qa["status"] != "PASS":
            raise HTTPException(
                status_code=422,
                detail={
                    "message": "Generated document did not pass PDF fit checks; no artifact was stored.",
                    "qa": qa,
                },
            )

        client = _storage_client()
        pdf_key = f"resumes/{render_id}/{pdf_path.name}"
        _upload_file(client, pdf_path, pdf_key, "application/pdf")
        pdf_url = _signed_download_url(client, pdf_key, pdf_path.name)

        docx_url = None
        if payload.include_docx:
            docx_key = f"resumes/{render_id}/{docx_path.name}"
            _upload_file(
                client,
                docx_path,
                docx_key,
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )
            docx_url = _signed_download_url(client, docx_key, docx_path.name)

        return RenderResponse(
            render_id=render_id,
            status=qa["status"],
            pdf_url=pdf_url,
            docx_url=docx_url,
            qa=qa,
            expires_in_minutes=math.ceil(SIGNED_URL_TTL_SECONDS / 60),
        )
    finally:
        shutil.rmtree(work, ignore_errors=True)
