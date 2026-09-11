# Resume Document Renderer API

An authenticated FastAPI service that converts structured resume content into a DOCX and PDF, evaluates basic one-page content fit, and returns short-lived artifact URLs from S3-compatible storage.

It is a portfolio implementation of document-generation and API design. The repository intentionally contains only fictional example data; it is not a hosted resume service and does not include any personal resumes, generated documents, credentials, or deployment configuration.

## What it demonstrates

- FastAPI request validation and bearer-token authentication
- Deterministic DOCX generation with `python-docx`
- Headless LibreOffice PDF conversion in Docker
- PDF text-layout checks with PyMuPDF
- S3-compatible artifact storage and short-lived presigned download URLs
- Failure handling that prevents artifacts from being stored when PDF fit checks fail

## Design and privacy choices

The render endpoint requires `RENDERER_API_KEY`. It only returns direct, short-lived presigned storage URLs after a successful render. There is no unauthenticated, persistent download gateway. Generated working files are removed after each request.

Do not commit real resumes, contact details, rendered files, `.env` files, credentials, bucket names, or deployment URLs. The included example is fictional.

## Run locally

Copy `.env.example` to `.env` and supply local test values. PDF conversion requires LibreOffice, so Docker is the easiest route:

```bash
docker build -t resume-document-renderer-api .
docker run --rm -p 10000:10000 --env-file .env resume-document-renderer-api
curl http://localhost:10000/health
```

To use object storage, configure an S3-compatible bucket through the environment variables in `.env.example`. Keep the bucket private and apply a retention policy appropriate for the documents you handle.

## Request example

```bash
curl -X POST http://localhost:10000/v1/render \
  -H "Authorization: Bearer $RENDERER_API_KEY" \
  -H "Content-Type: application/json" \
  --data @examples/fictional_resume.json
```

The endpoint returns `422` without storing any artifact when the PDF fails its fit checks. A successful response contains `PASS`, QA metadata, and direct URLs that expire after `SIGNED_URL_TTL_SECONDS`.

## Verification

Run the focused test suite with:

```bash
python -m unittest -v tests.test_api
```

The tests cover authentication, input validation, storage error handling, fictional fixture validation, PDF QA classification, and the no-upload-on-QA-failure rule.
