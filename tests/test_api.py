import os
import tempfile
import unittest
from pathlib import Path
from unittest.mock import MagicMock, patch

import fitz

os.environ["RENDER_ROOT"] = tempfile.mkdtemp(prefix="resume-renderer-test-")
from fastapi.testclient import TestClient

from app import main
from app.models import RenderRequest
from app.renderer import inspect_pdf


class RendererApiTests(unittest.TestCase):
    def setUp(self):
        self.client = TestClient(main.app)

    def test_missing_key_returns_service_unavailable(self):
        with patch.object(main, "API_KEY", ""):
            self.assertEqual(self.client.post("/v1/render", json={}).status_code, 503)

    def test_wrong_key_is_rejected(self):
        with patch.object(main, "API_KEY", "test-key"):
            response = self.client.post(
                "/v1/render", headers={"Authorization": "Bearer wrong"}, json={}
            )
        self.assertEqual(response.status_code, 401)

    def test_fictional_fixture_validates(self):
        fixture = Path("examples/fictional_resume.json").read_text()
        self.assertEqual(RenderRequest.model_validate_json(fixture).filename_stem, "Taylor_Rivera_Resume")

    def test_pdf_qa_identifies_underfill_and_overflow(self):
        with tempfile.TemporaryDirectory() as directory:
            one_page = Path(directory) / "one-page.pdf"
            document = fitz.open()
            document.new_page().insert_text((72, 72), "Fictional test")
            document.save(one_page)
            document.close()
            self.assertEqual(inspect_pdf(one_page)["status"], "UNDERFILL")

            two_page = Path(directory) / "two-page.pdf"
            document = fitz.open()
            document.new_page().insert_text((72, 72), "Fictional test")
            document.new_page().insert_text((72, 72), "Fictional test")
            document.save(two_page)
            document.close()
            self.assertEqual(inspect_pdf(two_page)["status"], "OVERFLOW")

    def test_no_storage_when_qa_does_not_pass(self):
        payload = RenderRequest.model_validate_json(Path("examples/fictional_resume.json").read_text())
        qa = {"status": "UNDERFILL", "page_count": 1, "warnings": []}
        with patch.object(main, "API_KEY", "test-key"), patch.object(main, "build_docx"), patch.object(main, "convert_to_pdf", return_value=Path("example.pdf")), patch.object(main, "inspect_pdf", return_value=qa), patch.object(main, "_storage_client") as storage:
            response = self.client.post(
                "/v1/render",
                headers={"Authorization": "Bearer test-key"},
                json=payload.model_dump(),
            )
        self.assertEqual(response.status_code, 422)
        storage.assert_not_called()

    def test_storage_errors_do_not_expose_provider_details(self):
        client = MagicMock()
        client.upload_file.side_effect = RuntimeError("provider-specific failure")
        with self.assertRaises(main.HTTPException) as context:
            main._upload_file(client, Path("fictional.pdf"), "key", "application/pdf")
        self.assertEqual(context.exception.detail, "Unable to store generated artifact")


if __name__ == "__main__":
    unittest.main(verbosity=2)
