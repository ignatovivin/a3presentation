from __future__ import annotations

import asyncio
import importlib
import os
import shutil
import tempfile
import unittest
from io import BytesIO
from pathlib import Path

from docx import Document
from fastapi import HTTPException
from starlette.datastructures import UploadFile

from a3presentation import settings as settings_module
from a3presentation.api import routes as routes_module
from a3presentation.domain.template import TemplateManifest
from a3presentation.domain.api import TextPlanRequest
from a3presentation.domain.presentation import PresentationPlan, SlideKind, SlideSpec


class ApiContractTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls._env_backup = {key: os.environ.get(key) for key in ("TEMPLATES_DIR", "OUTPUTS_DIR", "STORAGE_DIR")}
        cls._temp_dir = tempfile.TemporaryDirectory()
        cls._root = Path(cls._temp_dir.name)
        cls._templates_dir = cls._root / "templates"
        cls._outputs_dir = cls._root / "outputs"
        cls._templates_dir.mkdir(parents=True, exist_ok=True)
        cls._outputs_dir.mkdir(parents=True, exist_ok=True)

        source_templates = Path(__file__).resolve().parents[1] / "storage" / "templates"
        for template_id in ("corp_light_v1", "demo_business"):
            shutil.copytree(source_templates / template_id, cls._templates_dir / template_id)

        os.environ["TEMPLATES_DIR"] = str(cls._templates_dir)
        os.environ["OUTPUTS_DIR"] = str(cls._outputs_dir)
        os.environ["STORAGE_DIR"] = str(cls._root)

        importlib.reload(settings_module)
        importlib.reload(routes_module)

    @classmethod
    def tearDownClass(cls) -> None:
        cls._temp_dir.cleanup()
        for key, value in cls._env_backup.items():
            if value is None:
                os.environ.pop(key, None)
            else:
                os.environ[key] = value
        importlib.reload(settings_module)
        importlib.reload(routes_module)

    def test_health_endpoint_returns_ok(self) -> None:
        self.assertEqual(routes_module.healthcheck(), {"status": "ok"})

    def test_templates_endpoint_lists_available_templates(self) -> None:
        templates = routes_module.list_templates()
        template_ids = {item.template_id for item in templates}
        self.assertIn("corp_light_v1", template_ids)
        self.assertIn("demo_business", template_ids)

    def test_template_details_expose_missing_template_file(self) -> None:
        response = routes_module.get_template("demo_business")
        self.assertFalse(response.has_template_file)

    def test_template_details_reject_path_traversal_template_id(self) -> None:
        with self.assertRaises(HTTPException) as error:
            routes_module.get_template("..\\..\\outside")

        self.assertEqual(error.exception.status_code, 400)
        self.assertIn("escapes the storage root", error.exception.detail)

    def test_extract_text_endpoint_returns_blocks_tables_and_chart_assessments(self) -> None:
        document = Document()
        document.add_paragraph("Отчет по сегментам")
        document.add_heading("Рынок", level=1)
        table = document.add_table(rows=3, cols=2)
        table.cell(0, 0).text = "Сегмент"
        table.cell(0, 1).text = "Выручка"
        table.cell(1, 0).text = "SMB"
        table.cell(1, 1).text = "120"
        table.cell(2, 0).text = "Enterprise"
        table.cell(2, 1).text = "250"
        buffer = BytesIO()
        document.save(buffer)
        upload = UploadFile(filename="report.docx", file=BytesIO(buffer.getvalue()))

        payload = asyncio.run(routes_module.extract_document_text(upload))
        self.assertEqual(payload.file_name, "report.docx")
        self.assertTrue(payload.text)
        self.assertEqual(len(payload.tables), 1)
        self.assertEqual(len(payload.chart_assessments), 1)

    def test_plan_from_text_returns_presentation_plan(self) -> None:
        payload = routes_module.plan_from_text(
            TextPlanRequest(
                template_id="corp_light_v1",
                title="Demo",
                raw_text="Основные выводы\n- Рост выручки\n- Снижение churn",
            )
        )
        self.assertEqual(payload.template_id, "corp_light_v1")
        self.assertGreaterEqual(len(payload.slides), 1)

    def test_generate_and_download_presentation_for_valid_template(self) -> None:
        payload = routes_module.generate_presentation(
            PresentationPlan(
                template_id="corp_light_v1",
                title="Smoke Test",
                slides=[
                    SlideSpec(kind=SlideKind.TITLE, title="Smoke Test", subtitle="API contract"),
                    SlideSpec(kind=SlideKind.TEXT, title="Итог", text="Проверка generate/download через API."),
                ],
            )
        )
        self.assertTrue((self._outputs_dir / payload.file_name).exists())

        download = routes_module.download_presentation(payload.file_name)
        self.assertEqual(
            download.media_type,
            "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )
        self.assertEqual(Path(download.path).name, payload.file_name)

    def test_generate_returns_404_for_template_without_pptx(self) -> None:
        with self.assertRaises(HTTPException) as error:
            routes_module.generate_presentation(
                PresentationPlan(
                    template_id="demo_business",
                    title="Broken Template",
                    slides=[SlideSpec(kind=SlideKind.TITLE, title="Broken Template")],
                )
            )

        self.assertEqual(error.exception.status_code, 404)
        self.assertIn("Template PPTX not found", error.exception.detail)

    def test_upload_template_rejects_manifest_path_traversal(self) -> None:
        manifest = TemplateManifest(
            template_id="..\\..\\outside",
            display_name="Broken",
            source_pptx="template.pptx",
        )
        upload = UploadFile(
            filename="template.pptx",
            file=BytesIO(b"not-a-real-pptx"),
        )

        with self.assertRaises(HTTPException) as error:
            asyncio.run(
                routes_module.upload_template(
                    manifest_json=manifest.model_dump_json(),
                    template_file=upload,
                )
            )

        self.assertEqual(error.exception.status_code, 400)
        self.assertIn("escapes the storage root", error.exception.detail)


if __name__ == "__main__":
    unittest.main()
