"""End-to-end tests for DOCX → PDF conversion.

Skipped unless either local `soffice` or the `mcp-docx-soffice:latest` docker
image is available. Builds a real DOCX with python-docx, converts it, then
verifies the produced file is a valid PDF whose extracted text still contains
the source content (Cyrillic, numerals, structural markers).

Run:
    python -m unittest tests.test_pdf_converter
"""
import os
import sys
import tempfile
import unittest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
sys.path.insert(0, '/shared/business/scripts')

from core.pdf_converter import (  # noqa: E402
    convert_docx_to_pdf,
    _local_soffice,
    _docker_image_present,
    LibreOfficeNotInstalled,
)

try:
    from md_to_docx import build_docx
except ImportError as e:
    raise unittest.SkipTest(f"md_to_docx not importable: {e}")

try:
    from pypdf import PdfReader
except ImportError as e:
    raise unittest.SkipTest(f"pypdf not installed (pip install pypdf): {e}")


SAMPLE_MD = """# Договор № 42/2026

## 1. Предмет договора

1.1. Исполнитель оказывает Заказчику услуги.

1.2. Стоимость — 100 000 рублей в месяц.

## 2. Сроки

2.1. Срок: с 01.05.2026 до 30.04.2027.

2.2. Уникальная_метка: ZX9-VERIFY-CYRILLIC-АБВ
"""


def _converter_available() -> bool:
    return bool(_local_soffice()) or _docker_image_present()


@unittest.skipUnless(
    _converter_available(),
    "neither local soffice nor mcp-docx-soffice:latest docker image is available",
)
class DocxToPdfTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls._tmp = tempfile.mkdtemp(prefix="pdf-conv-test-")
        cls.md_path = os.path.join(cls._tmp, "contract.md")
        cls.docx_path = os.path.join(cls._tmp, "contract.docx")
        with open(cls.md_path, "w", encoding="utf-8") as f:
            f.write(SAMPLE_MD)
        build_docx(cls.md_path, cls.docx_path, accept=False)

    @classmethod
    def tearDownClass(cls):
        import shutil
        shutil.rmtree(cls._tmp, ignore_errors=True)

    def _read_pdf_text(self, path: str) -> str:
        reader = PdfReader(path)
        self.assertGreaterEqual(len(reader.pages), 1, "PDF has no pages")
        return "\n".join(page.extract_text() for page in reader.pages)

    def test_default_output_path_is_sibling_pdf(self):
        out = convert_docx_to_pdf(self.docx_path)
        self.addCleanup(os.unlink, out)
        self.assertEqual(out, os.path.splitext(self.docx_path)[0] + ".pdf")
        self.assertTrue(os.path.exists(out))
        self.assertGreater(os.path.getsize(out), 1000, "PDF suspiciously small")

    def test_custom_output_path_is_honored(self):
        custom = os.path.join(self._tmp, "out", "result.pdf")
        out = convert_docx_to_pdf(self.docx_path, output_path=custom)
        self.addCleanup(os.unlink, out)
        self.assertEqual(out, custom)
        self.assertTrue(os.path.exists(custom))

    def test_produced_file_has_pdf_signature(self):
        out = convert_docx_to_pdf(
            self.docx_path,
            output_path=os.path.join(self._tmp, "sig.pdf"),
        )
        self.addCleanup(os.unlink, out)
        with open(out, "rb") as f:
            self.assertEqual(f.read(5), b"%PDF-")

    def test_pdf_text_preserves_source_content(self):
        out = convert_docx_to_pdf(
            self.docx_path,
            output_path=os.path.join(self._tmp, "content.pdf"),
        )
        self.addCleanup(os.unlink, out)
        text = self._read_pdf_text(out)
        # Sanity: the unique marker pins this PDF to our source DOCX,
        # so a stale cached file or wrong conversion would fail loudly.
        for needle in (
            "Договор",
            "42/2026",
            "Предмет договора",
            "Исполнитель",
            "100 000",
            "01.05.2026",
            "30.04.2027",
            "ZX9-VERIFY-CYRILLIC",
            "АБВ",
        ):
            self.assertIn(needle, text, f"missing in PDF text: {needle!r}")

    def test_missing_input_raises(self):
        with self.assertRaises(FileNotFoundError):
            convert_docx_to_pdf("/nonexistent/path/missing.docx")


class NoConverterAvailableTests(unittest.TestCase):
    """Verifies the explicit error path when neither backend is available.

    We can't actually uninstall LibreOffice/docker, so we monkey-patch the
    detection helpers in the converter module to return None / False.
    """

    def test_raises_libreoffice_not_installed(self):
        from core import pdf_converter

        original_local = pdf_converter._local_soffice
        original_docker = pdf_converter._docker_image_present
        pdf_converter._local_soffice = lambda: None
        pdf_converter._docker_image_present = lambda: False
        try:
            with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as f:
                f.write(b"not really a docx, but passes existence check")
                tmp = f.name
            try:
                with self.assertRaises(LibreOfficeNotInstalled) as ctx:
                    convert_docx_to_pdf(tmp)
                self.assertIn("docker build", str(ctx.exception))
            finally:
                os.unlink(tmp)
        finally:
            pdf_converter._local_soffice = original_local
            pdf_converter._docker_image_present = original_docker


if __name__ == "__main__":
    unittest.main()
