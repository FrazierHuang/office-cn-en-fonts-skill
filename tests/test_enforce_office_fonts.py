#!/usr/bin/env python3
from __future__ import annotations

import shutil
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile


ROOT = Path(__file__).resolve().parents[1]
SCRIPT = ROOT / "scripts" / "enforce_office_fonts.py"


def make_docx(path: Path) -> None:
    with ZipFile(path, "w", ZIP_DEFLATED) as archive:
        archive.writestr(
            "[Content_Types].xml",
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>',
        )
        archive.writestr(
            "word/document.xml",
            """<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:tbl><w:tr><w:tc><w:tcPr><w:shd w:fill="FFFF00"/></w:tcPr><w:p><w:r><w:t>河蚬 As 暴露</w:t></w:r></w:p></w:tc></w:tr></w:tbl></w:body></w:document>""",
        )


def make_xlsx(path: Path) -> None:
    with ZipFile(path, "w", ZIP_DEFLATED) as archive:
        archive.writestr(
            "[Content_Types].xml",
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>',
        )
        archive.writestr(
            "xl/styles.xml",
            """<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><name val="Calibri"/><color rgb="FFFF0000"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/></patternFill></fill></fills><cellXfs count="1"><xf fontId="0" fillId="2"/></cellXfs></styleSheet>""",
        )
        archive.writestr(
            "xl/sharedStrings.xml",
            """<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si><t>河蚬 As 暴露</t></si></sst>""",
        )
        archive.writestr(
            "xl/worksheets/sheet1.xml",
            """<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row><c r="A1" t="inlineStr"><is><t>测试 DEF 和 As 暴露</t></is></c></row></sheetData></worksheet>""",
        )


def make_pptx(path: Path) -> None:
    with ZipFile(path, "w", ZIP_DEFLATED) as archive:
        archive.writestr(
            "[Content_Types].xml",
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>',
        )
        archive.writestr(
            "ppt/slides/slide1.xml",
            """<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:sp><p:txBody><a:p><a:r><a:rPr/><a:t>河蚬 As 暴露</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>""",
        )


class OfficeFontScriptTest(unittest.TestCase):
    def setUp(self) -> None:
        self.tmp = Path(tempfile.mkdtemp(prefix="office-font-test-"))
        make_docx(self.tmp / "sample.docx")
        make_xlsx(self.tmp / "sample.xlsx")
        make_pptx(self.tmp / "sample.pptx")

    def tearDown(self) -> None:
        shutil.rmtree(self.tmp)

    def run_script(self, *args: str) -> subprocess.CompletedProcess[str]:
        return subprocess.run(
            [sys.executable, str(SCRIPT), *args],
            cwd=ROOT,
            text=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            check=False,
        )

    def test_check_fails_before_and_passes_after_recursive_fix(self) -> None:
        before = self.run_script("--check", "--recursive", str(self.tmp))
        self.assertNotEqual(before.returncode, 0, before.stdout)
        self.assertIn("FAIL", before.stdout)

        fixed = self.run_script("--recursive", "--in-place", "--backup", str(self.tmp))
        self.assertEqual(fixed.returncode, 0, fixed.stderr)
        self.assertTrue((self.tmp / "sample.backup.docx").exists())

        after = self.run_script("--check", "--recursive", str(self.tmp))
        self.assertEqual(after.returncode, 0, after.stdout + after.stderr)
        self.assertIn("PASS", after.stdout)

        with ZipFile(self.tmp / "sample.docx") as archive:
            text = archive.read("word/document.xml").decode("utf-8")
        self.assertIn("河蚬As暴露", text)
        self.assertIn("Times New Roman", text)
        self.assertIn("宋体", text)
        self.assertNotIn("FFFF00", text)

    def test_non_in_place_writes_fontfixed_file(self) -> None:
        result = self.run_script(str(self.tmp / "sample.docx"))
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertTrue((self.tmp / "sample_fontfixed.docx").exists())


if __name__ == "__main__":
    unittest.main()
