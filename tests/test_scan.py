import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook, load_workbook
from unique_sheetreview import scan


class ScanTests(unittest.TestCase):
    def setUp(self):
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        self.folder = Path(tmp.name)

    def test_csv(self):
        path = self.folder / "data.csv"
        path.write_text("ID,Date,Name\n001,04/05/2026, Pat\n001,2026-01-01,=literal\n,,Ann\n")
        report = scan(path, required="ID", unique="ID")
        self.assertCountEqual([f.code for f in report.findings],
                              ["unique", "unique", "required", "ambiguous_date", "whitespace"])
        out = report.save_review_copy(self.folder / "review.xlsx")
        ws = load_workbook(out).active
        self.assertEqual(ws["A2"].value, "001")
        self.assertEqual(ws["C3"].data_type, "s")
        self.assertIn("Moderate", ws["B2"].comment.text)
        with self.assertRaises(FileExistsError):
            report.save_review_copy(out)

    def test_xlsx(self):
        book = Workbook()
        ws = book.active
        ws.append(["Amount", "Amount", "Name"])
        ws.append(["1.75", 2, "Bob"])
        ws.merge_cells("C1:D1")
        path = self.folder / "sample.xlsx"
        book.save(path)
        report = scan(path, required=["Code"])
        self.assertCountEqual([f.code for f in report.findings],
                              ["duplicate_header", "duplicate_header", "empty_header", "missing_column", "numeric_text"])
        with self.assertRaises(ValueError):
            report.save_review_copy(path)
        ws = load_workbook(report.save_review_copy(self.folder / "marked.xlsx")).active
        self.assertIn("D1", ws["C1"].comment.text)
        self.assertIsNone(load_workbook(path).active["A2"].comment)

    def test_clean(self):
        path = self.folder / "clean.csv"
        path.write_text("ID,Total\n001,42\n")
        self.assertEqual(scan(path).findings, [])
