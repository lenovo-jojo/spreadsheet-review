"""End-to-end checks using entirely invented catalog rows."""
import csv
import json
import os
from pathlib import Path
import subprocess
import tempfile
import unittest
import zipfile
from xml.sax.saxutils import escape


def workbook(path, rows):
    xml_rows = []
    for index, row in enumerate(rows, 1):
        cells = []
        for column, value in sorted(row.items(), key=lambda item: (len(item[0]), item[0])):
            cells.append(f'<c r="{column}{index}" t="inlineStr"><is><t>{escape(value)}</t></is></c>')
        xml_rows.append(f'<row r="{index}">{"".join(cells)}</row>')
    main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    rels = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    package = "http://schemas.openxmlformats.org/package/2006/relationships"
    # A complete minimal package, so desktop Excel can open it as well as the macOS parser.
    with zipfile.ZipFile(path, "w") as z:
        z.writestr("[Content_Types].xml", '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/report.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>')
        z.writestr("_rels/.rels", f'<Relationships xmlns="{package}"><Relationship Id="rId1" Type="{rels}/officeDocument" Target="xl/workbook.xml"/></Relationships>')
        z.writestr("xl/workbook.xml", f'<workbook xmlns="{main}" xmlns:r="{rels}"><sheets><sheet name="Report" sheetId="1" r:id="rId2"/></sheets></workbook>')
        z.writestr("xl/_rels/workbook.xml.rels", f'<Relationships xmlns="{package}"><Relationship Id="rId2" Type="{rels}/worksheet" Target="worksheets/report.xml"/></Relationships>')
        z.writestr("xl/worksheets/report.xml", f'<worksheet xmlns="{main}"><sheetData>{"".join(xml_rows)}</sheetData></worksheet>')


def read_csv(path):
    with path.open(newline="", encoding="utf-8-sig") as f:
        return list(csv.reader(f))


class ClaimantsTests(unittest.TestCase):
    def run_tool(self, path, **env):
        return subprocess.run([os.environ["CLAIMANTS_BINARY"], str(path)], env={**os.environ, "CLAIMANTS_NO_CLIPBOARD": "1", **env}, text=True, encoding="utf-8", capture_output=True)

    def test_filter_deduplicate_and_csv_escape(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "synthetic.xlsx"
            header = {"E": "Transfer", "K": "Member ID", "L": "Member", "M": "Publisher ID", "N": "Publisher", "Q": "Admin ID", "R": "Admin"}
            claimant = {"K": "001", "L": 'Example, "Studio"', "M": "002", "N": "Line\nTwo", "Q": "003", "R": "Synthetic Admin"}
            different = {**claimant, "Q": "004"}
            workbook(path, [header, {**claimant, "E": " Yes "}, {**claimant, "E": "(Y)"}, {**different, "E": "y"}, {**different, "E": "No"}])
            result = self.run_tool(path)
            self.assertEqual(result.returncode, 0, result.stderr)
            rows = read_csv(path.with_name("synthetic_yes_claimants.csv"))
            self.assertEqual(len(rows), 3)
            self.assertEqual(rows[0], ["Admin", "Admin P#", "Publisher", "Publisher P#"])
            self.assertIn("True Partial", result.stdout)
            self.assertEqual(rows[1], [claimant[c] for c in ["R", "Q", "N", "M"]])
            self.assertEqual(rows[2][1], "004")

    def test_classification_and_member_independent_deduplication(self):
        for answer, expected in [("", "True Full"), (" Yes ", "True Full"), ("(N)", "True Partial"), ("unknown", "Needs review")]:
            with self.subTest(answer=answer), tempfile.TemporaryDirectory() as directory:
                path = Path(directory) / "synthetic.xlsx"
                claimant = {"R": "Admin", "Q": "001", "N": "Publisher", "M": "002"}
                workbook(path, [{"E": "Transfer"}, {**claimant, "E": "YES", "K": "member1"}, {**claimant, "E": "Y", "K": "member2"}, {**claimant, "E": answer}])
                result = self.run_tool(path)
                self.assertEqual(result.returncode, 0, result.stderr)
                self.assertIn(expected, result.stdout)
                self.assertEqual(len(read_csv(path.with_name("synthetic_yes_claimants.csv"))), 2)

    def test_column_settings_file(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "synthetic.xlsx"
            settings = Path(directory) / "columns.json"
            settings.write_text(json.dumps({"answerColumn": "B", "fields": [{"header": "Name", "column": "C"}, {"header": "Code", "column": "D"}]}))
            workbook(path, [{"B": "Transfer"}, {"B": "Yes", "C": "Synthetic Name", "D": "123"}, {"B": "Y", "C": "Synthetic Name", "D": "123"}, {"B": "No", "C": "Skipped", "D": "999"}])
            result = self.run_tool(path, CLAIMANTS_COLUMNS=str(settings))
            self.assertEqual(result.returncode, 0, result.stderr)
            self.assertIn("True Partial", result.stdout)
            self.assertEqual(read_csv(path.with_name("synthetic_yes_claimants.csv")), [["Name", "Code"], ["Synthetic Name", "123"]])

    def test_invalid_column_settings_report_error(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "synthetic.xlsx"
            settings = Path(directory) / "columns.json"
            settings.write_text(json.dumps({"answerColumn": "5", "fields": []}))
            workbook(path, [{"E": "Transfer"}, {"E": "Yes"}])
            result = self.run_tool(path, CLAIMANTS_COLUMNS=str(settings))
            self.assertNotEqual(result.returncode, 0)
            self.assertIn("Column settings", result.stderr)

    def test_no_matches_reports_error(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "synthetic.xlsx"
            workbook(path, [{"E": "Transfer"}, {"E": "No"}])
            result = self.run_tool(path)
            self.assertNotEqual(result.returncode, 0)
            self.assertIn("No YES", result.stderr)

    def test_rejects_other_file_types(self):
        self.assertNotEqual(self.run_tool(Path("input.csv")).returncode, 0)


if __name__ == "__main__":
    unittest.main()
