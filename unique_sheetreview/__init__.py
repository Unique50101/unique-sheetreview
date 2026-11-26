"""Scan a spreadsheet and annotate its data quality findings."""
import csv
import re
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from pathlib import Path
from tempfile import TemporaryFile
from shutil import copyfileobj

from openpyxl import Workbook, load_workbook
from openpyxl.comments import Comment
from openpyxl.cell.cell import MergedCell
from openpyxl.styles import PatternFill

COLORS = {"High": "FFC7CE", "Moderate": "FFEB9C", "Low": "DDEBF7"}
DATE = re.compile(r"(\d{1,2})/(\d{1,2})/(\d{4})")
NUMBER = re.compile(r"[+-]?(?:0|[1-9]\d*)(?:\.\d+)?")
AMOUNT = re.compile(r"amount|price|cost|quantity|total|revenue|balance", re.I)


@dataclass(frozen=True)
class Finding:
    severity: str
    sheet: str
    cell: str
    code: str
    message: str


@dataclass
class Report:
    findings: list[Finding]
    _book: Workbook = field(repr=False)
    _source: Path = field(repr=False)

    def __str__(self):
        counts = Counter(f.severity for f in self.findings)
        lines = [" | ".join(f"{s}: {counts[s]}" for s in COLORS)]
        lines.extend(f"[{f.severity}] {f.sheet}!{f.cell} — {f.message}" for f in self.findings)
        return "\n".join(lines)

    def save_review_copy(self, path):
        """Create a new .xlsx review workbook. Existing files are protected."""
        target = Path(path)
        if target.suffix.lower() != ".xlsx" or target.resolve() == self._source:
            raise ValueError("Choose a new .xlsx output path.")
        if target.exists():
            raise FileExistsError(target)
        grouped = defaultdict(list)
        for finding in self.findings:
            ws = self._book[finding.sheet]
            address = finding.cell
            if isinstance(ws[address], MergedCell):
                area = next(r for r in ws.merged_cells.ranges if address in r)
                address = ws.cell(area.min_row, area.min_col).coordinate
            grouped[finding.sheet, address].append(finding)
        # Clone the scan snapshot so multiple exports never accumulate comments.
        with TemporaryFile() as snapshot:
            self._book.save(snapshot)
            snapshot.seek(0)
            book = load_workbook(snapshot)
        for (sheet, address), issues in grouped.items():
            cell = book[sheet][address]
            severity = next(s for s in COLORS if any(i.severity == s for i in issues))
            cell.fill = PatternFill("solid", fgColor=COLORS[severity])
            notes = [f"[{i.severity}] {i.cell}: {i.message}" for i in issues]
            if cell.comment:
                notes.insert(0, f"{cell.comment.author}: {cell.comment.text}")
            cell.comment = Comment("\n".join(notes), "SheetReview")
        with TemporaryFile() as output:
            book.save(output)
            output.seek(0)
            with target.open("xb") as dest:
                copyfileobj(output, dest)
        return target


def _open(path):
    if path.suffix.lower() == ".xlsx":
        return load_workbook(path)
    if path.suffix.lower() != ".csv":
        raise ValueError("Supported formats are .csv and .xlsx.")
    book = Workbook()
    ws = book.active
    ws.title = "Data"
    with path.open(encoding="utf-8-sig", newline="") as stream:
        for row_no, row in enumerate(csv.reader(stream, strict=True), 1):
            if len(row) > 16384 or row_no > 1048576:
                raise ValueError("CSV exceeds Excel worksheet limits.")
            for col, value in enumerate(row, 1):
                if len(value) > 32767:
                    raise ValueError(f"Value exceeds Excel cell limit on row {row_no}.")
                cell = ws.cell(row_no, col, value)
                cell.data_type = "s"  # Preserve IDs and literal text, including leading '='.
    return book


def _names(names):
    if isinstance(names, str):
        names = [names]
    return {name.strip().casefold() for name in names}


def scan(path, *, required=(), unique=(), sheet=None):
    """Inspect row-one headers and subsequent records, optionally on one sheet."""
    source = Path(path)
    book = _open(source)
    required, unique = _names(required), _names(unique)
    findings = []
    for ws in ([book[sheet]] if sheet is not None else book.worksheets):
        def flag(cell, severity, code, message):
            findings.append(Finding(severity, ws.title, cell.coordinate, code, message))

        headers = next(ws.iter_rows(min_row=1, max_row=1))
        names = [str(c.value).strip().casefold() if c.value is not None else "" for c in headers]
        counts = Counter(names)
        for c, name in zip(headers, names):
            if not name:
                flag(c, "High", "empty_header", "Give this column a name.")
            elif counts[name] > 1:
                flag(c, "High", "duplicate_header", "Give each column a distinct name.")
        for missing in sorted((required | unique) - set(names)):
            flag(headers[0], "High", "missing_column", f"Add the expected column '{missing}'.")
        seen = defaultdict(dict)
        repeated = set()
        for row in ws.iter_rows(min_row=2):
            if all(c.value is None or c.value == "" for c in row):
                flag(row[0], "Low", "blank_row", "Review this empty row before import.")
                continue
            for c, name in zip(row, names):
                value = c.value
                if value is None or (isinstance(value, str) and not value.strip()):
                    if name in required:
                        flag(c, "High", "required", "Enter the required value.")
                    continue
                if name in unique:
                    key = (type(value), value)
                    previous = seen[c.column].setdefault(key, c)
                    if previous != c:
                        msg = "This value repeats in a column declared unique."
                        if previous.coordinate not in repeated:
                            flag(previous, "High", "unique", msg)
                            repeated.add(previous.coordinate)
                        flag(c, "High", "unique", msg)
                if c.data_type == "e":
                    flag(c, "High", "excel_error", "Resolve the Excel error before import.")
                if not isinstance(value, str) or c.data_type == "f":
                    continue
                clean = value.strip()
                if clean != value:
                    flag(c, "Low", "whitespace", "Trim leading or trailing whitespace.")
                date = DATE.fullmatch(clean)
                if date:
                    a, b = map(int, date.groups()[:2])
                    if a != b and 1 <= a <= 12 and 1 <= b <= 12:
                        flag(c, "Moderate", "ambiguous_date", "Confirm day/month order or use YYYY-MM-DD.")
                if source.suffix.lower() == ".xlsx" and AMOUNT.search(name) and NUMBER.fullmatch(clean):
                    flag(c, "Moderate", "numeric_text", "Review this numeric text and convert to a number if appropriate.")
    return Report(findings, book, source.resolve())
