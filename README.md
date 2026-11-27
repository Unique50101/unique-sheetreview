# SheetReview

## Use in Colab or local Jupyter

[Open the quickstart in Colab](https://colab.research.google.com/github/Unique50101/unique-sheetreview/blob/main/Quickstart.ipynb) or download [Quickstart.ipynb](Quickstart.ipynb) and open it with a Python 3 Jupyter kernel. Run the cells from top to bottom. The example scans a CSV, prints severity levels, and writes a highlighted Excel review copy. Replace the sample path with your own CSV or workbook path.


Give SheetReview a CSV or Excel file and get a list of data-quality issues with the sheet, cell, priority, and suggested fix. Export the findings as an Excel copy with highlighted cells and comments for the person who supplied the data.

## Quick start

Requires Python 3.10 or later.

```sh
python -m pip install unique-sheetreview
```

```python
from unique_sheetreview import scan

report = scan("orders.xlsx")
print(report)
report.save_review_copy("orders_review.xlsx")
```

## Try the sample

From the repository root, use the included [orders.csv](examples/orders.csv). For this example, each row must have a distinct order ID.

```python
from unique_sheetreview import scan

report = scan("examples/orders.csv", required=["OrderID"], unique=["OrderID"])
print(report)
report.save_review_copy("orders_review.xlsx")
```

The sample produces **3 High, 1 Moderate, and 1 Low** findings.

| Priority | Cells | Finding |
| --- | --- | --- |
| High | A2, A3 | Order ID 101 occurs twice. |
| High | A4 | A required order ID is missing. |
| Moderate | B2 | 04/05/2026 has two possible day/month interpretations. |
| Low | C2 | The name has a leading space. |

## Checks

| Priority | What the scanner checks |
| --- | --- |
| High | Missing or duplicate headers, absent rule columns, empty required values, repeated values in declared unique columns, stored Excel error cells |
| Moderate | Ambiguous slash dates, numeric text in Excel amount, price, cost, quantity, total, revenue, or balance columns |
| Low | Blank rows and surrounding whitespace in text values |

Use `report.findings` to integrate with an upload form. Each finding has `severity`, `sheet`, `cell`, `code`, and `message`. Use `scan(path, sheet="Orders")` to restrict an Excel scan to one worksheet. By default the scanner runs on all worksheets, applying the same rules to each. Header rule names ignore case and surrounding spaces. Unique checks compare typed values exactly and skip blanks. Add a required rule when blanks should also be reported.

## Input and export behavior

- Supports comma-delimited UTF-8 CSV (including BOM) and ordinary .xlsx workbooks with headers in row one. Files are processed locally in memory.
- CSV identifiers keep their leading zeros. CSV values beginning with an equals sign remain literal text in the review workbook.
- The review copy uses a new .xlsx path and retains cell values. Review copies of complex workbooks are subject to openpyxl's format support. Formulas are preserved and are not evaluated.
- Review labels are based on the checks above. Business-specific rules, additional layouts, and large streaming inputs are outside this release.

## Development

```sh
python -m pip install -e .
python -m unittest discover -s tests -v
```

Released under the [MIT license](LICENSE).
