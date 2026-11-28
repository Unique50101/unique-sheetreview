# SheetReview

[![Open in Colab](https://colab.research.google.com/assets/colab-badge.svg)](https://colab.research.google.com/github/Unique50101/unique-sheetreview/blob/main/SheetReview.ipynb) [![Launch Jupyter](https://img.shields.io/badge/Jupyter-launch-F37626?logo=jupyter&logoColor=white)](https://mybinder.org/v2/gh/Unique50101/unique-sheetreview/main?urlpath=lab%2Ftree%2FSheetReview.ipynb) [![Open in VS Code](https://img.shields.io/badge/VS_Code-open-007ACC)](https://vscode.dev/github/Unique50101/unique-sheetreview/blob/main/SheetReview.ipynb)

## Install the interface by name

Run `%pip -q install --upgrade unique-interface` in a notebook to update the notebook interface. SheetReview uses its file picker and download panel. [Unique Interface developer guide](https://uniqueopensources.web.app/docs/unique-interface)

## Run the notebook

Run the notebook to open the file picker. Select a CSV or XLSX, click **Review file**, then download the marked workbook. A separate cell includes a runnable sample. The Jupyter badge launches Binder, a public hosted service. For local files, download [SheetReview.ipynb](https://raw.githubusercontent.com/Unique50101/unique-sheetreview/main/SheetReview.ipynb) and open it in local JupyterLab. The VS Code badge opens the notebook editor. Select a Python kernel or open the notebook in desktop VS Code to execute it.

Install the notebook extra with `python -m pip install "unique-sheetreview[notebook]"`. In a Python notebook, run:

```python
from unique_sheetreview.notebook import launch_review

panel = launch_review()
```

The interface is provided by [Unique Interface](https://github.com/Unique50101/unique-interface). The original [cell-based example](https://github.com/Unique50101/unique-sheetreview/blob/main/Quickstart.ipynb) is also available.

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
