"""Notebook upload and download using the Unique Interface package."""
from . import scan


def launch_review(*, required=(), unique=(), sheet=None):
    """Open a notebook file picker with optional column rules."""
    from unique_interface import FileInterface

    def review(path, folder):
        report = scan(path, required=required, unique=unique, sheet=sheet)
        return str(report), report.save_review_copy(folder / "sheet_review.xlsx")

    return FileInterface(review, title="SheetReview").launch()
