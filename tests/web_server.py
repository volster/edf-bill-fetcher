"""Isolated synthetic server for real browser tests; never uses personal evidence."""

from pathlib import Path
from tempfile import TemporaryDirectory

import uvicorn
from reportlab.pdfgen.canvas import Canvas

from edf_bill_fetcher.io.persistence.result_store import save_result
from edf_bill_fetcher.web.api import create_app
from tests.synthetic_analysis import snapshot


def main():
    """Start a disposable household workspace with one accepted synthetic analysis."""
    with TemporaryDirectory(prefix="edf-browser-") as temporary:
        root = Path(temporary)
        originals = root / "originals"
        originals.mkdir()
        canvas = Canvas(str(originals / "bill.pdf"))
        for i, text in enumerate(
            [
                "Invoice number: KI-12345",
                "Account number: A-12345678",
                "Date issued: 15 February 2024",
                "Your charges: 01 Feb 2024 - 29 Feb 2024",
                "Total charges for this period £125.00 debit",
                "Current balance £250.00 debit",
            ]
        ):
            canvas.drawString(40, 780 - i * 30, text)
        canvas.save()
        app = create_app(root / "server", roots={"Synthetic bills": str(originals)})
        store = app.state.store
        store.user("reviewer", "synthetic-browser-password")
        store.user("admin", "admin", is_admin=True, setup_required=True)
        store.user("viewer", "v")
        entry = store.create(
            "job",
            {
                "request": {"mode": "extract", "name": "Synthetic baseline", "formats": ["xlsx"]},
                "status": "complete",
                "created": 1,
                "cancel": False,
                "error": None,
            },
        )
        result = snapshot(
            [
                {
                    "Date": "01/02/2024",
                    "Entry Type": "New Bill",
                    "Invoice #": "BROWSER-001",
                    "Account": "A-123",
                    "Fuel": "Electricity",
                    "Period From": "01/01/2024",
                    "Period To": "31/01/2024",
                    "Period Charge (£)": 100,
                    "Balance (£)": 100,
                }
            ]
        )
        save_result(result, store.directory(entry) / "result.json", result_id=entry["id"])
        uvicorn.run(app, host="127.0.0.1", port=8766, access_log=False)


if __name__ == "__main__":
    main()
