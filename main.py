#!/usr/bin/env python3
"""EDF Evidence Collector — simple launch entry point."""
from __future__ import annotations

from edf_bill_fetcher.io.cli import main as _main


def main() -> None:
    _main()


if __name__ == "__main__":
    main()
