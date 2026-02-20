#!/usr/bin/env python3
"""Altered TCG Cards to Excel / Google Sheets Catalogue.

Fetches Common, Rare, and Exalted cards from the Altered.gg API
and generates an Excel spreadsheet and/or Google Sheet for collection cataloguing.
"""

import argparse

from src.api import fetch_all_cards
from src.config import OUTPUT_FILENAME


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Fetch Altered TCG cards and generate a catalogue."
    )
    parser.add_argument(
        "--excel", action="store_true", help="Generate an Excel (.xlsx) file"
    )
    parser.add_argument(
        "--gsheet", action="store_true", help="Generate a Google Sheet"
    )
    args = parser.parse_args()

    if not args.excel and not args.gsheet:
        args.excel = True
        args.gsheet = True

    print("=== Altered TCG Cards to Catalogue ===\n")

    print("[1] Fetching card data from Altered.gg API ...")
    cards = fetch_all_cards()

    if not cards:
        print("No cards fetched. Exiting.")
        return

    step = 2

    if args.excel:
        from src.excel_builder import build_catalogue

        print(f"\n[{step}] Building Excel catalogue ({len(cards)} cards) ...")
        build_catalogue(cards, OUTPUT_FILENAME)
        step += 1

    if args.gsheet:
        from src.gsheet_builder import build_gsheet

        print(f"\n[{step}] Building Google Sheet ({len(cards)} cards) ...")
        url = build_gsheet(cards)
        print(f"\n  Google Sheet URL: {url}")

    print("\nDone!")


if __name__ == "__main__":
    main()
