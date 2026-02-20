#!/usr/bin/env python3
"""Altered TCG Cards to Excel Catalogue.

Fetches Common, Rare, and Exalted cards from the Altered.gg API
and generates an Excel spreadsheet for collection cataloguing.
"""

from src.api import fetch_all_cards
from src.config import OUTPUT_FILENAME
from src.excel_builder import build_catalogue


def main() -> None:
    print("=== Altered TCG Cards to Excel Catalogue ===\n")

    print("[1/2] Fetching card data from Altered.gg API ...")
    cards = fetch_all_cards()

    if not cards:
        print("No cards fetched. Exiting.")
        return

    print(f"\n[2/2] Building Excel catalogue ({len(cards)} cards) ...")
    build_catalogue(cards, OUTPUT_FILENAME)

    print("\nDone!")


if __name__ == "__main__":
    main()
