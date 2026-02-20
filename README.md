# Altered TCG Cards to Excel Catalogue

Fetches **Common**, **Rare**, and **Exalted** cards from the [Altered.gg](https://www.altered.gg) public API and generates an Excel spreadsheet for cataloguing your collection.

## Features

- Pulls all Common, Rare, and Exalted cards from every faction
- Embeds a miniature card thumbnail in each row
- Clickable link to view the full-resolution card image
- Editable **Quantity** column so you can track how many copies you own
- Autofilters on Set, Faction, Rarity, Type, Cost, Recall Cost, and Power columns
- Frozen header row for easy scrolling

## Setup

```bash
python -m venv .venv
source .venv/bin/activate   # macOS / Linux
pip install -r requirements.txt
```

## Usage

```bash
python main.py
```

This will:

1. Fetch card data from `https://api.altered.gg/cards`
2. Download card thumbnails to `temp/`
3. Generate `altered_cards_catalogue.xlsx` in the project root

Open the `.xlsx` file in Excel, Google Sheets, or LibreOffice Calc to start cataloguing.

## Configuration

Edit `src/config.py` to change:

| Setting | Default | Description |
|---------|---------|-------------|
| `LANGUAGE` | `"en-us"` | Language for card names and images |
| `RARITIES` | `["COMMON", "RARE", "EXALTED"]` | Which rarities to include |
| `OUTPUT_FILENAME` | `"altered_cards_catalogue.xlsx"` | Output file name |
| `ITEMS_PER_PAGE` | `36` | API page size |
| `FACTIONS` | All 7 factions | Which factions to fetch |

## Disclaimer

All card names, images, artwork, and related intellectual property belong to [Equinox / Altered](https://www.altered.gg). This project is **not** affiliated with, endorsed by, or associated with Equinox in any way. Card data and images are retrieved from their public API purely for personal collection-tracking purposes. No ownership of any Altered TCG content is claimed.
