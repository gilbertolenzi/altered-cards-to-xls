# Altered TCG Cards to Excel / Google Sheets Catalogue

Fetches **Common**, **Rare**, and **Exalted** cards from the [Altered.gg](https://www.altered.gg) public API and generates a catalogue for tracking your collection.

Two output formats are supported:

| Format | File | Filters | Images |
|--------|------|---------|--------|
| **Excel** (.xlsx) | Local file | Table with sort & filter dropdowns | Embedded thumbnails |
| **Google Sheets** | Shareable URL | Native filter dropdowns on every column | `=IMAGE()` loaded from URL |

## Quick Start

```bash
git clone https://github.com/gilbertolenzi/altered-cards-to-xls.git
cd altered-cards-to-xls
python -m venv .venv
source .venv/bin/activate   # macOS / Linux
pip install -r requirements.txt
```

### Generate the Excel file only (no Google setup needed)

```bash
python main.py --excel
```

Output: `altered_cards_catalogue.xlsx` in the project root.

### Generate the Google Sheet only

```bash
python main.py --gsheet
```

Requires a `credentials.json` service-account file (see [Google Sheets Setup](#google-sheets-setup) below).

### Generate both at once

```bash
python main.py
```

Running with no flags produces both outputs.

## Features

- Pulls all Common, Rare, and Exalted cards across every faction
- **Excel**: Proper Excel Table with built-in sort/filter dropdowns, banded rows, embedded card thumbnails, and clickable "View Full" image links
- **Google Sheets**: `=IMAGE()` thumbnails rendered in-cell, `=HYPERLINK()` to full card image, native filter dropdowns, frozen header, formatted columns -- shared as an editable link anyone can open
- Editable **Quantity** column (defaults to 0) for collection tracking
- Filters on Set, Faction, Rarity, Type, Cost, Recall Cost, and Power columns

## Google Sheets Setup

To use the `--gsheet` option you need a Google Cloud service account. This is a one-time setup:

1. Go to the [Google Cloud Console](https://console.cloud.google.com/)
2. Create a new project (or use an existing one)
3. Enable the **Google Sheets API** and **Google Drive API**:
   - Navigate to **APIs & Services > Library**
   - Search for "Google Sheets API" and click **Enable**
   - Search for "Google Drive API" and click **Enable**
4. Create a service account:
   - Go to **APIs & Services > Credentials**
   - Click **Create Credentials > Service account**
   - Give it a name and click **Done**
5. Create a key:
   - Click on the service account you just created
   - Go to the **Keys** tab
   - Click **Add Key > Create new key > JSON**
   - Save the downloaded file as `credentials.json` in the project root

> `credentials.json` is listed in `.gitignore` and will never be committed.

## Configuration

Edit `src/config.py` to change:

| Setting | Default | Description |
|---------|---------|-------------|
| `LANGUAGE` | `"en-us"` | Language for card names and images |
| `RARITIES` | `["COMMON", "RARE", "EXALTED"]` | Which rarities to include |
| `OUTPUT_FILENAME` | `"altered_cards_catalogue.xlsx"` | Excel output file name |
| `ITEMS_PER_PAGE` | `36` | API page size |
| `FACTIONS` | All 7 factions | Which factions to fetch |
| `GOOGLE_SHEET_NAME` | `"Altered TCG Catalogue"` | Name of the created Google Sheet |

## Disclaimer

All card names, images, artwork, and related intellectual property belong to [Equinox / Altered](https://www.altered.gg). This project is **not** affiliated with, endorsed by, or associated with Equinox in any way. Card data and images are retrieved from their public API purely for personal collection-tracking purposes. No ownership of any Altered TCG content is claimed.
