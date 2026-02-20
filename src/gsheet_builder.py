import math
import os

import gspread
from google.oauth2.service_account import Credentials

from src.config import GOOGLE_CREDENTIALS_FILE, GOOGLE_SHEET_NAME

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

HEADERS = [
    "Thumbnail",
    "Name",
    "Set",
    "Faction",
    "Rarity",
    "Type",
    "Cost",
    "Recall Cost",
    "Mountain",
    "Ocean",
    "Forest",
    "Quantity",
    "Full Image",
]

COL_WIDTHS = [120, 220, 200, 120, 100, 120, 70, 100, 85, 85, 85, 85, 110]

BATCH_SIZE = 500


def _authenticate() -> gspread.Client:
    creds_path = GOOGLE_CREDENTIALS_FILE
    if not os.path.exists(creds_path):
        raise FileNotFoundError(
            f"Google credentials not found at {creds_path}\n"
            "See README for setup instructions."
        )
    creds = Credentials.from_service_account_file(creds_path, scopes=SCOPES)
    return gspread.authorize(creds)


def _card_to_row(card: dict, row_num: int) -> list:
    """Convert a card dict to a spreadsheet row (1-indexed row_num for formulas)."""
    image_url = card["image_url"]
    thumb_formula = f'=IMAGE("{image_url}", 2)' if image_url else ""
    link_formula = f'=HYPERLINK("{image_url}", "View Full")' if image_url else ""

    return [
        thumb_formula,
        card["name"],
        card["card_set"],
        card["faction"],
        card["rarity"],
        card["card_type"],
        _to_int(card["main_cost"]),
        _to_int(card["recall_cost"]),
        _to_int(card["mountain_power"]),
        _to_int(card["ocean_power"]),
        _to_int(card["forest_power"]),
        0,
        link_formula,
    ]


def build_gsheet(cards: list[dict]) -> str:
    """Create a Google Sheet catalogue. Returns the spreadsheet URL."""
    print("  Authenticating with Google ...")
    gc = _authenticate()

    print(f"  Creating spreadsheet '{GOOGLE_SHEET_NAME}' ...")
    sh = gc.create(GOOGLE_SHEET_NAME)
    ws = sh.sheet1
    ws.update_title("Altered Cards")

    total = len(cards)
    total_rows = total + 1  # header + data
    total_cols = len(HEADERS)

    ws.resize(rows=total_rows, cols=total_cols)

    print("  Writing headers ...")
    ws.update("A1", [HEADERS], value_input_option="RAW")

    print(f"  Writing {total} card rows ...")
    num_batches = math.ceil(total / BATCH_SIZE)
    for batch_idx in range(num_batches):
        start = batch_idx * BATCH_SIZE
        end = min(start + BATCH_SIZE, total)
        batch_rows = [
            _card_to_row(cards[i], i + 2) for i in range(start, end)
        ]
        start_row = start + 2
        end_row = end + 1
        cell_range = f"A{start_row}:M{end_row}"
        ws.update(cell_range, batch_rows, value_input_option="USER_ENTERED")
        print(f"    Batch {batch_idx + 1}/{num_batches} ({end}/{total})")

    print("  Applying formatting ...")
    sheet_id = ws.id

    requests = []

    # Freeze header row
    requests.append({
        "updateSheetProperties": {
            "properties": {
                "sheetId": sheet_id,
                "gridProperties": {"frozenRowCount": 1},
            },
            "fields": "gridProperties.frozenRowCount",
        }
    })

    # Header formatting
    requests.append({
        "repeatCell": {
            "range": {
                "sheetId": sheet_id,
                "startRowIndex": 0,
                "endRowIndex": 1,
                "startColumnIndex": 0,
                "endColumnIndex": total_cols,
            },
            "cell": {
                "userEnteredFormat": {
                    "backgroundColor": {"red": 0.184, "green": 0.329, "blue": 0.588},
                    "textFormat": {
                        "bold": True,
                        "fontSize": 11,
                        "foregroundColor": {"red": 1, "green": 1, "blue": 1},
                    },
                    "horizontalAlignment": "CENTER",
                    "verticalAlignment": "MIDDLE",
                }
            },
            "fields": "userEnteredFormat",
        }
    })

    # Row height for thumbnail rows
    requests.append({
        "updateDimensionProperties": {
            "range": {
                "sheetId": sheet_id,
                "dimension": "ROWS",
                "startIndex": 1,
                "endIndex": total_rows,
            },
            "properties": {"pixelSize": 120},
            "fields": "pixelSize",
        }
    })

    # Column widths
    for col_idx, width in enumerate(COL_WIDTHS):
        requests.append({
            "updateDimensionProperties": {
                "range": {
                    "sheetId": sheet_id,
                    "dimension": "COLUMNS",
                    "startIndex": col_idx,
                    "endIndex": col_idx + 1,
                },
                "properties": {"pixelSize": width},
                "fields": "pixelSize",
            }
        })

    # Center-align numeric columns (Cost through Forest = cols 6-10, Quantity = 11)
    for col_idx in [6, 7, 8, 9, 10, 11]:
        requests.append({
            "repeatCell": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": 1,
                    "endRowIndex": total_rows,
                    "startColumnIndex": col_idx,
                    "endColumnIndex": col_idx + 1,
                },
                "cell": {
                    "userEnteredFormat": {
                        "horizontalAlignment": "CENTER",
                        "verticalAlignment": "MIDDLE",
                    }
                },
                "fields": "userEnteredFormat.horizontalAlignment,"
                          "userEnteredFormat.verticalAlignment",
            }
        })

    # Basic filter on all columns
    requests.append({
        "setBasicFilter": {
            "filter": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": 0,
                    "endRowIndex": total_rows,
                    "startColumnIndex": 0,
                    "endColumnIndex": total_cols,
                },
            }
        }
    })

    sh.batch_update({"requests": requests})

    # Make it accessible to anyone with the link
    sh.share(None, perm_type="anyone", role="writer")

    url = sh.url
    print(f"  Spreadsheet created: {url}")
    return url


def _to_int(val: str):
    if val is None or val == "":
        return ""
    try:
        return int(val)
    except (ValueError, TypeError):
        return val
