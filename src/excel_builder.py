import io
import os

import requests
import xlsxwriter
from PIL import Image as PilImage

from src.config import TEMP_DIR, THUMB_HEIGHT, THUMB_WIDTH

COLUMNS = [
    ("Qty", 8),
    ("Thumbnail", 16),
    ("Name", 30),
    ("Set", 26),
    ("Faction", 16),
    ("Rarity", 14),
    ("Type", 16),
    ("Cost", 10),
    ("Recall Cost", 14),
    ("Mountain", 12),
    ("Ocean", 12),
    ("Forest", 12),
    ("Full Image", 16),
]

ROW_HEIGHT_PT = (THUMB_HEIGHT + 4) * 0.75


def _download_thumbnail(image_url: str, reference: str) -> str | None:
    if not image_url:
        return None

    os.makedirs(TEMP_DIR, exist_ok=True)
    thumb_path = os.path.join(TEMP_DIR, f"{reference}.png")

    if os.path.exists(thumb_path):
        return thumb_path

    try:
        resp = requests.get(image_url, timeout=30)
        resp.raise_for_status()
    except requests.RequestException:
        return None

    try:
        img = PilImage.open(io.BytesIO(resp.content))
        img.thumbnail((THUMB_WIDTH, THUMB_HEIGHT), PilImage.LANCZOS)
        img.save(thumb_path, "PNG")
        return thumb_path
    except Exception:
        return None


def build_catalogue(cards: list[dict], output_path: str) -> None:
    wb = xlsxwriter.Workbook(output_path, {"use_zip64": True})
    ws = wb.add_worksheet("Altered Cards")

    num_fmt = wb.add_format({"align": "center", "valign": "vcenter"})
    qty_fmt = wb.add_format(
        {"align": "center", "valign": "vcenter", "num_format": "0"}
    )
    link_fmt = wb.add_format(
        {
            "valign": "vcenter",
            "font_color": "#0563C1",
            "underline": True,
        }
    )
    cell_fmt = wb.add_format({"valign": "vcenter"})

    for col_idx, (_, width) in enumerate(COLUMNS):
        ws.set_column(col_idx, col_idx, width)

    total = len(cards)
    for idx, card in enumerate(cards):
        row = idx + 1
        if idx % 50 == 0:
            print(f"  Building row {row}/{total} ...")

        ws.set_row(row, ROW_HEIGHT_PT)

        ws.write(row, 0, 0, qty_fmt)

        thumb_path = _download_thumbnail(card["image_url"], card["reference"])
        if thumb_path:
            try:
                ws.insert_image(
                    row,
                    1,
                    thumb_path,
                    {
                        "x_scale": THUMB_WIDTH / PilImage.open(thumb_path).width,
                        "y_scale": THUMB_HEIGHT / PilImage.open(thumb_path).height,
                        "x_offset": 2,
                        "y_offset": 2,
                        "object_position": 1,
                    },
                )
            except Exception:
                pass

        ws.write(row, 2, card["name"], cell_fmt)
        ws.write(row, 3, card["card_set"], cell_fmt)
        ws.write(row, 4, card["faction"], cell_fmt)
        ws.write(row, 5, card["rarity"], cell_fmt)
        ws.write(row, 6, card["card_type"], cell_fmt)
        ws.write(row, 7, _to_int(card["main_cost"]), num_fmt)
        ws.write(row, 8, _to_int(card["recall_cost"]), num_fmt)
        ws.write(row, 9, _to_int(card["mountain_power"]), num_fmt)
        ws.write(row, 10, _to_int(card["ocean_power"]), num_fmt)
        ws.write(row, 11, _to_int(card["forest_power"]), num_fmt)

        if card["image_url"]:
            ws.write_url(row, 12, card["image_url"], link_fmt, "View Full")

    last_row = len(cards)
    last_col = len(COLUMNS) - 1
    table_columns = [{"header": name} for name, _ in COLUMNS]
    table_columns[0]["format"] = qty_fmt
    table_columns[12]["format"] = link_fmt
    for i in range(7, 12):
        table_columns[i]["format"] = num_fmt

    ws.add_table(
        0,
        0,
        last_row,
        last_col,
        {
            "columns": table_columns,
            "style": "Table Style Medium 2",
            "banded_rows": True,
            "autofilter": True,
            "name": "AlteredCards",
        },
    )

    ws.freeze_panes(1, 1)

    wb.close()
    print(f"Saved catalogue to {output_path}")


def _to_int(val: str):
    if val is None or val == "":
        return ""
    try:
        return int(val)
    except (ValueError, TypeError):
        return val
