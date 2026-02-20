import io
import os

import requests
import xlsxwriter
from PIL import Image as PilImage

from src.config import TEMP_DIR, THUMB_HEIGHT, THUMB_WIDTH

COLUMNS = [
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
    ("Quantity", 12),
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

    header_fmt = wb.add_format(
        {
            "bold": True,
            "font_name": "Calibri",
            "font_size": 11,
            "font_color": "#FFFFFF",
            "bg_color": "#2F5496",
            "align": "center",
            "valign": "vcenter",
            "text_wrap": True,
            "border": 1,
            "border_color": "#D9D9D9",
        }
    )
    cell_fmt = wb.add_format(
        {
            "valign": "vcenter",
            "border": 1,
            "border_color": "#D9D9D9",
        }
    )
    alt_fmt = wb.add_format(
        {
            "valign": "vcenter",
            "border": 1,
            "border_color": "#D9D9D9",
            "bg_color": "#F2F6FC",
        }
    )
    link_fmt = wb.add_format(
        {
            "valign": "vcenter",
            "border": 1,
            "border_color": "#D9D9D9",
            "font_color": "#0563C1",
            "underline": True,
        }
    )
    alt_link_fmt = wb.add_format(
        {
            "valign": "vcenter",
            "border": 1,
            "border_color": "#D9D9D9",
            "font_color": "#0563C1",
            "underline": True,
            "bg_color": "#F2F6FC",
        }
    )
    qty_fmt = wb.add_format(
        {
            "valign": "vcenter",
            "align": "center",
            "border": 1,
            "border_color": "#D9D9D9",
            "num_format": "0",
        }
    )
    alt_qty_fmt = wb.add_format(
        {
            "valign": "vcenter",
            "align": "center",
            "border": 1,
            "border_color": "#D9D9D9",
            "num_format": "0",
            "bg_color": "#F2F6FC",
        }
    )
    num_fmt = wb.add_format(
        {
            "valign": "vcenter",
            "align": "center",
            "border": 1,
            "border_color": "#D9D9D9",
        }
    )
    alt_num_fmt = wb.add_format(
        {
            "valign": "vcenter",
            "align": "center",
            "border": 1,
            "border_color": "#D9D9D9",
            "bg_color": "#F2F6FC",
        }
    )

    for col_idx, (header, width) in enumerate(COLUMNS):
        ws.write(0, col_idx, header, header_fmt)
        ws.set_column(col_idx, col_idx, width)

    ws.set_row(0, 24)
    ws.freeze_panes(1, 0)

    last_col = len(COLUMNS) - 1
    ws.autofilter(0, 0, len(cards), last_col)

    total = len(cards)
    for idx, card in enumerate(cards):
        row = idx + 1
        if idx % 50 == 0:
            print(f"  Building row {row}/{total} ...")

        is_alt = idx % 2 == 1
        fmt = alt_fmt if is_alt else cell_fmt
        nfmt = alt_num_fmt if is_alt else num_fmt
        qfmt = alt_qty_fmt if is_alt else qty_fmt
        lfmt = alt_link_fmt if is_alt else link_fmt

        ws.set_row(row, ROW_HEIGHT_PT)

        thumb_path = _download_thumbnail(card["image_url"], card["reference"])
        if thumb_path:
            try:
                ws.insert_image(
                    row,
                    0,
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
        ws.write_blank(row, 0, "", fmt)

        ws.write(row, 1, card["name"], fmt)
        ws.write(row, 2, card["card_set"], fmt)
        ws.write(row, 3, card["faction"], fmt)
        ws.write(row, 4, card["rarity"], fmt)
        ws.write(row, 5, card["card_type"], fmt)

        ws.write(row, 6, _to_int(card["main_cost"]), nfmt)
        ws.write(row, 7, _to_int(card["recall_cost"]), nfmt)
        ws.write(row, 8, _to_int(card["mountain_power"]), nfmt)
        ws.write(row, 9, _to_int(card["ocean_power"]), nfmt)
        ws.write(row, 10, _to_int(card["forest_power"]), nfmt)

        ws.write(row, 11, 0, qfmt)

        if card["image_url"]:
            ws.write_url(row, 12, card["image_url"], lfmt, "View Full")
        else:
            ws.write_blank(row, 12, "", fmt)

    wb.close()
    print(f"Saved catalogue to {output_path}")


def _to_int(val: str):
    if val is None or val == "":
        return ""
    try:
        return int(val)
    except (ValueError, TypeError):
        return val
