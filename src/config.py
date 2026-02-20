import os

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

API_BASE_URL = "https://api.altered.gg"
LANGUAGE = "en-us"
LANGUAGE_HEADER = {"Accept-Language": "en-us"}

RARITIES = ["COMMON", "RARE", "EXALTED"]
FACTIONS = ["AX", "BR", "LY", "MU", "OR", "YZ", "NE"]
ITEMS_PER_PAGE = 36
MAX_RETRIES = 5

TEMP_DIR = os.path.join(BASE_DIR, "temp")
OUTPUT_FILENAME = os.path.join(BASE_DIR, "altered_cards_catalogue.xlsx")

THUMB_WIDTH = 80
THUMB_HEIGHT = 110
