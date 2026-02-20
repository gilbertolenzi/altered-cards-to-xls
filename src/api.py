import time
import requests

from src.config import (
    API_BASE_URL,
    FACTIONS,
    ITEMS_PER_PAGE,
    LANGUAGE_HEADER,
    MAX_RETRIES,
    RARITIES,
)


def _build_url(page: int, faction: str) -> str:
    rarity_params = "&".join(f"rarity[]={r}" for r in RARITIES)
    return (
        f"{API_BASE_URL}/cards"
        f"?{rarity_params}"
        f"&factions[]={faction}"
        f"&itemsPerPage={ITEMS_PER_PAGE}"
        f"&page={page}"
    )


def _get_page(page: int, faction: str) -> tuple[list[dict], int]:
    url = _build_url(page, faction)
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            resp = requests.get(url, headers=LANGUAGE_HEADER, timeout=30)
            resp.raise_for_status()
            data = resp.json()
            return data["hydra:member"], data["hydra:totalItems"]
        except (requests.RequestException, KeyError) as exc:
            if attempt == MAX_RETRIES:
                raise RuntimeError(
                    f"Failed to fetch {url} after {MAX_RETRIES} attempts"
                ) from exc
            wait = 2 ** attempt
            print(f"  Retry {attempt}/{MAX_RETRIES} for {url} (waiting {wait}s)...")
            time.sleep(wait)
    return [], 0  # unreachable, keeps type-checkers happy


def _flatten_card(raw: dict) -> dict:
    elements = raw.get("elements", {})
    image_paths = raw.get("allImagePath", {})
    image_url = image_paths.get("en-us", raw.get("imagePath", ""))

    return {
        "reference": raw.get("reference", ""),
        "name": raw.get("name", ""),
        "card_type": raw.get("cardType", {}).get("name", ""),
        "card_set": raw.get("cardSet", {}).get("name", ""),
        "card_set_ref": raw.get("cardSet", {}).get("reference", ""),
        "faction": raw.get("mainFaction", {}).get("name", ""),
        "faction_ref": raw.get("mainFaction", {}).get("reference", ""),
        "rarity": raw.get("rarity", {}).get("name", ""),
        "main_cost": elements.get("MAIN_COST", ""),
        "recall_cost": elements.get("RECALL_COST", ""),
        "mountain_power": elements.get("MOUNTAIN_POWER", ""),
        "ocean_power": elements.get("OCEAN_POWER", ""),
        "forest_power": elements.get("FOREST_POWER", ""),
        "image_url": image_url,
    }


def fetch_all_cards() -> list[dict]:
    """Fetch all Common, Rare, and Exalted cards across every faction."""
    all_cards: list[dict] = []

    for faction in FACTIONS:
        print(f"Fetching faction {faction} ...")
        page = 1
        members, total = _get_page(page, faction)
        cards_so_far = list(members)

        total_pages = max(1, -(-total // ITEMS_PER_PAGE))  # ceil division
        print(f"  {total} cards, {total_pages} page(s)")

        for page in range(2, total_pages + 1):
            members, _ = _get_page(page, faction)
            cards_so_far.extend(members)
            print(f"  page {page}/{total_pages}")

        all_cards.extend(_flatten_card(c) for c in cards_so_far)

    print(f"\nTotal cards fetched: {len(all_cards)}")
    return all_cards
