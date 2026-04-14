#!/usr/bin/env python3
"""
MTG Packing Slip Organizer
Parses TCGplayer packing slips and reorganizes cards by color and rarity.
Uses Scryfall API to look up card colors.
"""

import re
import time
import json
import urllib.request
import urllib.parse
import urllib.error
from pathlib import Path
from dataclasses import dataclass
from collections import defaultdict
from typing import Optional

# Try to import pdfplumber, provide helpful error if not available
try:
    import pdfplumber
except ImportError:
    print("Error: pdfplumber is required. Install it with:")
    print("  python3 -m venv venv && source venv/bin/activate && pip install pdfplumber")
    exit(1)


@dataclass
class Card:
    quantity: int
    set_name: str
    card_name: str
    collector_number: str
    rarity: str
    condition: str
    is_foil: bool
    price: float
    total_price: float
    variant: Optional[str] = None
    language: Optional[str] = None  # Non-English language if applicable
    color: str = "Colorless"  # Will be populated from Scryfall
    image_url: Optional[str] = None  # Scryfall card image URL
    order_group: Optional[str] = None  # 'A', 'B', or 'C' for multi-order pulls


# Rarity mapping
RARITY_ORDER = {"M": 0, "R": 1, "U": 2, "C": 3, "S": 4}
RARITY_NAMES = {"M": "Mythic Rare", "R": "Rare", "U": "Uncommon", "C": "Common", "S": "Special"}

# Order group colors for multi-order pull sheets (max 3 concurrent orders)
ORDER_GROUP_COLORS = {
    'A': '#4a9eff',  # blue
    'B': '#ff9800',  # orange
    'C': '#66bb6a',  # green
}

# Variant/border treatment styling — maps lowercase keyword to (css_class, label, color)
VARIANT_STYLES = {
    'borderless': ('variant-borderless', 'Borderless', '#00bcd4'),
    'extended art': ('variant-extended', 'Extended Art', '#9c27b0'),
    'showcase': ('variant-showcase', 'Showcase', '#e91e63'),
    'retro frame': ('variant-retro', 'Retro Frame', '#ff8f00'),
    'white border': ('variant-white-border', 'White Border', '#b0bec5'),
    'foil etched': ('variant-etched', 'Foil Etched', '#b8860b'),
    'full art': ('variant-full-art', 'Full Art', '#3f51b5'),
    'future sight': ('variant-future-sight', 'Future Sight', '#00e5ff'),
    'surge foil': ('variant-surge', 'Surge Foil', '#ff6f00'),
}


def get_variant_style(variant: Optional[str]) -> tuple[str, str, str]:
    """Return (css_class, label, color) for a variant, or a generic fallback."""
    if not variant:
        return ('', '', '')
    lower = variant.lower()
    for key, style in VARIANT_STYLES.items():
        if key in lower:
            return style
    # Generic fallback for unknown variants
    return ('variant-other', variant, '#78909c')

# Color order for sorting (WUBRG + multicolor + colorless + land)
COLOR_ORDER = {
    "White": 0,
    "Blue": 1,
    "Black": 2,
    "Red": 3,
    "Green": 4,
    "Multicolor": 5,
    "Colorless": 6,
    "Land": 7,
}

# TCGPlayer-specific set name overrides that don't match Scryfall naming.
# These supplement the dynamically-fetched Scryfall set list.
TCGPLAYER_SET_OVERRIDES = {
    "SecretLairDropSeries": "sld",
    "Secret Lair Drop Series": "sld",
    "SecretLairCountdownKit": "slc",
    "Secret Lair Countdown Kit": "slc",
    "Avatar:TheLastAirbender:Eternal-Legal": "tle",
    "Avatar: The Last Airbender: Eternal-Legal": "tle",
    "MarvelUniverseEternal-Legal": "mar",
    "Marvel Universe Eternal-Legal": "mar",
    "TheListReprints": "plst",
    "The List Reprints": "plst",
    "TimeSpiral:Remastered": "tsr",
    "Time Spiral: Remastered": "tsr",
}

# Global cache for Scryfall set mapping (populated on first use)
_scryfall_set_cache: Optional[dict] = None
# Global cache for set prefixes used in PDF parsing (populated alongside set cache)
_set_prefix_cache: Optional[list] = None
# Track how many sets were loaded (for diagnostics)
_scryfall_set_count: int = 0
# Track the latest sets by release date (for diagnostics)
_latest_sets: list[dict] = []


def fetch_scryfall_sets() -> dict:
    """
    Fetch all sets from Scryfall API and build a mapping from various name formats
    to set codes. Also builds the set prefix list used for PDF parsing.
    This ensures new sets are always supported without code changes.
    """
    global _scryfall_set_cache, _set_prefix_cache, _scryfall_set_count, _latest_sets

    if _scryfall_set_cache is not None:
        return _scryfall_set_cache

    print("Syncing set list from Scryfall...")

    try:
        url = "https://api.scryfall.com/sets"
        req = urllib.request.Request(url, headers={"User-Agent": "MTG-Packing-Slip-Organizer/1.0", "Accept": "application/json"})
        with urllib.request.urlopen(req, timeout=30) as response:
            data = json.loads(response.read().decode())

        mapping = {}
        prefixes = set()
        sets_data = data.get("data", [])
        _scryfall_set_count = len(sets_data)

        # Track the 5 most recently released sets (by release_at date)
        sorted_by_release = sorted(
            [s for s in sets_data if s.get("released_at")],
            key=lambda s: s["released_at"],
            reverse=True,
        )
        _latest_sets = [
            {"name": s["name"], "code": s["code"], "released_at": s["released_at"]}
            for s in sorted_by_release[:5]
        ]

        for set_info in sets_data:
            code = set_info.get("code", "")
            name = set_info.get("name", "")

            if not code or not name:
                continue

            # Skip art series, tokens, promos, etc. for the primary mapping
            # These have different collector numbers than the main sets
            set_type = set_info.get("set_type", "")
            if set_type in ("token", "memorabilia", "promo", "alchemy"):
                continue

            # Add the exact name
            mapping[name] = code

            # Add version without spaces (TCGPlayer PDF format)
            no_spaces = name.replace(" ", "")
            mapping[no_spaces] = code
            prefixes.add(no_spaces)

            # Add version with colons but no spaces around them
            colon_no_space = name.replace(": ", ":")
            mapping[colon_no_space] = code
            prefixes.add(colon_no_space)

            # Add version with no spaces at all (including around colons)
            all_no_spaces = name.replace(" ", "").replace(":", "")
            if all_no_spaces not in mapping:
                mapping[all_no_spaces] = code
            prefixes.add(all_no_spaces)

        # Add TCGPlayer-specific overrides
        mapping.update(TCGPLAYER_SET_OVERRIDES)
        prefixes.update(TCGPLAYER_SET_OVERRIDES.keys())

        _scryfall_set_cache = mapping
        _set_prefix_cache = sorted(prefixes, key=len, reverse=True)
        print(f"  Loaded {_scryfall_set_count} sets from Scryfall ({len(prefixes)} parsing prefixes)")
        return mapping

    except Exception as e:
        print(f"  Warning: Could not fetch sets from Scryfall: {e}")
        print("  Falling back to basic name matching")
        _scryfall_set_cache = {}
        _set_prefix_cache = sorted(TCGPLAYER_SET_OVERRIDES.keys(), key=len, reverse=True)
        return {}


def get_set_prefixes() -> list[str]:
    """
    Return the list of known set name prefixes for PDF parsing.
    Dynamically built from Scryfall data on first call.
    """
    if _set_prefix_cache is None:
        fetch_scryfall_sets()
    return _set_prefix_cache or []


def get_set_sync_status() -> dict:
    """Return diagnostic info about the current set data."""
    if _scryfall_set_cache is None:
        fetch_scryfall_sets()
    return {
        "sets_loaded": _scryfall_set_count,
        "prefix_count": len(_set_prefix_cache) if _set_prefix_cache else 0,
        "tcgplayer_overrides": len(TCGPLAYER_SET_OVERRIDES),
        "cache_populated": _scryfall_set_cache is not None and len(_scryfall_set_cache) > 0,
        "latest_sets": _latest_sets,
    }


def get_scryfall_set_code(set_name: str) -> Optional[str]:
    """Get the Scryfall set code for a TCGPlayer set name."""
    mapping = fetch_scryfall_sets()

    # Try exact match first
    if set_name in mapping:
        return mapping[set_name]

    # Try without spaces
    no_spaces = set_name.replace(" ", "")
    if no_spaces in mapping:
        return mapping[no_spaces]

    # Try lowercase
    lower = set_name.lower()
    for key, code in mapping.items():
        if key.lower() == lower:
            return code

    return None


def extract_text_from_pdf(pdf_path: str) -> str:
    """Extract all text from a PDF file."""
    text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text()
            if page_text:
                text += page_text + "\n"
    return text


def add_spaces_to_card_name(name: str) -> str:
    """
    Add spaces to a card name that has been concatenated without spaces.
    E.g., "EloquentFirst-Year" -> "Eloquent First-Year"
    E.g., "Abigale,EloquentFirst-Year" -> "Abigale, Eloquent First-Year"
    """
    if not name:
        return name

    # First, handle special cases and preserve intentional concatenations
    # Don't add space before apostrophe-s
    # Don't add space in the middle of hyphenated words

    result = []
    i = 0
    while i < len(name):
        char = name[i]

        # Check if we need to add a space before this character
        if i > 0:
            prev_char = name[i - 1]

            # Add space between lowercase and uppercase (camelCase)
            # But not if prev char is an apostrophe or we're at start of hyphenated word
            if (prev_char.islower() and char.isupper() and
                prev_char != "'" and
                (i < 2 or name[i-2] != '-')):
                result.append(' ')

            # Add space between letter and number in some cases
            # e.g., "Momo2" but NOT for things like "C-3PO"

        result.append(char)
        i += 1

    name = ''.join(result)

    # Add space after comma if not present
    name = re.sub(r',([^\s])', r', \1', name)

    # Add space after colon if not present (but not for things like "1:1")
    name = re.sub(r':([A-Za-z])', r': \1', name)

    # Handle concatenated words with prepositions (from PDF no-space format)
    # Common word endings before "of": -ion, -ter, -ath, -lic, etc.
    name = re.sub(r'(ion|ter|ler|ant|ent|int|ard|ack|ock|uck|ime|ame|ome|ple|tle|nce|ise|ose|use|ine|one|ure|ire|are|ore|ide|ade|ude|ive|ave|ove|all|ell|ill|ull|ath|eth|ith|oth|uth|lic|ric|sic|tic|nic|pic)of', r'\1 of', name, flags=re.IGNORECASE)

    # Handle "sof" pattern (Ripples of, Champions of, etc.)
    name = re.sub(r'([^o])sof', r'\1s of', name, flags=re.IGNORECASE)

    # Handle similar patterns for "to"
    name = re.sub(r'(ack|ome|urn)to', r'\1 to', name, flags=re.IGNORECASE)

    # Fix "ofthe" "tothe" etc. patterns
    name = re.sub(r'\bof ?the\b', 'of the', name, flags=re.IGNORECASE)
    name = re.sub(r'\bto ?the\b', 'to the', name, flags=re.IGNORECASE)
    name = re.sub(r'\bat ?the\b', 'at the', name, flags=re.IGNORECASE)
    name = re.sub(r'\bin ?the\b', 'in the', name, flags=re.IGNORECASE)
    name = re.sub(r'\bfor ?the\b', 'for the', name, flags=re.IGNORECASE)
    name = re.sub(r'\bfrom ?the\b', 'from the', name, flags=re.IGNORECASE)
    name = re.sub(r'\bon ?the\b', 'on the', name, flags=re.IGNORECASE)
    name = re.sub(r'\band ?the\b', 'and the', name, flags=re.IGNORECASE)

    # Handle "the" followed by capital (thePerished -> the Perished)
    name = re.sub(r'\bthe([A-Z])', r'the \1', name)

    # Fix "First-Year" type patterns - ensure space before if preceded by lowercase
    # "EloquentFirst-Year" should become "Eloquent First-Year"

    # Clean up any double spaces
    name = re.sub(r'\s+', ' ', name)

    return name.strip()


def extract_set_and_card(description: str) -> tuple[str, str]:
    """
    Extract the set name and card name from the description.
    Returns (set_name, remaining_description)
    """
    # Use dynamically-built prefix list (already sorted longest-first)
    for prefix in get_set_prefixes():
        if description.startswith(prefix + "-"):
            return (prefix, description[len(prefix) + 1:])

    # Fallback: try to find the set name by looking for the pattern
    # SetName-CardName-#Number
    # The set name typically doesn't contain -#

    # Find the first occurrence of -# which marks the collector number
    number_match = re.search(r'-#\d+', description)
    if number_match:
        before_number = description[:number_match.start()]
        # The set name is everything up to the last hyphen before a capitalized word
        # that looks like a card name

        # Try to split on hyphen and find where the card name starts
        parts = before_number.split('-')
        if len(parts) >= 2:
            # Heuristic: the set name usually ends when we hit a part that
            # looks like it starts a card name (starts with capital, not a known set word)
            set_parts = []
            card_parts = []
            found_card = False

            for j, part in enumerate(parts):
                if not found_card:
                    # Check if this looks like it could be start of a card name
                    # Card names often start with capital letters and are not
                    # common set name components
                    if (j > 0 and part and part[0].isupper() and
                        not any(part.startswith(x) for x in ['Commander', 'Eternal', 'Legal', 'Remastered', 'Promo'])):
                        found_card = True
                        card_parts.append(part)
                    else:
                        set_parts.append(part)
                else:
                    card_parts.append(part)

            if set_parts and card_parts:
                return ('-'.join(set_parts), '-'.join(card_parts) + description[number_match.start():])

    # Last resort: split on first hyphen
    first_hyphen = description.find('-')
    if first_hyphen > 0:
        return (description[:first_hyphen], description[first_hyphen + 1:])

    return ("Unknown", description)


def parse_card_line(line: str) -> Optional[Card]:
    """Parse a single line from the packing slip into a Card object."""
    line = line.strip()

    # Skip non-card lines
    if not re.match(r'^\d+\s+Magic-', line):
        return None
    if "Quantity" in line or "Description" in line:
        return None

    # Extract quantity (first number)
    qty_match = re.match(r"^(\d+)\s+", line)
    if not qty_match:
        return None
    quantity = int(qty_match.group(1))

    # Extract prices (last two dollar amounts)
    price_pattern = r"\$(\d+\.?\d*)"
    prices = re.findall(price_pattern, line)
    if len(prices) >= 2:
        price = float(prices[-2])
        total_price = float(prices[-1])
    else:
        price = 0.0
        total_price = 0.0

    # Remove prices from line to get clean description
    # This handles cases where info appears AFTER prices due to line wrapping
    line_no_prices = re.sub(r'\$\d+\.?\d*', '', line)

    # Extract everything after "Magic-"
    magic_match = re.search(r'Magic-(.+)', line_no_prices)
    if not magic_match:
        return None
    full_description = magic_match.group(1).strip()

    # Extract set name and the rest
    set_name, remainder = extract_set_and_card(full_description)

    # Now parse the remainder: CardName(Variant)-#Number-Rarity-Condition
    # The full line (with info that may have been after prices) is now in remainder

    # Extract collector number - look for #XXX pattern anywhere
    number_match = re.search(r'#(\d+)', remainder)
    collector_number = number_match.group(1) if number_match else ""

    # Extract rarity - single letter M/R/U/C/S preceded by hyphen
    # Look for pattern like -R-, -M-, -R-Near, etc.
    rarity_match = re.search(r'-([MRUCS])-', remainder)
    if not rarity_match:
        # Try at end of string or before condition
        rarity_match = re.search(r'-([MRUCS])(?:$|-?Near|-?Lightly|-?Moderately|-?Heavily|-?Foil)', remainder)
    if not rarity_match:
        # Try rarity followed by space or price (e.g., "-M $5.81")
        rarity_match = re.search(r'-([MRUCS])[\s\$]', line)
    if not rarity_match:
        # Handle continuation lines where rarity appears after price (e.g., "$1.70M-NearMint")
        # Search the full line for price followed by rarity
        rarity_match = re.search(r'\$\d+\.?\d*([MRUCS])-', line)
    rarity = rarity_match.group(1) if rarity_match else "R"

    # Check for foil - can appear anywhere in the line
    is_foil = "Foil" in line  # Check original line, not just remainder

    # Extract condition
    condition = "Near Mint"
    if "LightlyPlayed" in line or "Lightly Played" in line:
        condition = "Lightly Played"
    elif "ModeratelyPlayed" in line or "Moderately Played" in line:
        condition = "Moderately Played"
    elif "HeavilyPlayed" in line or "Heavily Played" in line:
        condition = "Heavily Played"

    # Extract language - check for non-English languages in the line
    # TCGPlayer includes language as part of the product description
    # Be careful to avoid false positives (e.g., "RetroFrame" matching "FRA",
    # "Titan" matching "ITA")
    language = None
    # Use word boundaries or hyphen boundaries for language detection
    # TCGPlayer format typically has language as a separate field like "-Japanese-" or at end
    language_patterns = [
        # Full language names with word boundaries
        (r'[-\s]Japanese[-\s]|Japanese$', 'Japanese'),
        (r'[-\s]German[-\s]|German$', 'German'),
        (r'[-\s]French[-\s]|French$', 'French'),
        (r'[-\s]Italian[-\s]|Italian$', 'Italian'),
        (r'[-\s]Spanish[-\s]|Spanish$', 'Spanish'),
        (r'[-\s]Portuguese[-\s]|Portuguese$', 'Portuguese'),
        (r'[-\s]Russian[-\s]|Russian$', 'Russian'),
        (r'[-\s]Korean[-\s]|Korean$', 'Korean'),
        (r'ChineseSimplified|SimplifiedChinese', 'Chinese (Simplified)'),
        (r'ChineseTraditional|TraditionalChinese', 'Chinese (Traditional)'),
        (r'[-\s]Phyrexian[-\s]|Phyrexian$', 'Phyrexian'),
    ]
    for pattern, lang_name in language_patterns:
        if re.search(pattern, line):
            language = lang_name
            break

    # Extract card name - everything before the collector number
    if number_match:
        # Find position of #NUMBER in remainder
        num_pos = remainder.find('#' + collector_number)
        card_name_part = remainder[:num_pos]
        # Remove trailing hyphen
        card_name_part = card_name_part.rstrip('-')
    else:
        # Fallback: take everything before rarity
        if rarity_match:
            card_name_part = remainder[:rarity_match.start()]
            card_name_part = card_name_part.rstrip('-')
        else:
            card_name_part = remainder.split('-')[0]

    # Extract variant if present (in parentheses)
    variant = None
    variant_match = re.search(r'\(([^)]+)\)', card_name_part)
    if variant_match:
        variant = variant_match.group(1)
        # Clean up variant name - add spaces (e.g., "ExtendedArt" -> "Extended Art")
        variant = add_spaces_to_card_name(variant)
        card_name = card_name_part[:variant_match.start()].strip()
    else:
        card_name = card_name_part

    # Clean up card name - add proper spacing
    card_name = add_spaces_to_card_name(card_name)

    # Also clean up set name for display
    set_name_display = add_spaces_to_card_name(set_name)
    # Fix specific set name formatting
    set_name_display = set_name_display.replace("FINALFANTASY", "FINAL FANTASY")
    set_name_display = re.sub(r'Phyrexia:All Will Be One', 'Phyrexia: All Will Be One', set_name_display)
    set_name_display = re.sub(r'Avatar:The Last Airbender', 'Avatar: The Last Airbender', set_name_display)
    set_name_display = re.sub(r'Tarkir:Dragonstorm', 'Tarkir: Dragonstorm', set_name_display)
    set_name_display = re.sub(r'Commander:', 'Commander: ', set_name_display)
    set_name_display = re.sub(r'Promo Pack:', 'Promo Pack: ', set_name_display)
    set_name_display = re.sub(r'From the Vault:', 'From the Vault: ', set_name_display)
    set_name_display = re.sub(r'\s+', ' ', set_name_display)  # Clean double spaces

    return Card(
        quantity=quantity,
        set_name=set_name_display,
        card_name=card_name,
        collector_number=collector_number,
        rarity=rarity,
        condition=condition,
        is_foil=is_foil,
        price=price,
        total_price=total_price,
        variant=variant,
        language=language,
    )


def search_scryfall_by_set(set_code: str, collector_number: str) -> Optional[dict]:
    """Look up a specific card printing by set code and collector number."""
    # Scryfall endpoint for exact set/number lookup
    url = f"https://api.scryfall.com/cards/{set_code}/{collector_number}"

    try:
        req = urllib.request.Request(url, headers={"User-Agent": "MTG-Packing-Slip-Organizer/1.0", "Accept": "application/json"})
        with urllib.request.urlopen(req, timeout=10) as response:
            data = json.loads(response.read().decode())
            return data
    except urllib.error.HTTPError:
        return None
    except Exception:
        return None


def search_scryfall(card_name: str, set_name: str = None, collector_number: str = None) -> Optional[dict]:
    """Search Scryfall for a card and return its data.

    If set_name and collector_number are provided, tries to look up the exact printing first.
    Falls back to fuzzy name search if exact lookup fails.
    """
    # First, try exact lookup by set code and collector number if available
    if set_name and collector_number:
        # Try to map TCGPlayer set name to Scryfall set code
        set_code = get_scryfall_set_code(set_name)

        if set_code:
            result = search_scryfall_by_set(set_code, collector_number)
            if result:
                return result

    # Fall back to fuzzy name search
    search_name = card_name.strip()

    # Remove any remaining artifacts
    search_name = re.sub(r'\s+', ' ', search_name)

    # Use the search endpoint with fuzzy matching
    base_url = "https://api.scryfall.com/cards/named"
    params = {"fuzzy": search_name}

    url = f"{base_url}?{urllib.parse.urlencode(params)}"

    try:
        req = urllib.request.Request(url, headers={"User-Agent": "MTG-Packing-Slip-Organizer/1.0", "Accept": "application/json"})
        with urllib.request.urlopen(req, timeout=10) as response:
            data = json.loads(response.read().decode())
            return data
    except urllib.error.HTTPError as e:
        if e.code == 404:
            # Try exact search
            params = {"exact": search_name}
            url = f"{base_url}?{urllib.parse.urlencode(params)}"
            try:
                req = urllib.request.Request(url, headers={"User-Agent": "MTG-Packing-Slip-Organizer/1.0", "Accept": "application/json"})
                with urllib.request.urlopen(req, timeout=10) as response:
                    data = json.loads(response.read().decode())
                    return data
            except:
                pass

            # Try with simpler name (remove commas and extra words)
            simple_name = search_name.split(',')[0].strip()
            if simple_name != search_name:
                params = {"fuzzy": simple_name}
                url = f"{base_url}?{urllib.parse.urlencode(params)}"
                try:
                    req = urllib.request.Request(url, headers={"User-Agent": "MTG-Packing-Slip-Organizer/1.0", "Accept": "application/json"})
                    with urllib.request.urlopen(req, timeout=10) as response:
                        data = json.loads(response.read().decode())
                        return data
                except:
                    pass
        return None
    except Exception as e:
        print(f"  Error searching for '{search_name}': {e}")
        return None


def get_card_color(scryfall_data: dict) -> str:
    """Determine the color category from Scryfall data.

    For double-faced cards (transform, modal, etc.), we use the FRONT face
    to determine classification, since that's how you'll find it in your collection.
    """
    if not scryfall_data:
        return "Colorless"

    # For double-faced cards, use the front face for type and color determination
    if "card_faces" in scryfall_data and len(scryfall_data["card_faces"]) > 0:
        front_face = scryfall_data["card_faces"][0]
        type_line = front_face.get("type_line", "")
        colors = front_face.get("colors", [])
    else:
        type_line = scryfall_data.get("type_line", "")
        colors = scryfall_data.get("colors", [])

    # Check if it's a land first (front face only)
    if "Land" in type_line and "Creature" not in type_line:
        return "Land"

    # Handle case where colors is None (some Scryfall responses)
    if colors is None:
        colors = []

    if len(colors) == 0:
        return "Colorless"
    elif len(colors) > 1:
        return "Multicolor"
    else:
        color_map = {"W": "White", "U": "Blue", "B": "Black", "R": "Red", "G": "Green"}
        return color_map.get(colors[0], "Colorless")


def get_card_image_url(scryfall_data: dict) -> Optional[str]:
    """Extract the card image URL from Scryfall data."""
    if not scryfall_data:
        return None

    # Try top-level image_uris first (single-faced cards)
    if "image_uris" in scryfall_data:
        return scryfall_data["image_uris"].get("normal")

    # For double-faced cards, use the front face image
    if "card_faces" in scryfall_data and len(scryfall_data["card_faces"]) > 0:
        face = scryfall_data["card_faces"][0]
        if "image_uris" in face:
            return face["image_uris"].get("normal")

    return None


def merge_continuation_lines(lines: list[str]) -> list[str]:
    """
    Merge lines that are continuations of previous lines.
    TCGplayer PDFs split long entries across multiple lines.
    """
    merged = []
    current = ""

    for line in lines:
        line = line.strip()
        if not line:
            continue

        # Skip page headers/footers
        if line.startswith("Quantity Description") or line.startswith("OrderNumber:"):
            if current:
                merged.append(current)
                current = ""
            continue

        # Skip the total line at the end (e.g., "201 Total $524.25")
        if re.match(r'^\d+\s+Total\s+\$', line):
            if current:
                merged.append(current)
                current = ""
            continue

        # Check if this line starts a new card entry
        if re.match(r'^\d+\s+Magic-', line):
            # Save previous entry if exists
            if current:
                merged.append(current)
            current = line
        elif current:
            # This is a continuation line - append it
            current += line
        # else: skip lines before first card entry

    # Don't forget the last entry
    if current:
        merged.append(current)

    return merged


def parse_packing_slip(pdf_path: str) -> list[Card]:
    """Parse a TCGplayer packing slip PDF and return a list of Cards."""
    print(f"Parsing PDF: {pdf_path}")
    text = extract_text_from_pdf(pdf_path)

    lines = text.split("\n")

    # Merge continuation lines first
    merged_lines = merge_continuation_lines(lines)

    print(f"Found {len(merged_lines)} potential card entries")

    cards = []
    for line in merged_lines:
        card = parse_card_line(line)
        if card:
            cards.append(card)
        else:
            # Debug: show lines that didn't parse
            if "Magic-" in line:
                print(f"  Warning: Could not parse line: {line[:80]}...")

    print(f"Successfully parsed {len(cards)} cards")
    return cards


def get_search_name(card_name: str) -> str:
    """
    Extract the core card name for Scryfall lookup, removing variant info
    that might interfere with the search.
    """
    search_name = card_name.strip()

    # Remove common suffixes that aren't part of the card name
    # These are parsing artifacts from the PDF
    search_name = re.sub(r'\(Extended$', '', search_name).strip()
    search_name = re.sub(r'\(Borderless$', '', search_name).strip()
    search_name = re.sub(r'\(Showcase$', '', search_name).strip()
    search_name = re.sub(r'\(Retro Frame$', '', search_name).strip()
    search_name = re.sub(r'\(Foil Etched$', '', search_name).strip()
    search_name = re.sub(r'\(White Border$', '', search_name).strip()
    search_name = re.sub(r'\(Future Sight$', '', search_name).strip()

    return search_name


def fetch_colors_from_scryfall(cards: list[Card], on_progress=None) -> list[Card]:
    """Fetch color and image information for all cards from Scryfall.

    Args:
        cards: List of Card objects to look up.
        on_progress: Optional callback called with (current_index, total, card_name, status)
                     for each card processed. Used by the API for streaming progress.
    """
    print("\nFetching card data from Scryfall...")

    # Cache to avoid duplicate lookups
    # For variants, we cache by set+collector_number to get exact art
    # For color lookups, we still cache by card name
    image_cache = {}  # (set_name, collector_number) -> image_url
    color_cache = {}  # card_name -> color

    # Track failed lookups for summary
    failed_lookups = []

    total = len(cards)
    for i, card in enumerate(cards):
        # Get the core card name for searching (without variant artifacts)
        search_name = get_search_name(card.card_name)

        # Create cache keys
        # Image cache uses set+collector for exact variant art
        image_cache_key = (card.set_name, card.collector_number)
        # Color cache uses card name (color is same across printings)
        color_cache_key = search_name.lower().strip()

        # Check if we have cached data for this exact printing
        if image_cache_key in image_cache and color_cache_key in color_cache:
            card.color = color_cache[color_cache_key]
            card.image_url = image_cache[image_cache_key]
            status = f"{card.color} (cached)"
            print(f"  [{i+1}/{total}] {card.card_name}: {status}")
        else:
            # Scryfall rate limit: 10 requests per second
            time.sleep(0.1)

            # Try to get exact printing with set code and collector number
            scryfall_data = search_scryfall(search_name, card.set_name, card.collector_number)
            if scryfall_data:
                card.color = get_card_color(scryfall_data)
                card.image_url = get_card_image_url(scryfall_data)
                official_name = scryfall_data.get("name", search_name)
                status = card.color
                print(f"  [{i+1}/{total}] {card.card_name} (searched: {search_name}) -> {official_name}: {card.color}")
                # DO NOT overwrite card.card_name - keep the original for display
            else:
                card.color = "Colorless"
                card.image_url = None
                status = "NOT FOUND"
                failed_lookups.append(f"{card.card_name} (searched: {search_name})")
                print(f"  [{i+1}/{total}] {card.card_name}: NOT FOUND (defaulting to Colorless)")

            # Cache both color and image
            color_cache[color_cache_key] = card.color
            image_cache[image_cache_key] = card.image_url

        if on_progress:
            on_progress(i + 1, total, card.card_name, status)

    if failed_lookups:
        print(f"\n  Warning: {len(failed_lookups)} cards could not be found on Scryfall:")
        for name in failed_lookups[:10]:  # Show first 10
            print(f"    - {name}")
        if len(failed_lookups) > 10:
            print(f"    ... and {len(failed_lookups) - 10} more")

    return cards


def generate_html(cards: list[Card], output_path: str = None, order_number: str = "",
                   order_numbers: dict[str, str] = None):
    """Generate an HTML page organized by color and rarity."""
    import hashlib

    # Create a unique generation ID from the card data so progress resets on new orders
    card_fingerprint = "|".join(
        f"{c.card_name}-{c.set_name}-{c.collector_number}-{c.quantity}"
        for c in sorted(cards, key=lambda c: c.card_name)
    )
    generation_id = hashlib.md5(card_fingerprint.encode()).hexdigest()[:12]

    # Detect multi-order mode
    active_groups = sorted(set(c.order_group for c in cards if c.order_group))
    is_multi_order = len(active_groups) > 1
    if not order_numbers:
        order_numbers = {}

    # Group cards by color, then by rarity
    organized = defaultdict(lambda: defaultdict(list))

    for card in cards:
        organized[card.color][card.rarity].append(card)

    # Sort colors and rarities
    sorted_colors = sorted(organized.keys(), key=lambda c: COLOR_ORDER.get(c, 99))

    # Calculate totals
    total_cards = sum(card.quantity for card in cards)
    total_value = sum(card.total_price for card in cards)

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>MTG Order - Organized by Color & Rarity</title>
    <style>
        * {{
            box-sizing: border-box;
        }}
        body {{
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Oxygen, Ubuntu, sans-serif;
            margin: 0;
            padding: 20px;
            background: #1a1a2e;
            color: #eee;
        }}
        h1 {{
            text-align: center;
            color: #fff;
            margin-bottom: 10px;
        }}
        .order-info {{
            text-align: center;
            color: #aaa;
            margin-bottom: 20px;
        }}
        .summary {{
            display: flex;
            justify-content: center;
            gap: 30px;
            margin-bottom: 30px;
            flex-wrap: wrap;
        }}
        .summary-item {{
            background: #16213e;
            padding: 15px 25px;
            border-radius: 8px;
            text-align: center;
        }}
        .summary-item .number {{
            font-size: 2em;
            font-weight: bold;
            color: #e94560;
        }}
        .summary-item .label {{
            color: #aaa;
            font-size: 0.9em;
        }}
        .color-section {{
            margin-bottom: 30px;
            background: #16213e;
            border-radius: 12px;
            overflow: hidden;
        }}
        .color-section.collapsed .rarity-section {{
            display: none;
        }}
        .color-section.section-complete .color-header {{
            opacity: 0.6;
        }}
        .color-header {{
            padding: 15px 20px;
            font-size: 1.4em;
            font-weight: bold;
            display: flex;
            align-items: center;
            gap: 10px;
            cursor: pointer;
            user-select: none;
        }}
        .color-header:hover {{
            background: rgba(255,255,255,0.05);
        }}
        .collapse-icon {{
            margin-left: auto;
            font-size: 0.8em;
            transition: transform 0.2s;
        }}
        .color-section.collapsed .collapse-icon {{
            transform: rotate(-90deg);
        }}
        .section-progress {{
            font-size: 0.7em;
            font-weight: normal;
            color: #888;
            margin-left: 10px;
        }}
        .section-progress.complete {{
            color: #4caf50;
        }}
        .color-pip {{
            width: 24px;
            height: 24px;
            border-radius: 50%;
            border: 2px solid rgba(255,255,255,0.3);
        }}
        .color-White {{ background: linear-gradient(135deg, #f8f6d8, #e8e4c9); }}
        .color-Blue {{ background: linear-gradient(135deg, #0e68ab, #1a9bc7); }}
        .color-Black {{ background: linear-gradient(135deg, #393939, #1a1a1a); }}
        .color-Red {{ background: linear-gradient(135deg, #d32029, #f44336); }}
        .color-Green {{ background: linear-gradient(135deg, #00733e, #2e7d32); }}
        .color-Multicolor {{ background: linear-gradient(135deg, #c9a227, #ffd700); }}
        .color-Colorless {{ background: linear-gradient(135deg, #9e9e9e, #bdbdbd); }}
        .color-Land {{ background: linear-gradient(135deg, #795548, #a1887f); }}

        .header-White {{ background: linear-gradient(90deg, rgba(248,246,216,0.3), transparent); }}
        .header-Blue {{ background: linear-gradient(90deg, rgba(14,104,171,0.3), transparent); }}
        .header-Black {{ background: linear-gradient(90deg, rgba(57,57,57,0.5), transparent); }}
        .header-Red {{ background: linear-gradient(90deg, rgba(211,32,41,0.3), transparent); }}
        .header-Green {{ background: linear-gradient(90deg, rgba(0,115,62,0.3), transparent); }}
        .header-Multicolor {{ background: linear-gradient(90deg, rgba(201,162,39,0.3), transparent); }}
        .header-Colorless {{ background: linear-gradient(90deg, rgba(158,158,158,0.3), transparent); }}
        .header-Land {{ background: linear-gradient(90deg, rgba(121,85,72,0.3), transparent); }}

        .rarity-section {{
            padding: 10px 20px;
        }}
        .rarity-header {{
            font-size: 1.1em;
            font-weight: 600;
            padding: 8px 0;
            border-bottom: 1px solid rgba(255,255,255,0.1);
            margin-bottom: 10px;
        }}
        .rarity-M {{ color: #ff9800; }}
        .rarity-R {{ color: #ffd700; }}
        .rarity-U {{ color: #90caf9; }}
        .rarity-C {{ color: #aaa; }}
        .rarity-S {{ color: #ce93d8; }}

        .card-list {{
            display: grid;
            gap: 8px;
        }}
        .card-item {{
            display: grid;
            grid-template-columns: 40px 1fr auto;
            gap: 15px;
            padding: 10px 15px;
            background: rgba(255,255,255,0.05);
            border-radius: 6px;
            align-items: center;
            cursor: pointer;
        }}
        .card-item:hover {{
            background: rgba(255,255,255,0.1);
        }}
        .card-qty {{
            font-weight: bold;
            font-size: 1.2em;
            color: #e94560;
            text-align: center;
        }}
        .card-info {{
            display: flex;
            flex-direction: column;
            gap: 2px;
        }}
        .card-name {{
            font-weight: 500;
        }}
        .card-details {{
            font-size: 0.85em;
            color: #888;
        }}
        .card-foil {{
            color: #ffd700;
            font-weight: bold;
        }}
        .card-language {{
            color: #ff6b6b;
            font-weight: bold;
            font-size: 0.9em;
        }}
        .card-variant {{
            font-size: 0.75em;
            font-weight: 600;
            padding: 2px 7px;
            border-radius: 4px;
            margin-left: 6px;
            display: inline-block;
            vertical-align: middle;
            letter-spacing: 0.02em;
        }}
        .variant-borderless {{ background: rgba(0,188,212,0.2); color: #00bcd4; }}
        .variant-extended {{ background: rgba(156,39,176,0.2); color: #ce93d8; }}
        .variant-showcase {{ background: rgba(233,30,99,0.2); color: #f48fb1; }}
        .variant-retro {{ background: rgba(255,143,0,0.2); color: #ffb74d; }}
        .variant-white-border {{ background: rgba(176,190,197,0.2); color: #b0bec5; }}
        .variant-etched {{ background: rgba(184,134,11,0.2); color: #daa520; }}
        .variant-full-art {{ background: rgba(63,81,181,0.2); color: #7986cb; }}
        .variant-future-sight {{ background: rgba(0,229,255,0.2); color: #00e5ff; }}
        .variant-surge {{ background: rgba(255,111,0,0.2); color: #ff9800; }}
        .variant-other {{ background: rgba(120,144,156,0.2); color: #90a4ae; }}
        .card-item.has-variant {{
            border-left: 3px solid var(--variant-color, #78909c);
        }}
        /* Order group pill */
        .order-pill {{
            width: 26px;
            height: 26px;
            border-radius: 50%;
            display: flex;
            align-items: center;
            justify-content: center;
            font-weight: 700;
            font-size: 0.8em;
            color: #fff;
            flex-shrink: 0;
        }}
        .order-pill-A {{ background: #4a9eff; }}
        .order-pill-B {{ background: #ff9800; }}
        .order-pill-C {{ background: #66bb6a; }}
        .card-item.multi-order {{
            grid-template-columns: 30px 40px 1fr auto;
            border-left: 3px solid var(--group-color, transparent);
        }}
        .card-item.multi-order.has-variant {{
            border-left: 3px solid var(--variant-color, #78909c);
            border-right: 3px solid var(--group-color, transparent);
        }}
        /* Order group legend */
        .order-legend {{
            display: flex;
            justify-content: center;
            gap: 20px;
            margin-bottom: 20px;
            flex-wrap: wrap;
        }}
        .order-legend-item {{
            display: flex;
            align-items: center;
            gap: 8px;
            background: #16213e;
            padding: 10px 18px;
            border-radius: 8px;
        }}
        .order-legend-label {{
            color: #ccc;
            font-size: 0.9em;
        }}
        .order-legend-number {{
            color: #888;
            font-size: 0.8em;
        }}
        .card-price {{
            text-align: right;
            color: #4caf50;
            font-weight: 500;
        }}
        .checkbox {{
            width: 20px;
            height: 20px;
            cursor: pointer;
        }}
        .card-item.checked {{
            opacity: 0.5;
        }}
        .card-item.checked .card-name {{
            text-decoration: line-through;
        }}

        .nav {{
            position: sticky;
            top: 0;
            background: #1a1a2e;
            padding: 10px 0;
            margin-bottom: 20px;
            z-index: 100;
            border-bottom: 1px solid #333;
        }}
        .nav-links {{
            display: flex;
            justify-content: center;
            gap: 10px;
            flex-wrap: wrap;
        }}
        .nav-link {{
            padding: 8px 16px;
            border-radius: 20px;
            text-decoration: none;
            color: #fff;
            font-weight: 500;
            transition: transform 0.2s;
        }}
        .nav-link:hover {{
            transform: scale(1.05);
        }}
        .nav-link.complete {{
            opacity: 0.5;
            text-decoration: line-through;
        }}
        .nav-link .nav-progress {{
            font-size: 0.75em;
            opacity: 0.8;
        }}
        /* Dark text for light-colored nav buttons */
        .nav-link.color-White,
        .nav-link.color-Colorless,
        .nav-link.color-Multicolor {{
            color: #1a1a2e;
        }}
        .nav-link.color-White .nav-progress,
        .nav-link.color-Colorless .nav-progress,
        .nav-link.color-Multicolor .nav-progress {{
            opacity: 0.7;
        }}
        .nav-controls {{
            display: flex;
            justify-content: center;
            gap: 10px;
            margin-top: 10px;
            padding-top: 10px;
            border-top: 1px solid #333;
        }}
        .nav-btn {{
            padding: 6px 14px;
            border-radius: 6px;
            border: 1px solid #444;
            background: #2a2a3e;
            color: #ccc;
            font-size: 0.85em;
            cursor: pointer;
            transition: background 0.2s, color 0.2s;
        }}
        .nav-btn:hover {{
            background: #3a3a4e;
            color: #fff;
        }}
        .nav-btn.danger {{
            border-color: #c0392b;
            color: #e74c3c;
        }}
        .nav-btn.danger:hover {{
            background: #c0392b;
            color: #fff;
        }}

        .progress-bar {{
            width: 100%;
            height: 8px;
            background: #333;
            border-radius: 4px;
            margin-bottom: 20px;
            overflow: hidden;
        }}
        .progress-fill {{
            height: 100%;
            background: linear-gradient(90deg, #4caf50, #8bc34a);
            transition: width 0.3s;
            border-radius: 4px;
        }}
        .progress-text {{
            text-align: center;
            color: #aaa;
            margin-bottom: 10px;
            font-size: 0.9em;
        }}

        /* Card image hover preview */
        #card-preview {{
            display: none;
            position: fixed;
            z-index: 1000;
            pointer-events: none;
            border-radius: 12px;
            box-shadow: 0 8px 32px rgba(0,0,0,0.5);
            max-width: 250px;
            max-height: 350px;
        }}
        #card-preview img {{
            display: block;
            width: 100%;
            height: auto;
            border-radius: 12px;
        }}

        @media print {{
            body {{
                background: #fff;
                color: #000;
            }}
            .color-section {{
                background: #f5f5f5;
                break-inside: avoid;
            }}
            .card-item {{
                background: #eee;
            }}
            .nav, .progress-bar, .progress-text {{
                display: none;
            }}
        }}
    </style>
</head>
<body>
    <!-- Card image hover preview -->
    <div id="card-preview"><img src="" alt="Card Preview"></div>

    <h1>MTG Order - Pull Sheet</h1>
    <div class="order-info">{order_number if order_number else 'TCGplayer Order'}</div>

    <div class="progress-text">Progress: <span id="progress-count">0</span> / {len(cards)} items pulled</div>
    <div class="progress-bar">
        <div class="progress-fill" id="progress-fill" style="width: 0%"></div>
    </div>

    <div class="summary">
        <div class="summary-item">
            <div class="number">{total_cards}</div>
            <div class="label">Total Cards</div>
        </div>
        <div class="summary-item">
            <div class="number">${total_value:.2f}</div>
            <div class="label">Total Value</div>
        </div>
        <div class="summary-item">
            <div class="number">{len(cards)}</div>
            <div class="label">Line Items</div>
        </div>
    </div>
{"".join(f'''
    <div class="order-legend">
''' + "".join(f'''        <div class="order-legend-item">
            <div class="order-pill order-pill-{g}">{g}</div>
            <div>
                <div class="order-legend-label">Order {g}</div>
                <div class="order-legend-number">{order_numbers.get(g, '')}</div>
            </div>
            <div class="order-legend-label">({sum(1 for c in cards if c.order_group == g)} items)</div>
        </div>
''' for g in active_groups) + '''    </div>
''') if is_multi_order else ''}

    <nav class="nav">
        <div class="nav-links">
"""

    # Add navigation links with progress counters
    for color in sorted_colors:
        color_cards = organized[color]
        section_total = sum(1 for rarity in color_cards.values() for _ in rarity)
        html += f'            <a href="#{color.lower()}" class="nav-link color-{color}" data-section="{color.lower()}">{color} <span class="nav-progress">(<span class="nav-remaining">{section_total}</span>/{section_total})</span></a>\n'

    html += """        </div>
        <div class="nav-controls">
            <button class="nav-btn" onclick="expandAll()">Expand All</button>
            <button class="nav-btn" onclick="collapseAll()">Collapse All</button>
            <button class="nav-btn danger" onclick="resetProgress()">Reset Progress</button>
        </div>
    </nav>
"""

    # Add card sections
    card_index = 0
    for color in sorted_colors:
        color_cards = organized[color]
        sorted_rarities = sorted(color_cards.keys(), key=lambda r: RARITY_ORDER.get(r, 99))

        color_total = sum(c.quantity for rarity in color_cards.values() for c in rarity)

        section_item_count = sum(1 for rarity in color_cards.values() for _ in rarity)
        html += f"""
    <div class="color-section" id="{color.lower()}" data-section-total="{section_item_count}">
        <div class="color-header header-{color}" onclick="toggleSection(this.parentElement)">
            <span class="color-pip color-{color}"></span>
            {color} ({color_total} cards)
            <span class="section-progress"><span class="section-remaining">{section_item_count}</span> remaining</span>
            <span class="collapse-icon">▼</span>
        </div>
"""

        for rarity in sorted_rarities:
            # Sort cards: variants first, then by card name, then by order group
            # This keeps duplicate cards from different orders adjacent
            rarity_cards = sorted(
                color_cards[rarity],
                key=lambda c: (0 if c.variant else 1, c.card_name, c.order_group or '')
            )
            rarity_name = RARITY_NAMES.get(rarity, rarity)

            html += f"""
        <div class="rarity-section">
            <div class="rarity-header rarity-{rarity}">{rarity_name} ({sum(c.quantity for c in rarity_cards)})</div>
            <div class="card-list">
"""

            for card in rarity_cards:
                foil_badge = '<span class="card-foil"> ★ FOIL</span>' if card.is_foil else ''
                language_badge = f'<span class="card-language"> [{card.language}]</span>' if card.language else ''
                image_attr = f' data-image="{card.image_url}"' if card.image_url else ''

                # Variant badge
                variant_css, variant_label, variant_color = get_variant_style(card.variant)
                if variant_css:
                    variant_badge = f'<span class="card-variant {variant_css}">{variant_label}</span>'
                    variant_style = f' style="--variant-color: {variant_color}"'
                    variant_class = ' has-variant'
                else:
                    variant_badge = ''
                    variant_style = ''
                    variant_class = ''

                # Order group pill (only in multi-order mode)
                if is_multi_order and card.order_group:
                    group = card.order_group
                    group_color = ORDER_GROUP_COLORS.get(group, '#888')
                    order_pill = f'<div class="order-pill order-pill-{group}">{group}</div>'
                    multi_class = ' multi-order'
                    group_style = f' --group-color: {group_color};'
                    data_group = f' data-group="{group}"'
                else:
                    order_pill = ''
                    multi_class = ''
                    group_style = ''
                    data_group = ''

                # Combine inline styles
                combined_style = ''
                if variant_style or group_style:
                    style_parts = []
                    if variant_color:
                        style_parts.append(f'--variant-color: {variant_color}')
                    if group_style:
                        style_parts.append(group_style.strip().rstrip(';'))
                    combined_style = f' style="{"; ".join(style_parts)}"'

                # Escape HTML in card name
                safe_card_name = card.card_name.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
                safe_set_name = card.set_name.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

                html += f"""                <div class="card-item{variant_class}{multi_class}" data-index="{card_index}"{data_group}{image_attr}{combined_style} onclick="toggleCard(this)">
                    {order_pill}<div class="card-qty">{card.quantity}x</div>
                    <div class="card-info">
                        <div class="card-name">{safe_card_name}{variant_badge}{foil_badge}{language_badge}</div>
                        <div class="card-details">{safe_set_name} #{card.collector_number} - {card.condition}</div>
                    </div>
                    <div class="card-price">${card.total_price:.2f}</div>
                </div>
"""
                card_index += 1

            html += """            </div>
        </div>
"""

        html += """    </div>
"""

    html += f"""
    <script>
        const totalItems = {len(cards)};

        function updateProgress() {{
            const checked = document.querySelectorAll('.card-item.checked').length;
            document.getElementById('progress-count').textContent = checked;
            document.getElementById('progress-fill').style.width = (checked / totalItems * 100) + '%';
        }}

        function updateSectionProgress() {{
            document.querySelectorAll('.color-section').forEach((section) => {{
                const total = parseInt(section.dataset.sectionTotal);
                const checked = section.querySelectorAll('.card-item.checked').length;
                const remaining = total - checked;
                const sectionId = section.id;

                // Update section header
                const progressSpan = section.querySelector('.section-remaining');
                const progressContainer = section.querySelector('.section-progress');
                if (progressSpan) {{
                    progressSpan.textContent = remaining;
                }}
                if (progressContainer) {{
                    if (remaining === 0) {{
                        progressContainer.classList.add('complete');
                        progressContainer.innerHTML = '✓ Complete';
                        section.classList.add('section-complete');
                    }} else {{
                        progressContainer.classList.remove('complete');
                        progressContainer.innerHTML = '<span class="section-remaining">' + remaining + '</span> remaining';
                        section.classList.remove('section-complete');
                    }}
                }}

                // Update nav link
                const navLink = document.querySelector('.nav-link[data-section="' + sectionId + '"]');
                if (navLink) {{
                    const navRemaining = navLink.querySelector('.nav-remaining');
                    if (navRemaining) {{
                        navRemaining.textContent = remaining;
                    }}
                    if (remaining === 0) {{
                        navLink.classList.add('complete');
                    }} else {{
                        navLink.classList.remove('complete');
                    }}
                }}
            }});
        }}

        function toggleSection(section) {{
            section.classList.toggle('collapsed');
            const sectionId = section.id;
            if (section.classList.contains('collapsed')) {{
                localStorage.setItem('section-' + sectionId + '-collapsed', 'true');
            }} else {{
                localStorage.removeItem('section-' + sectionId + '-collapsed');
            }}
        }}

        function expandAll() {{
            document.querySelectorAll('.color-section').forEach((section) => {{
                section.classList.remove('collapsed');
                localStorage.removeItem('section-' + section.id + '-collapsed');
            }});
        }}

        function collapseAll() {{
            document.querySelectorAll('.color-section').forEach((section) => {{
                section.classList.add('collapsed');
                localStorage.setItem('section-' + section.id + '-collapsed', 'true');
            }});
        }}

        function resetProgress() {{
            if (!confirm('Reset all progress? This will uncheck all cards and cannot be undone.')) {{
                return;
            }}
            document.querySelectorAll('.card-item').forEach((item) => {{
                item.classList.remove('checked');
                localStorage.removeItem('card-' + item.dataset.index);
            }});
            document.querySelectorAll('.color-section').forEach((section) => {{
                section.classList.remove('collapsed');
                localStorage.removeItem('section-' + section.id + '-collapsed');
            }});
            updateProgress();
            updateSectionProgress();
        }}

        function toggleCard(element) {{
            element.classList.toggle('checked');
            const index = element.dataset.index;
            if (element.classList.contains('checked')) {{
                localStorage.setItem('card-' + index, 'checked');
            }} else {{
                localStorage.removeItem('card-' + index);
            }}
            updateProgress();
            updateSectionProgress();
        }}

        // Auto-reset progress when a new order is loaded
        const generationId = '{generation_id}';
        if (localStorage.getItem('generation-id') !== generationId) {{
            // New order detected — clear all old progress
            document.querySelectorAll('.card-item').forEach((item) => {{
                localStorage.removeItem('card-' + item.dataset.index);
            }});
            document.querySelectorAll('.color-section').forEach((section) => {{
                localStorage.removeItem('section-' + section.id + '-collapsed');
            }});
            localStorage.setItem('generation-id', generationId);
        }} else {{
            // Same order — restore checked state
            document.querySelectorAll('.card-item').forEach((item) => {{
                const index = item.dataset.index;
                if (localStorage.getItem('card-' + index) === 'checked') {{
                    item.classList.add('checked');
                }}
            }});

            // Restore collapsed state
            document.querySelectorAll('.color-section').forEach((section) => {{
                const sectionId = section.id;
                if (localStorage.getItem('section-' + sectionId + '-collapsed') === 'true') {{
                    section.classList.add('collapsed');
                }}
            }});
        }}

        updateProgress();
        updateSectionProgress();

        // Auto-expand section when clicking nav link
        document.querySelectorAll('.nav-link').forEach((link) => {{
            link.addEventListener('click', (e) => {{
                const sectionId = link.dataset.section;
                const section = document.getElementById(sectionId);
                if (section && section.classList.contains('collapsed')) {{
                    section.classList.remove('collapsed');
                    localStorage.removeItem('section-' + sectionId + '-collapsed');
                }}
            }});
        }});

        // Card image hover preview
        const preview = document.getElementById('card-preview');
        const previewImg = preview.querySelector('img');

        document.querySelectorAll('.card-item[data-image]').forEach((item) => {{
            item.addEventListener('mouseenter', (e) => {{
                const imageUrl = item.dataset.image;
                if (imageUrl) {{
                    previewImg.src = imageUrl;
                    preview.style.display = 'block';
                }}
            }});

            item.addEventListener('mousemove', (e) => {{
                // Position preview to the right of cursor, or left if near edge
                const padding = 20;
                let x = e.clientX + padding;
                let y = e.clientY - 100;

                // Keep preview on screen
                if (x + 260 > window.innerWidth) {{
                    x = e.clientX - 260 - padding;
                }}
                if (y < 10) {{
                    y = 10;
                }}
                if (y + 360 > window.innerHeight) {{
                    y = window.innerHeight - 360;
                }}

                preview.style.left = x + 'px';
                preview.style.top = y + 'px';
            }});

            item.addEventListener('mouseleave', () => {{
                preview.style.display = 'none';
            }});
        }});
    </script>
</body>
</html>
"""

    if output_path:
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(html)
        print(f"\nGenerated HTML: {output_path}")

    return html


def main():
    import sys

    if len(sys.argv) < 2:
        print("Usage: python mtg_packing_slip_organizer.py <slip1.pdf> [slip2.pdf] [slip3.pdf] [-o output.html]")
        print("\nThis tool parses TCGplayer packing slips and generates an HTML page")
        print("organized by card color and rarity for easier order fulfillment.")
        print("Pass up to 3 PDFs to merge into a single multi-order pull sheet.")
        sys.exit(1)

    # Parse arguments: PDF files and optional -o output path
    pdf_paths = []
    output_path = None
    args = sys.argv[1:]
    i = 0
    while i < len(args):
        if args[i] == '-o' and i + 1 < len(args):
            output_path = args[i + 1]
            i += 2
        else:
            pdf_paths.append(args[i])
            i += 1

    if len(pdf_paths) > 3:
        print("Error: Maximum 3 PDFs supported for multi-order pull sheets.")
        sys.exit(1)

    # Validate all files exist
    for p in pdf_paths:
        if not Path(p).exists():
            print(f"Error: File not found: {p}")
            sys.exit(1)

    if not output_path:
        output_path = pdf_paths[0].rsplit('.', 1)[0] + '_organized.html'

    group_labels = ['A', 'B', 'C']
    all_cards = []
    order_numbers = {}
    is_multi = len(pdf_paths) > 1

    for idx, pdf_path in enumerate(pdf_paths):
        group = group_labels[idx] if is_multi else None
        if is_multi:
            print(f"\n--- Order {group} ---")

        cards = parse_packing_slip(pdf_path)
        if not cards:
            print(f"Warning: No cards found in {pdf_path}")
            continue

        # Assign order group
        if group:
            for card in cards:
                card.order_group = group

        # Extract order number
        text = extract_text_from_pdf(pdf_path)
        order_match = re.search(r"Order\s*Number:\s*([A-Z0-9-]+)", text)
        order_num = order_match.group(1) if order_match else ""
        if group:
            order_numbers[group] = order_num

        all_cards.extend(cards)

    if not all_cards:
        print("No cards found in any PDF. Please check the file formats.")
        sys.exit(1)

    # Fetch colors from Scryfall
    all_cards = fetch_colors_from_scryfall(all_cards)

    # Generate HTML
    order_label = order_numbers.get('A', '') if not is_multi else "Multi-Order Pull Sheet"
    generate_html(all_cards, output_path, order_label, order_numbers=order_numbers if is_multi else None)

    print(f"\nDone! Open {output_path} in your browser.")


if __name__ == "__main__":
    main()
