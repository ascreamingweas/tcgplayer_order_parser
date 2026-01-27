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


# Rarity mapping
RARITY_ORDER = {"M": 0, "R": 1, "U": 2, "C": 3, "S": 4}
RARITY_NAMES = {"M": "Mythic Rare", "R": "Rare", "U": "Uncommon", "C": "Common", "S": "Special"}

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

# Known set name patterns (to help with parsing)
# These are patterns that help identify where set names end
KNOWN_SET_PREFIXES = [
    "LorwynEclipsed",
    "Avatar:TheLastAirbender:Eternal-Legal",
    "Avatar:TheLastAirbender",
    "MarvelUniverseEternal-Legal",
    "Marvel'sSpider-Man",
    "EdgeofEternities",
    "Commander:EdgeofEternities",
    "FINALFANTASY",
    "Commander:FINALFANTASY",
    "Tarkir:Dragonstorm",
    "Commander:Tarkir:Dragonstorm",
    "Aetherdrift",
    "Phyrexia:AllWillBeOne",
    "WaroftheSpark",
    "Foundations",
    "CommanderLegends:BattleforBaldur'sGate",
    "Commander:OutlawsofThunderJunction",
    "Commander:StreetsofNewCapenna",
    "Commander2016",
    "ModernHorizons3",
    "RavnicaRemastered",
    "TimeSpiral:Remastered",
    "TheListReprints",
    "MysteryBooster2",
    "SecretLairDropSeries",
    "SecretLairCountdownKit",
    "Urza'sLegacy",
    "FromtheVault:Lore",
    "PromoPack:OutlawsofThunderJunction",
    "PromoPack:MarchoftheMachine",
    "PromoPack:Kamigawa:NeonDynasty",
    "CommanderMasters",
    "Innistrad:MidnightHunt",
    "OutlawsofThunderJunction",
    "MurdersatKarlovManor",
]

# Global cache for Scryfall set mapping (populated on first use)
_scryfall_set_cache: Optional[dict] = None


def fetch_scryfall_sets() -> dict:
    """
    Fetch all sets from Scryfall API and build a mapping from various name formats
    to set codes. This ensures we always have the latest sets.
    """
    global _scryfall_set_cache

    if _scryfall_set_cache is not None:
        return _scryfall_set_cache

    print("Syncing set list from Scryfall...")

    try:
        url = "https://api.scryfall.com/sets"
        req = urllib.request.Request(url, headers={"User-Agent": "MTG-Packing-Slip-Organizer/1.0"})
        with urllib.request.urlopen(req, timeout=30) as response:
            data = json.loads(response.read().decode())

        mapping = {}
        sets_data = data.get("data", [])

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

            # Add version with colons but no spaces around them
            colon_no_space = name.replace(": ", ":")
            mapping[colon_no_space] = code

            # Add version with no spaces at all (including around colons)
            all_no_spaces = name.replace(" ", "").replace(":", "")
            if all_no_spaces not in mapping:
                mapping[all_no_spaces] = code

        # Add some manual overrides for TCGPlayer-specific naming quirks
        # TCGPlayer sometimes uses different names than Scryfall
        manual_overrides = {
            # TCGPlayer uses "Drop Series" suffix
            "SecretLairDropSeries": "sld",
            "Secret Lair Drop Series": "sld",
            # TCGPlayer uses "Countdown Kit"
            "SecretLairCountdownKit": "slc",
            "Secret Lair Countdown Kit": "slc",
            # TCGPlayer uses "Eternal-Legal" suffix for some sets
            "Avatar:TheLastAirbender:Eternal-Legal": "tle",
            "Avatar: The Last Airbender: Eternal-Legal": "tle",
            "MarvelUniverseEternal-Legal": "mar",
            "Marvel Universe Eternal-Legal": "mar",
            # The List
            "TheListReprints": "plst",
            "The List Reprints": "plst",
            # Time Spiral Remastered (TCGPlayer uses colon)
            "TimeSpiral:Remastered": "tsr",
            "Time Spiral: Remastered": "tsr",
        }
        mapping.update(manual_overrides)

        _scryfall_set_cache = mapping
        print(f"  Loaded {len(sets_data)} sets from Scryfall")
        return mapping

    except Exception as e:
        print(f"  Warning: Could not fetch sets from Scryfall: {e}")
        print("  Falling back to basic name matching")
        _scryfall_set_cache = {}
        return {}


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
    # Sort known prefixes by length (longest first) to match most specific first
    sorted_prefixes = sorted(KNOWN_SET_PREFIXES, key=len, reverse=True)

    for prefix in sorted_prefixes:
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
        req = urllib.request.Request(url, headers={"User-Agent": "MTG-Packing-Slip-Organizer/1.0"})
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
        req = urllib.request.Request(url, headers={"User-Agent": "MTG-Packing-Slip-Organizer/1.0"})
        with urllib.request.urlopen(req, timeout=10) as response:
            data = json.loads(response.read().decode())
            return data
    except urllib.error.HTTPError as e:
        if e.code == 404:
            # Try exact search
            params = {"exact": search_name}
            url = f"{base_url}?{urllib.parse.urlencode(params)}"
            try:
                req = urllib.request.Request(url, headers={"User-Agent": "MTG-Packing-Slip-Organizer/1.0"})
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
                    req = urllib.request.Request(url, headers={"User-Agent": "MTG-Packing-Slip-Organizer/1.0"})
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


def fetch_colors_from_scryfall(cards: list[Card]) -> list[Card]:
    """Fetch color and image information for all cards from Scryfall."""
    print("\nFetching card data from Scryfall...")

    # Cache to avoid duplicate lookups
    # For variants, we cache by set+collector_number to get exact art
    # For color lookups, we still cache by card name
    image_cache = {}  # (set_name, collector_number) -> image_url
    color_cache = {}  # card_name -> color

    # Track failed lookups for summary
    failed_lookups = []

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
            print(f"  [{i+1}/{len(cards)}] {card.card_name}: {card.color} (cached)")
        else:
            # Scryfall rate limit: 10 requests per second
            time.sleep(0.1)

            # Try to get exact printing with set code and collector number
            scryfall_data = search_scryfall(search_name, card.set_name, card.collector_number)
            if scryfall_data:
                card.color = get_card_color(scryfall_data)
                card.image_url = get_card_image_url(scryfall_data)
                official_name = scryfall_data.get("name", search_name)
                print(f"  [{i+1}/{len(cards)}] {card.card_name} (searched: {search_name}) -> {official_name}: {card.color}")
                # DO NOT overwrite card.card_name - keep the original for display
            else:
                card.color = "Colorless"
                card.image_url = None
                failed_lookups.append(f"{card.card_name} (searched: {search_name})")
                print(f"  [{i+1}/{len(cards)}] {card.card_name}: NOT FOUND (defaulting to Colorless)")

            # Cache both color and image
            color_cache[color_cache_key] = card.color
            image_cache[image_cache_key] = card.image_url

    if failed_lookups:
        print(f"\n  Warning: {len(failed_lookups)} cards could not be found on Scryfall:")
        for name in failed_lookups[:10]:  # Show first 10
            print(f"    - {name}")
        if len(failed_lookups) > 10:
            print(f"    ... and {len(failed_lookups) - 10} more")

    return cards


def generate_html(cards: list[Card], output_path: str, order_number: str = ""):
    """Generate an HTML page organized by color and rarity."""

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
        .color-header {{
            padding: 15px 20px;
            font-size: 1.4em;
            font-weight: bold;
            display: flex;
            align-items: center;
            gap: 10px;
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

    <nav class="nav">
        <div class="nav-links">
"""

    # Add navigation links
    for color in sorted_colors:
        html += f'            <a href="#{color.lower()}" class="nav-link color-{color}">{color}</a>\n'

    html += """        </div>
    </nav>
"""

    # Add card sections
    card_index = 0
    for color in sorted_colors:
        color_cards = organized[color]
        sorted_rarities = sorted(color_cards.keys(), key=lambda r: RARITY_ORDER.get(r, 99))

        color_total = sum(c.quantity for rarity in color_cards.values() for c in rarity)

        html += f"""
    <div class="color-section" id="{color.lower()}">
        <div class="color-header header-{color}">
            <span class="color-pip color-{color}"></span>
            {color} ({color_total} cards)
        </div>
"""

        for rarity in sorted_rarities:
            # Sort cards: variants (non-traditional borders) first, then by card name
            # This groups special printings together at the top for easier pulling
            rarity_cards = sorted(
                color_cards[rarity],
                key=lambda c: (0 if c.variant else 1, c.card_name)
            )
            rarity_name = RARITY_NAMES.get(rarity, rarity)

            html += f"""
        <div class="rarity-section">
            <div class="rarity-header rarity-{rarity}">{rarity_name} ({sum(c.quantity for c in rarity_cards)})</div>
            <div class="card-list">
"""

            for card in rarity_cards:
                foil_badge = '<span class="card-foil"> ★ FOIL</span>' if card.is_foil else ''
                variant_text = f" ({card.variant})" if card.variant else ""
                language_badge = f'<span class="card-language"> [{card.language}]</span>' if card.language else ''
                image_attr = f' data-image="{card.image_url}"' if card.image_url else ''

                # Escape HTML in card name
                safe_card_name = card.card_name.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
                safe_set_name = card.set_name.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

                html += f"""                <div class="card-item" data-index="{card_index}"{image_attr} onclick="toggleCard(this)">
                    <div class="card-qty">{card.quantity}x</div>
                    <div class="card-info">
                        <div class="card-name">{safe_card_name}{variant_text}{foil_badge}{language_badge}</div>
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

        function toggleCard(element) {{
            element.classList.toggle('checked');
            const index = element.dataset.index;
            if (element.classList.contains('checked')) {{
                localStorage.setItem('card-' + index, 'checked');
            }} else {{
                localStorage.removeItem('card-' + index);
            }}
            updateProgress();
        }}

        // Restore checked state on load
        document.querySelectorAll('.card-item').forEach((item) => {{
            const index = item.dataset.index;
            if (localStorage.getItem('card-' + index) === 'checked') {{
                item.classList.add('checked');
            }}
        }});
        updateProgress();

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

    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html)

    print(f"\nGenerated HTML: {output_path}")


def main():
    import sys

    if len(sys.argv) < 2:
        print("Usage: python mtg_packing_slip_organizer.py <packing_slip.pdf> [output.html]")
        print("\nThis tool parses TCGplayer packing slips and generates an HTML page")
        print("organized by card color and rarity for easier order fulfillment.")
        sys.exit(1)

    pdf_path = sys.argv[1]

    if not Path(pdf_path).exists():
        print(f"Error: File not found: {pdf_path}")
        sys.exit(1)

    # Default output path
    output_path = sys.argv[2] if len(sys.argv) > 2 else pdf_path.rsplit('.', 1)[0] + '_organized.html'

    # Parse the PDF
    cards = parse_packing_slip(pdf_path)

    if not cards:
        print("No cards found in the PDF. Please check the file format.")
        sys.exit(1)

    # Fetch colors from Scryfall
    cards = fetch_colors_from_scryfall(cards)

    # Extract order number from PDF text
    text = extract_text_from_pdf(pdf_path)
    order_match = re.search(r"Order\s*Number:\s*([A-Z0-9-]+)", text)
    order_number = order_match.group(1) if order_match else ""

    # Generate HTML
    generate_html(cards, output_path, order_number)

    print(f"\nDone! Open {output_path} in your browser.")


if __name__ == "__main__":
    main()
