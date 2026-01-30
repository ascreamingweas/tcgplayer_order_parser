# TCGplayer Order Parser - Project Context

## Overview

This tool parses TCGplayer packing slip PDFs and generates an interactive HTML pull sheet organized by Magic: The Gathering card color and rarity. It's designed to streamline the order fulfillment process for TCG sellers.

## Project Structure

```
tcgplayer_order_parser/
├── mtg_packing_slip_organizer.py   # Main application (~1,275 lines)
├── README.md                        # User documentation
├── Makefile                         # Build/run automation
├── Context.md                       # This file
└── venv/                            # Python virtual environment
```

## How It Works

### Workflow

1. **PDF Parsing** - Extracts text from TCGplayer packing slip PDF using `pdfplumber`
2. **Card Extraction** - Parses card entries (quantity, name, set, collector number, condition, price, etc.)
3. **Scryfall Sync** - Fetches latest Magic set list from Scryfall API
4. **Card Lookup** - Queries Scryfall for each card's color, rarity, and image URL
5. **Organization** - Groups cards by Color → Rarity → Variant type
6. **HTML Generation** - Creates interactive pull sheet with progress tracking

### Card Data Structure

```python
@dataclass
class Card:
    quantity: int              # Number of copies
    set_name: str             # Magic set name (e.g., "Foundations")
    card_name: str            # Card name (e.g., "Lightning Bolt")
    collector_number: str     # Set-specific card ID
    rarity: str               # M/R/U/C/S (Mythic/Rare/Uncommon/Common/Special)
    condition: str            # Near Mint, Lightly Played, etc.
    is_foil: bool             # Foil or regular
    price: float              # Individual card price
    total_price: float        # Quantity × price
    variant: Optional[str]    # Extended Art, Showcase, etc.
    language: Optional[str]   # Non-English language if applicable
    color: str                # Determined from Scryfall API
    image_url: Optional[str]  # Scryfall image URL
```

### Key Features

| Feature | Description |
|---------|-------------|
| Color Organization | WUBRG order + Multicolor + Colorless + Land |
| Rarity Sorting | Mythic → Rare → Uncommon → Common → Special |
| Variant Prioritization | Special printings listed first within each rarity |
| Double-Faced Cards | Classified by front face for proper searching |
| Card Images | Hover preview showing actual card art |
| Language Detection | Flags non-English cards with visible badge |
| Progress Tracking | Click to mark pulled; persists in localStorage |

## Dependencies

- **Python 3.8+**
- **pdfplumber** - PDF text extraction
- **Scryfall API** - Card data lookup (no auth required, rate limited ~10 req/sec)

## Usage

```bash
# Setup (one time)
make setup

# Process a packing slip
make run PDF=TCGplayer_PackingSlips_20260127_114359.pdf

# With custom output filename
make run PDF=packing_slip.pdf OUTPUT=my_orders.html

# List available PDFs
make list
```

## Key Code Sections

### Entry Point
- `main()` function at bottom of file
- Accepts PDF path as first argument, optional output filename as second

### PDF Parsing
- `parse_packing_slip()` - Main parsing orchestrator
- `parse_card_line()` - Extracts individual card data from text lines
- Handles continuation lines split across multiple PDF lines

### Scryfall Integration
- `fetch_scryfall_sets()` - Gets current set list (cached globally)
- `get_scryfall_set_code()` - Maps TCGPlayer set names to Scryfall codes
- `lookup_card_on_scryfall()` - Fetches card color and image URL
- Includes manual overrides for TCGPlayer-specific naming quirks

### Card Processing
- `add_spaces_to_card_name()` - Fixes CamelCase card names from PDF parsing
- `extract_set_name()` - Parses set name from card description
- `detect_rarity()` - Extracts rarity from various PDF formats
- `detect_language()` - Identifies non-English cards

### HTML Generation
- `generate_html()` - Creates the complete HTML output
- Dark theme with color-coded sections
- Responsive grid layout
- JavaScript for interactive features (mark as pulled, progress tracking)

## Common Issues & Solutions

### Set Name Mismatches
TCGPlayer and Scryfall sometimes use different set names. Manual overrides are defined around line 156:
```python
set_overrides = {
    "Murders at Karlov Manor: Clue Edition": "mkc",
    # ... more overrides
}
```

### Card Name Parsing
PDF extraction can produce CamelCase names. The `add_spaces_to_card_name()` function handles common patterns.

### Rarity Detection
Multiple regex patterns handle various PDF formatting variations where rarity codes appear.

## API Notes

### Scryfall API
- Base URL: `https://api.scryfall.com/`
- No authentication required
- Rate limit: ~10 requests/second (0.1s delay between requests)
- Endpoints used:
  - `/sets` - Get all Magic sets
  - `/cards/{set_code}/{collector_number}` - Get card by set and number
  - `/cards/named?fuzzy={name}` - Fuzzy name search fallback

## Future Enhancement Ideas

- Cache Scryfall responses to disk for offline use
- Support for other TCGs (Pokemon, Yu-Gi-Oh)
- Batch processing of multiple PDFs
- Export to other formats (CSV, JSON)
