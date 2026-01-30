# TCGplayer Packing Slip Organizer

Parses TCGplayer packing slip PDFs and generates an HTML pull sheet organized by card color and rarity for easier Magic: The Gathering order fulfillment.

## Features

- **Color-organized output** - Cards grouped by color (WUBRG order), then Multicolor, Colorless, and Lands
- **Rarity sorting** - Within each color, cards sorted by rarity (Mythic → Rare → Uncommon → Common)
- **Variants first** - Within each rarity, special printings (Borderless, Showcase, Extended Art, Retro Frame, etc.) appear before traditional border cards
- **Preserves fulfillment details** - Card name, variant, foil status, set, collector number, condition, and price
- **Language detection** - Non-English cards (Japanese, German, French, Italian, etc.) are flagged with a visible badge
- **Card image preview** - Hover over any card to see the actual card image from Scryfall
- **Exact variant art** - Shows the correct art for each specific printing (Extended Art, Showcase, etc.)
- **Interactive HTML** - Click cards to mark them as pulled; progress persists in your browser
- **Scryfall integration** - Automatically syncs set list and looks up card data via the Scryfall API
- **Double-faced card support** - Transform and modal cards are classified by their front face

## Installation

```bash
# Clone the repository
git clone git@github.com:ascreamingweas/tcgplayer_order_parser.git
cd tcgplayer_order_parser

# Set up virtual environment and install dependencies
make setup
```

## Usage

```bash
# Run the organizer on a packing slip PDF (supports full paths)
make run PDF=~/Downloads/TCGplayer_PackingSlips.pdf

# Optionally specify output filename
make run PDF=~/Downloads/packing_slip.pdf OUTPUT=my_orders.html

# List generated HTML files
make list

# Clean up generated files
make clean
```

The script will:
1. Parse the PDF and extract all card entries
2. Sync the latest set list from Scryfall (ensures new sets are always supported)
3. Look up each card's color and image on Scryfall (cached to avoid duplicate lookups)
4. Generate an HTML file organized by color, rarity, and variant type

## Output

Generated HTML files are saved to the `output/` directory.

The generated HTML includes:
- **Summary stats** - Total cards, total value, line items
- **Progress bar** - Track how many items you've pulled
- **Color navigation** - Jump to specific color sections
- **Card image hover** - See the actual card art when hovering over any card
- **Click-to-mark** - Click any card to mark it as pulled (strikethrough + faded)
- **Print-friendly** - Clean output when printing

## Requirements

- Python 3.8+
- pdfplumber
- Internet connection (for Scryfall API)

## License

MIT
