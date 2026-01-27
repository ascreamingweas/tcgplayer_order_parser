# TCGplayer Packing Slip Organizer

Parses TCGplayer packing slip PDFs and generates an HTML pull sheet organized by card color and rarity for easier Magic: The Gathering order fulfillment.

## Features

- **Color-organized output** - Cards grouped by color (WUBRG order), then Multicolor, Colorless, and Lands
- **Rarity sorting** - Within each color, cards sorted by rarity (Mythic → Rare → Uncommon → Common)
- **Preserves fulfillment details** - Card name, variant (Extended Art, Borderless, Showcase, etc.), foil status, set, collector number, condition, and price
- **Language detection** - Non-English cards (Japanese, German, French, Italian, etc.) are flagged with a visible badge
- **Interactive HTML** - Click cards to mark them as pulled; progress persists in your browser
- **Scryfall integration** - Automatically looks up card colors via the Scryfall API

## Installation

```bash
# Clone the repository
git clone git@github.com:ascreamingweas/tcgplayer_order_parser.git
cd tcgplayer_order_parser

# Create a virtual environment and install dependencies
python3 -m venv venv
source venv/bin/activate
pip install pdfplumber
```

## Usage

```bash
# Activate the virtual environment
source venv/bin/activate

# Run the organizer on a packing slip PDF
python3 mtg_packing_slip_organizer.py "path/to/TCGplayer_PackingSlip.pdf"

# Optionally specify output filename
python3 mtg_packing_slip_organizer.py "packing_slip.pdf" "output.html"
```

The script will:
1. Parse the PDF and extract all card entries
2. Look up each card's color on Scryfall (cached to avoid duplicate lookups)
3. Generate an HTML file organized by color and rarity

## Output

The generated HTML includes:
- **Summary stats** - Total cards, total value, line items
- **Progress bar** - Track how many items you've pulled
- **Color navigation** - Jump to specific color sections
- **Click-to-mark** - Click any card to mark it as pulled (strikethrough + faded)
- **Print-friendly** - Clean output when printing

## Requirements

- Python 3.8+
- pdfplumber

## License

MIT
