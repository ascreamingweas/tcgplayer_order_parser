# TCGplayer Packing Slip Organizer

Parses TCGplayer packing slip PDFs (or a TCGplayer pull sheet CSV covering all open orders) and generates an HTML pull sheet organized by card color and rarity for easier Magic: The Gathering order fulfillment.

## Features

- **Pull sheet CSV support** - Drop TCGplayer's pull sheet export (every open order in one file) instead of individual packing slips. Non-Magic cards (Lorcana, Gundam, etc.) are skipped from the sheet but still counted per order
- **Orders panel** - For pull sheets, a table of every order with its short ID, which TCGs it contains, and card counts; click an order to show only its cards (quantities adjust for cards shared between orders)
- **Order ID chips** - Each card row in a pull sheet shows which orders it goes to and how many each, e.g. `0710DC ×1`
- **Bulk fulfillment** - Upload the pull sheet together with TCGplayer's shipping export and the report becomes a set of steps (see below)

- **Color-organized output** - Cards grouped by color (WUBRG order), then Multicolor, Colorless, and Lands
- **Rarity sorting** - Within each color, cards sorted by rarity (Mythic → Rare → Uncommon → Common)
- **Variants first** - Within each rarity, special printings (Borderless, Showcase, Extended Art, Retro Frame, etc.) appear before traditional border cards
- **Premium Treatments** - Borderless, Showcase, Extended Art, Retro Frame, White Border, Full Art, Alternate Art, every bonus-sheet card (Scryfall's "masterpiece" sets: Mystical Archive, Retro Frame Artifacts, Breaking News, Enchanting Tales, Multiverse Legends, Special Guests, and any new ones), every Innistrad: Double Feature card, and Special-rarity cards (e.g. Time Spiral Remastered's timeshifted cards) are grouped under Premium Treatments in each color
- **Preserves fulfillment details** - Card name, variant, foil status, set, collector number, condition, and price
- **Language detection** - Non-English cards (Japanese, German, French, Italian, etc.) are flagged with a visible badge
- **Card image preview** - Hover over any card to see the actual card image from Scryfall
- **Exact variant art** - Shows the correct art for each specific printing (Extended Art, Showcase, etc.)
- **Interactive HTML** - Click cards to mark them as pulled; progress persists in your browser
- **Scryfall integration** - Automatically syncs set list and looks up card data from a local copy of Scryfall's card data, falling back to the live Scryfall API (see below)
- **Double-faced card support** - Transform and modal cards are classified by their front face

## Bulk fulfillment (pull sheet + shipping export)

Choose **Pull Sheet (CSV)** and drop both CSVs from TCGplayer: the pull sheet and the shipping export for the same orders (either order). The report then has steps:

1. **Upload** - back to the start for a new batch.
2. **Pull** - the usual color/rarity pull sheet with every Magic single from every order. Each order chip shows its label number (`#4 F1175B ×1`); orders outside the bulk process have a dashed orange chip, and hovering it says why.
3. **Labels** - address labels for the bulk orders, numbered #1, #2, ... Pick a layout (4, 2, 6 or 1 per Rollo 4x6 sheet, address-label rolls, or a custom size), the matching line (`#4 · F1175B · 2x` by default; the count never says "cards"), capitals, cut lines, and an optional return address. The preview is the PDF you print. Mark labels printed so the next pass can hide them.
4. **Pack** - one box per label number with the cards that go in that envelope. Click a card once it's in; the envelope shows packed when every card is in.
5. **Set aside** - orders to fulfill another way, with the reason and their Magic cards.

**Missing cards:** right-click a card on the Pull step (or a card in an envelope on Pack / Set aside) to say how many copies you couldn't find, per order. The card gets a red "missing" badge, its envelope shows only what's left to pack, and the Pack step lists **Refunds to send** with each buyer, order number, and missing cards ("Copy details" copies them for your message to the buyer). An order with every card missing is set aside. Labels count only the cards that ship, so mark missing cards before printing labels when you can.

An order is **bulk** when it paid at most $5.00 shipping, is worth under $50, holds only Magic singles, ships within the US, and is in both files. Orders with sealed Magic (pull sheet rows with no card number or rarity) or another game are set aside, but their Magic singles are still on the pull sheet. Both dollar limits are settings on the Labels step. Orders with more than 15 cards (also a setting) stay bulk but get flagged.

Labels are built by the web app (`/api/labels.pdf`), so run `make serve` (or use the deployed site). A report made with `make pullsheet CSV=... SHIP=...` has every step, but printing labels from it needs the web app. Settings carry over between batches; printed and packed marks are kept per batch in your browser.

## Local Scryfall data

Card lookups use a local copy of Scryfall's daily "Default Cards" download (every English printing, ~80 MB download, ~70 MB on disk in `data/scryfall.sqlite`). A pull sheet of a few hundred cards looks up in about a second instead of a couple of minutes, without Scryfall rate limits.

- The web app downloads it on startup if it's missing and checks for a new copy daily. For the command line, run `make scryfall-refresh` once (and whenever you like after).
- Anything the local copy can't answer (a card printed or previewed since the last download, or a set Scryfall doesn't have) is looked up live as before, and recorded.
- `make scryfall-status` (or `GET /api/scryfall` on the web app) shows when the copy was downloaded and the recent cards it had to look up live.
- Set `SCRYFALL_DB=/path/to/file.sqlite` to keep it somewhere else, e.g. a persistent disk on the deployed server.

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

# Or process a TCGplayer pull sheet CSV (all open orders at once)
make pullsheet CSV=~/Downloads/TCGplayer_PullSheet.csv

# Run the web app (needed to print labels) at http://localhost:8000
make serve

# Optionally specify output filename
make run PDF=~/Downloads/packing_slip.pdf OUTPUT=my_orders.html

# List generated HTML files
make list

# Clean up generated files
make clean
```

The pull sheet export has no prices or buyer names, so pull sheets show order counts instead of total value.

The script will:
1. Parse the PDF (or pull sheet CSV) and extract all card entries
2. Sync the latest set list from Scryfall (ensures new sets are always supported)
3. Look up each card's color and image on Scryfall (cached to avoid duplicate lookups)
4. Generate an HTML file organized by color, rarity, and variant type

## Output

Generated HTML files are saved to the `output/` directory.

The generated HTML includes:
- **Summary stats** - Total cards, total value, line items
- **Progress bar** - Track how many items you've pulled
- **Color navigation** - Jump to specific color sections; shows remaining count per section
- **Collapsible sections** - Click any color header to collapse/expand; auto-expands on nav click
- **Section progress** - Each section shows remaining items; displays "✓ Complete" when done
- **Card image hover** - See the actual card art when hovering over any card
- **Click-to-mark** - Click any card to mark it as pulled (strikethrough + faded)
- **Expand/Collapse All** - Buttons to quickly manage all sections
- **Reset Progress** - Clear all checked items when starting fresh or regenerating
- **Print-friendly** - Clean output when printing

All progress and collapse states persist in your browser's localStorage.

## Requirements

- Python 3.8+
- pdfplumber
- Internet connection (for Scryfall API)

## License

MIT
