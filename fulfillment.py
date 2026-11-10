"""Bulk fulfillment: tie a TCGplayer shipping export to a pull sheet.

The pull sheet says which cards go to which order; the shipping export has each
order's address, value, and shipping paid. Together they drive the Labels and Pack
steps of the pull sheet report: cheap-shipping Magic-singles orders get numbered
address labels, everything else is set aside to be fulfilled another way.

Which orders count as bulk is decided in the report's JavaScript (the limits are
settings the user can change there); this module only gathers the data.
"""

import csv
import io
import json
from dataclasses import asdict, dataclass
from html import escape as html_escape

import labels

SHIPPING_EXPORT_COLUMNS = {"Order #", "FirstName", "LastName", "Address1", "City", "State",
                           "PostalCode", "Shipping Fee Paid"}


@dataclass
class Shipment:
    order_number: str
    first_name: str
    last_name: str
    address1: str
    address2: str
    city: str
    state: str
    postal_code: str
    country: str
    order_date: str
    item_count: int
    product_value: float
    shipping_fee: float
    tracking_number: str


def csv_kind(text: str) -> str:
    """'pull_sheet', 'shipping_export', or 'unknown', from the header row."""
    header = text.lstrip("﻿").split("\n", 1)[0]
    if "Order Quantity" in header and "Product Name" in header:
        return "pull_sheet"
    if "Order #" in header and "Shipping Fee Paid" in header:
        return "shipping_export"
    return "unknown"


def _money(value: str) -> float:
    try:
        return float((value or "0").replace("$", "").replace(",", "").strip() or 0)
    except ValueError:
        return 0.0


def _int(value: str) -> int:
    try:
        return int(float((value or "0").strip() or 0))
    except ValueError:
        return 0


def parse_shipping_export(text: str) -> list[Shipment]:
    reader = csv.DictReader(io.StringIO(text.lstrip("﻿")))
    missing = SHIPPING_EXPORT_COLUMNS - set(reader.fieldnames or [])
    if missing:
        raise ValueError(f"Not a TCGplayer shipping export (missing columns: {', '.join(sorted(missing))})")

    shipments = []
    for row in reader:
        def get(name: str) -> str:
            return (row.get(name) or "").strip()

        if not get("Order #"):
            continue
        shipments.append(Shipment(
            order_number=get("Order #"),
            first_name=get("FirstName"),
            last_name=get("LastName"),
            address1=get("Address1"),
            address2=get("Address2"),
            city=get("City"),
            state=get("State"),
            postal_code=get("PostalCode"),
            country=get("Country") or "US",
            order_date=get("Order Date"),
            item_count=_int(get("Item Count")),
            product_value=_money(get("Value Of Products")),
            shipping_fee=_money(get("Shipping Fee Paid")),
            tracking_number=get("Tracking #"),
        ))
    return shipments


MAGIC_LINE = "Magic"
MAGIC_SEALED_LINE = "Magic Sealed"
PULL_SHEET_FOOTER = "Orders Contained in Pull Sheet:"


def product_line_of(row: dict) -> str:
    """The pull sheet's Product Line, with sealed Magic split out as "Magic Sealed".

    Sealed Magic (boosters, decks) has no collector number or rarity, which is how
    it's told apart from Magic singles.
    """
    line = (row.get("Product Line") or "").strip()
    if line == MAGIC_LINE and not (row.get("Number") or "").strip() and not (row.get("Rarity") or "").strip():
        return MAGIC_SEALED_LINE
    return line


def _order_quantities(value: str) -> dict[str, int]:
    """"ORDER:QTY | ORDER:QTY" -> {order: qty}."""
    orders: dict[str, int] = {}
    for part in (value or "").split("|"):
        order, sep, qty = part.strip().rpartition(":")
        if sep and order and qty.strip().isdigit():
            orders[order] = orders.get(order, 0) + int(qty)
    return orders


def parse_other_products(text: str) -> list[dict]:
    """Pull sheet rows that aren't Magic singles: other games (Lorcana, Gundam, ...)
    and sealed product. These have no Scryfall data; they're listed as the pull
    sheet gives them so their orders can still be pulled, labeled, and packed."""
    reader = csv.DictReader(io.StringIO(text.lstrip("\ufeff")))
    products = []
    for row in reader:
        line = product_line_of(row)
        if not line or line == MAGIC_LINE or line == PULL_SHEET_FOOTER:
            continue
        orders = _order_quantities(row.get("Order Quantity") or "")
        if not orders:
            continue
        products.append({
            "index": len(products),
            "product_line": line,
            "name": (row.get("Product Name") or "").strip(),
            "set_name": (row.get("Set") or "").strip(),
            "number": (row.get("Number") or "").strip(),
            "condition": (row.get("Condition") or "").strip(),
            "rarity": (row.get("Rarity") or "").strip(),
            "orders": orders,
        })
    return products


def build_orders(pull_sheet_orders: list[dict], shipments: list[Shipment]) -> list[dict]:
    """One record per order in either file, oldest order first.

    This order is what numbers the bulk labels (#1 is the oldest bulk order). The
    export only has the day, so same-day orders keep the shipping export's order;
    orders with no shipping row (no date) come last.
    """
    by_number = {s.order_number: s for s in shipments}
    position = {s.order_number: i for i, s in enumerate(shipments)}
    records = []
    seen = set()
    for o in pull_sheet_orders:
        seen.add(o["order_number"])
        records.append(_order_record(o, by_number.get(o["order_number"])))
    for s in shipments:
        if s.order_number not in seen:
            records.append(_order_record(None, s))
    records.sort(key=lambda r: (
        not r.get("order_date"), r.get("order_date") or "",
        position.get(r["order_number"], len(position)), r["order_number"],
    ))
    return records


def _order_record(summary: dict | None, shipment: Shipment | None) -> dict:
    number = summary["order_number"] if summary else shipment.order_number
    parts = number.split("-")
    record = {
        "order_number": number,
        "short_id": summary["short_id"] if summary else (parts[1] if len(parts) == 3 else number),
        "tcgs": summary["tcgs"] if summary else {},
        "magic_cards": summary["magic_cards"] if summary else 0,
        "in_pull_sheet": summary is not None,
        "has_shipping": shipment is not None,
    }
    if shipment:
        record.update({k: v for k, v in asdict(shipment).items() if k != "order_number"})
    return record


def card_records(cards_in_page_order: list, badges_of=lambda variant: [], price_of=lambda card: None) -> list[dict]:
    """Card rows as the report renders them; index matches each row's data-index.

    badges_of gives (css_class, label) per treatment, reusing the pull sheet's badge
    styles; price_of gives an approximate market price per copy.
    """
    return [{
        "index": i,
        "name": c.card_name,
        "set_name": c.set_name,
        "number": c.collector_number,
        "condition": c.condition,
        "foil": c.is_foil,
        "variant": c.variant,
        "badges": [{"cls": cls, "label": label} for cls, label in badges_of(c.variant)],
        "language": c.language,
        "price": price_of(c),
        "image": c.image_url,
        "orders": c.orders or {},
    } for i, c in enumerate(cards_in_page_order)]


def report_data(orders: list[dict], cards: list[dict], products: list[dict] = ()) -> str:
    """JSON for the report, safe to drop inside a <script> tag."""
    data = {
        "orders": orders,
        "cards": cards,
        "products": list(products),
        "layouts": [p.describe() for p in labels.PRESETS],
        "default_layout": labels.DEFAULT_LAYOUT,
    }
    return json.dumps(data).replace("</", "<\\/")


def steps_nav(order_count: int) -> str:
    return f"""
    <nav class="steps" id="steps">
        <a class="step-tab done" href="/" title="Start over with new files"><span class="step-num">1</span> Upload</a>
        <button class="step-tab" data-step="pull"><span class="step-num">2</span> Pull</button>
        <button class="step-tab" data-step="labels"><span class="step-num">3</span> Labels <span class="step-count" id="labels-count"></span></button>
        <button class="step-tab" data-step="pack"><span class="step-num">4</span> Pack <span class="step-count" id="pack-count"></span></button>
    </nav>
    <div class="lane-filter" id="lane-filter" role="group" aria-label="Show orders">
        <span class="lane-filter-label">Show</span>
        <button class="lane-btn" data-lane="all">All orders <span id="lane-count-all"></span></button>
        <button class="lane-btn" data-lane="bulk">Bulk <span id="lane-count-bulk"></span></button>
        <button class="lane-btn" data-lane="other">Other <span id="lane-count-other"></span></button>
    </div>
    <div class="bulk-banner" id="bulk-banner">{order_count} orders in this batch</div>
"""


def step_panels() -> str:
    """Labels and Pack steps for every order. Their content is filled in by FULFILLMENT_JS."""
    return """
    <div class="ff-menu" id="ff-menu" role="menu" hidden></div>

    <section class="step-panel" id="step-labels" hidden>
        <div class="labels-layout">
            <aside class="ff-card ff-settings">
                <h2>Label settings</h2>
                <label class="ff-field">Sheet layout
                    <select id="ff-layout"></select>
                    <span class="ff-hint" id="ff-layout-hint"></span>
                </label>
                <div id="ff-custom" hidden>
                    <div class="ff-row2">
                        <label class="ff-field">Sheet width (in)<input id="ff-c-page_width" type="number" step="0.125" min="0.5" max="12"></label>
                        <label class="ff-field">Sheet height (in)<input id="ff-c-page_height" type="number" step="0.125" min="0.5" max="12"></label>
                        <label class="ff-field">Columns<input id="ff-c-cols" type="number" min="1" max="10"></label>
                        <label class="ff-field">Rows<input id="ff-c-rows" type="number" min="1" max="10"></label>
                    </div>
                    <label class="ff-check"><input type="checkbox" id="ff-c-rotate"> Lay labels sideways on the sheet</label>
                </div>
                <label class="ff-field">Inner margin (in)<input id="ff-margin" type="number" step="0.02" min="0" max="1"></label>
                <label class="ff-field">Matching line on each label
                    <select id="ff-id_line">
                        <option value="num_id_count">#4 · F1175B · 2x</option>
                        <option value="id_count">F1175B · 2x</option>
                        <option value="id">F1175B</option>
                        <option value="none">Nothing</option>
                    </select>
                </label>
                <label class="ff-check"><input type="checkbox" id="ff-uppercase"> Address in capitals (what USPS prefers)</label>
                <label class="ff-check"><input type="checkbox" id="ff-cut_lines"> Dashed cut lines between labels</label>
                <label class="ff-field">Return address (optional)
                    <textarea id="ff-return_address" rows="3" placeholder="Your name&#10;Street&#10;City, ST ZIP"></textarea>
                </label>

                <h2>Which orders are bulk</h2>
                <p class="ff-hint">Bulk is cheap-shipping, Magic-singles orders. Every other order is in the Other lane, with the reason it isn't bulk.</p>
                <label class="ff-field">Needs tracking if shipping paid is over
                    <span class="ff-money">$<input id="ff-tracking_limit" type="number" step="0.01" min="0"></span>
                    <span class="ff-hint">Above this the buyer paid for tracking.</span>
                </label>
                <label class="ff-field">Needs tracking if the order is worth at least
                    <span class="ff-money">$<input id="ff-value_limit" type="number" step="1" min="0"></span>
                </label>
                <label class="ff-field">Flag orders with more than
                    <span class="ff-money"><input id="ff-big_order" type="number" min="1"> cards</span>
                    <span class="ff-hint">Still bulk; just might not fit a stamped envelope.</span>
                </label>
                <p class="ff-hint">Orders that need tracking get a "printed elsewhere" checkbox instead of an address label. Orders with sealed product or other games, shipping outside the US, or missing from either file are Other too.</p>
                <button class="nav-btn" id="ff-reset-settings">Reset settings</button>
            </aside>

            <div class="ff-card">
                <div class="ff-toolbar">
                    <h2>Orders</h2>
                    <span class="ff-spacer"></span>
                    <label class="ff-check"><input type="checkbox" id="ff-hide-printed"> Hide printed</label>
                    <button class="nav-btn" id="ff-select-all">Select all</button>
                    <button class="nav-btn" id="ff-select-none">Select none</button>
                </div>
                <table class="ff-table">
                    <thead><tr><th></th><th class="num">Label</th><th>Order</th><th>Ship to</th><th class="num">Items</th><th class="num">Value</th><th class="num">Shipping</th><th>Status</th></tr></thead>
                    <tbody id="ff-label-rows"></tbody>
                </table>
                <p class="ff-empty" id="ff-label-empty" hidden></p>
            </div>

            <div class="ff-card ff-preview">
                <div class="ff-toolbar">
                    <h2>Preview</h2>
                    <span class="ff-spacer"></span>
                    <span class="ff-hint" id="ff-sheet-count"></span>
                </div>
                <iframe id="ff-preview" title="Label preview" hidden></iframe>
                <div class="ff-preview-empty" id="ff-preview-empty">Select orders to see their labels.</div>
                <div class="ff-toolbar ff-actions">
                    <button class="nav-btn primary" id="ff-open-pdf" disabled>Open PDF to print</button>
                    <button class="nav-btn" id="ff-mark-printed" disabled>Mark selected as printed</button>
                    <button class="nav-btn danger" id="ff-clear-printed">Clear printed marks</button>
                </div>
                <ul class="ff-tips">
                    <li>Print at 100% / "Actual size", paper 4 x 6 (or your roll's size).</li>
                    <li>Print one sheet first and check the cut lines before printing the batch.</li>
                </ul>
            </div>
        </div>
    </section>

    <section class="step-panel" id="step-pack" hidden>
        <div class="ff-toolbar">
            <h2>Pack each envelope</h2>
            <span class="ff-hint" id="ff-pack-progress"></span>
            <span class="ff-spacer"></span>
            <label class="ff-check">Top-load cards worth $<input id="ff-topload" type="number" step="0.5" min="0" class="ff-topload-input"> or more</label>
            <label class="ff-check"><input type="checkbox" id="ff-hide-packed"> Hide packed</label>
            <button class="nav-btn danger" id="ff-reset-pack">Reset packing</button>
        </div>
        <p class="ff-hint">Match each card to its label number. Click a card when it's in the envelope, or tick the envelope's box to mark all its cards in at once. Prices are TCGplayer market prices (via Scryfall) per copy, a guide to which cards need a toploader, not what the buyer paid. Right-click a card to mark it missing.</p>
        <div class="ff-refunds" id="ff-refunds" hidden></div>
        <div class="envelopes" id="ff-envelopes"></div>
    </section>

"""


# End of the Pull step: products that aren't Magic singles, filled in by FULFILLMENT_JS.
OTHER_PULL_HTML = """
    <section class="other-pull" id="ff-other-pull" hidden>
        <div class="other-pull-head">
            <h2>Other products</h2>
            <span class="ff-hint" id="ff-other-progress"></span>
        </div>
        <p class="ff-hint">Other games and sealed product from the pull sheet, by game and set. Click a row once it's pulled; right-click to mark it missing.</p>
        <div id="ff-other-groups"></div>
    </section>
"""


FULFILLMENT_CSS = """
        .steps { display: flex; gap: 6px; justify-content: center; flex-wrap: wrap; margin: 6px 0 14px; }
        .step-tab { font: inherit; color: #ccc; background: #16213e; border: 1px solid #2c3a5e; border-radius: 999px;
            padding: 7px 16px; cursor: pointer; text-decoration: none; display: inline-flex; align-items: center; gap: 8px; }
        .step-tab:hover { border-color: #4a5a85; color: #fff; }
        .step-tab.active { background: #e94560; border-color: #e94560; color: #fff; }
        .step-tab.aside:not(.active) { border-style: dashed; }
        .step-num { display: inline-grid; place-items: center; width: 20px; height: 20px; border-radius: 50%;
            background: rgba(255,255,255,0.12); font-size: 0.8em; font-weight: 700; }
        .step-tab.done .step-num::after { content: ''; }
        .step-count { font-size: 0.85em; opacity: 0.85; }
        .bulk-banner { text-align: center; color: #aaa; margin-bottom: 12px; }
        .lane-filter { display: flex; gap: 6px; justify-content: center; align-items: center; margin-bottom: 18px; }
        .lane-filter[hidden] { display: none; }
        .lane-filter-label { color: #8a93a8; font-size: 0.85em; margin-right: 4px; }
        .lane-btn { font: inherit; font-size: 0.9em; color: #ccc; background: transparent; border: 1px solid #2c3a5e; border-radius: 6px; padding: 4px 12px; cursor: pointer; }
        .lane-btn:hover { color: #fff; border-color: #4a5a85; }
        .lane-btn.active { background: #2c3a5e; color: #fff; border-color: #4a5a85; }
        .ff-chip.lane-bulk { color: #8fdc93; background: rgba(76,175,80,0.15); }
        .ff-chip.lane-other { color: #f0b35a; background: rgba(240,179,90,0.15); }
        .ff-chip.reason { color: #c9cfdc; background: rgba(255,255,255,0.07); }
        .ff-chip.track { color: #9ecbff; background: rgba(74,158,255,0.18); }
        .label-no-big.other, .env-no.other { color: #f0b35a; }
        .envelope.lane-other-env { border: 1px dashed #6b5326; }
        .envelope.lane-other-env.packed { border-color: #2f6f3a; }
        .elsewhere { margin: 0 6px 3px 0; display: inline-flex; font-size: 0.85em; color: #9ecbff; }
        .card-variant.variant-product { background: rgba(240,179,90,0.18); color: #f0b35a; }

        .other-pull { background: #16213e; border-radius: 12px; padding: 16px 20px; margin: 30px 0; }
        .other-pull[hidden] { display: none; }
        .other-pull-head { display: flex; align-items: baseline; gap: 12px; }
        .other-pull h2 { margin: 0; font-size: 1.3em; color: #fff; }
        .other-game h3 { color: #f0b35a; margin: 18px 0 6px; font-size: 1.05em; }
        .other-set-name { color: #8a93a8; font-size: 0.85em; text-transform: uppercase; letter-spacing: 0.04em; margin: 10px 0 4px; }
        .other-item { display: flex; gap: 12px; padding: 8px 10px; border-radius: 8px; background: rgba(255,255,255,0.03); margin-bottom: 6px; cursor: pointer; align-items: baseline; }
        .other-item:hover { background: rgba(255,255,255,0.06); }
        .other-item.pulled { opacity: 0.45; }
        .other-item.pulled .name { text-decoration: line-through; }
        .other-item.has-missing { box-shadow: inset 4px 0 0 #ef5350; }
        .other-item .qty { font-weight: 700; color: #e94560; min-width: 3ch; }
        .other-item .name { color: #fff; }
        .other-item .meta { color: #8a93a8; font-size: 0.85em; }
        .other-body { flex: 1; min-width: 0; }
        .step-panel[hidden] { display: none; }
        .step-panel h2 { font-size: 1.1em; margin: 0; color: #fff; }

        .order-chip.bulk-label { border-color: #4caf50; }
        .order-chip .label-no { font-weight: 700; color: #8fdc93; margin-right: 4px; }
        .order-chip.set-aside { border-style: dashed; color: #f0b35a; }
        .order-chip.set-aside .label-no { color: #f0b35a; }

        .labels-layout { display: grid; grid-template-columns: 320px minmax(0, 1fr) 400px; gap: 18px; align-items: start; }
        @media (max-width: 1350px) { .labels-layout { grid-template-columns: 320px minmax(0, 1fr); } .ff-preview { grid-column: 1 / -1; } }
        .ff-card { background: #16213e; border-radius: 12px; padding: 16px 18px; }
        .ff-settings h2 { margin: 4px 0 12px; }
        .ff-settings h2 + * { margin-top: 0; }
        .ff-settings h2:not(:first-child) { margin-top: 22px; }
        .ff-field { display: grid; gap: 4px; margin-bottom: 12px; font-size: 0.8em; color: #aaa; text-transform: uppercase; letter-spacing: 0.03em; }
        .ff-field input, .ff-field select, .ff-field textarea { text-transform: none; letter-spacing: 0; font-size: 1.2em; }
        .ff-hint { font-size: 0.85rem; color: #8a93a8; text-transform: none; letter-spacing: 0; }
        .ff-check { display: flex; gap: 8px; align-items: center; margin-bottom: 8px; color: #ddd; cursor: pointer; }
        .ff-row2 { display: grid; grid-template-columns: 1fr 1fr; gap: 8px; }
        .ff-money { display: flex; align-items: center; gap: 4px; color: #ddd; font-size: 1.2em; text-transform: none; }
        .ff-money input { width: 90px; font-size: 1em; }
        .ff-card input, .ff-card select, .ff-card textarea { background: #0f1830; color: #eee; border: 1px solid #2c3a5e;
            border-radius: 6px; padding: 6px 8px; font-family: inherit; }
        .ff-card select, .ff-card textarea { width: 100%; }
        .ff-card input[type=checkbox] { accent-color: #e94560; width: 16px; height: 16px; }
        .ff-toolbar { display: flex; gap: 8px; align-items: center; flex-wrap: wrap; margin-bottom: 12px; }
        .ff-toolbar .ff-check { margin: 0; }
        .ff-spacer { flex: 1; }
        .nav-btn.primary { background: #e94560; border-color: #e94560; color: #fff; }
        .nav-btn:disabled { opacity: 0.45; cursor: default; }
        .ff-table { width: 100%; border-collapse: collapse; font-size: 0.92em; }
        .ff-table th, .ff-table td { text-align: left; padding: 7px 6px; border-bottom: 1px solid #26314f; vertical-align: top; }
        .ff-table th { color: #8a93a8; font-size: 0.8em; text-transform: uppercase; letter-spacing: 0.03em; font-weight: 600; }
        .ff-table .num { text-align: right; font-variant-numeric: tabular-nums; }
        .ff-table tr.off td:not(:first-child) { opacity: 0.4; }
        .ff-table .sub { color: #8a93a8; font-size: 0.85em; }
        .label-no-big { font-weight: 700; color: #8fdc93; font-size: 1.1em; }
        .ff-chip { display: inline-block; font-size: 0.8em; border-radius: 999px; padding: 1px 9px; margin: 0 4px 3px 0; white-space: nowrap; }
        .ff-chip.ok { color: #8fdc93; background: rgba(76,175,80,0.15); }
        .ff-chip.warn { color: #f0b35a; background: rgba(240,179,90,0.15); }
        .ff-chip.done { color: #aaa; background: rgba(255,255,255,0.08); }
        .ff-empty { color: #8a93a8; padding: 10px 4px; }
        #ff-preview { width: 100%; height: 620px; border: 0; border-radius: 8px; background: #fff; }
        .ff-preview-empty { height: 620px; display: grid; place-items: center; text-align: center; padding: 20px;
            color: #8a93a8; border: 1px dashed #2c3a5e; border-radius: 8px; }
        .ff-actions { margin: 12px 0 0; }
        .ff-tips { color: #8a93a8; font-size: 0.85em; padding-left: 18px; margin: 10px 0 0; }

        .envelopes { display: grid; grid-template-columns: repeat(auto-fill, minmax(330px, 1fr)); gap: 14px; margin-top: 14px; }
        .envelope { background: #16213e; border-radius: 12px; padding: 12px 14px; border: 1px solid transparent; }
        .envelope.packed { opacity: 0.55; border-color: #2f6f3a; }
        .env-head { display: flex; align-items: baseline; gap: 10px; margin-bottom: 6px; }
        .env-no { font-size: 1.6em; font-weight: 800; color: #8fdc93; min-width: 2ch; }
        .env-title { flex: 1; min-width: 0; }
        .env-title b { color: #fff; }
        .env-status { font-size: 0.85em; color: #8a93a8; white-space: nowrap; }
        .envelope.packed .env-status { color: #8fdc93; }
        .env-done { display: inline-flex; align-items: center; gap: 6px; cursor: pointer; }
        .env-done input { accent-color: #4caf50; width: 18px; height: 18px; cursor: pointer; }
        .env-cards { list-style: none; margin: 6px 0 0; padding: 0; }
        .env-card { display: flex; gap: 8px; padding: 5px 6px; border-radius: 6px; cursor: pointer; align-items: baseline; }
        .env-card:hover { background: rgba(255,255,255,0.05); }
        .env-card.in { text-decoration: line-through; opacity: 0.5; }
        .env-card .qty { font-weight: 700; min-width: 2.5ch; }
        .env-card .meta { color: #8a93a8; font-size: 0.85em; }
        .env-card-body { flex: 1; min-width: 0; }
        .env-card .card-variant { font-size: 0.72em; }
        .env-price { margin-left: auto; padding-left: 8px; color: #8a93a8; font-size: 0.85em; white-space: nowrap; font-variant-numeric: tabular-nums; }
        .env-price.top { color: #1a1a2e; background: #ffd54f; border-radius: 999px; padding: 1px 8px; font-weight: 700; }
        .env-card.valuable { box-shadow: inset 3px 0 0 #ffd54f; }
        .env-value { color: #ddd; font-weight: 600; }
        .env-tops { color: #ffd54f; font-size: 0.85em; font-weight: 600; margin-top: 2px; }
        .ff-topload-input { width: 70px; background: #0f1830; color: #eee; border: 1px solid #2c3a5e; border-radius: 6px; padding: 4px 6px; font: inherit; }
        .env-card.static { cursor: default; }
        .env-other { color: #f0b35a; font-size: 0.9em; margin-top: 6px; }
        .env-card.gone { opacity: 0.55; }
        .env-card.gone .qty, .env-card.gone > span:nth-child(2) > :first-child { text-decoration: line-through; }
        .missing-note { color: #ff7b7b; font-size: 0.85em; font-weight: 600; margin-left: 6px; text-decoration: none; }
        .ff-chip.bad { color: #ff9b9b; background: rgba(239,83,80,0.18); }

        .card-item.has-missing { box-shadow: inset 4px 0 0 #ef5350; }
        .missing-badge { display: inline-block; margin-left: 8px; font-size: 0.75em; font-weight: 700; color: #fff;
            background: #c62828; border-radius: 999px; padding: 1px 8px; vertical-align: middle; }

        .ff-menu { position: fixed; z-index: 2000; min-width: 260px; max-width: 340px; background: #0f1830; color: #eee;
            border: 1px solid #3a4a75; border-radius: 10px; box-shadow: 0 10px 30px rgba(0,0,0,0.5); padding: 10px 12px; font-size: 0.92em; }
        .ff-menu[hidden] { display: none; }
        .ff-menu h3 { margin: 0 0 2px; font-size: 1em; color: #fff; }
        .ff-menu .meta { color: #8a93a8; font-size: 0.85em; }
        .ff-menu-row { display: flex; align-items: center; gap: 8px; padding: 6px 0; border-top: 1px solid #26314f; }
        .ff-menu .ff-menu-q { margin-bottom: 6px; }
        .ff-menu-row .who { flex: 1; min-width: 0; }
        .ff-menu-row button { font: inherit; width: 28px; height: 28px; border-radius: 6px; border: 1px solid #3a4a75;
            background: #16213e; color: #eee; cursor: pointer; }
        .ff-menu-row button:disabled { opacity: 0.35; cursor: default; }
        .ff-menu-row .count { min-width: 5.5em; text-align: center; font-variant-numeric: tabular-nums; }
        .ff-menu-row .count.some { color: #ff7b7b; font-weight: 700; }
        .ff-menu-actions { display: flex; gap: 8px; justify-content: flex-end; margin-top: 8px; }

        .ff-refunds { background: rgba(198,40,40,0.12); border: 1px solid #7a2a2a; border-radius: 12px; padding: 12px 14px; margin-top: 14px; }
        .ff-refunds[hidden] { display: none; }
        .ff-refunds h3 { margin: 0 0 8px; font-size: 1em; color: #ffb4b4; }
        .refund { display: grid; grid-template-columns: 1fr auto; gap: 4px 12px; padding: 8px 0; border-top: 1px solid rgba(255,255,255,0.08); }
        .refund:first-of-type { border-top: 0; }
        .refund .who b { color: #fff; }
        .refund ul { margin: 2px 0 0; padding-left: 18px; color: #ddd; }
        .env-reasons { margin-top: 2px; }
"""


FULFILLMENT_JS = r"""
    // ---- Order fulfillment: every order, in two lanes (Bulk / Other) -----------
    // Bulk: cheap-shipping, Magic-singles-only orders (stamped envelopes). Other:
    // everything else, with the reasons why. Both lanes share the Labels and Pack
    // steps; orders that need tracking get a "printed elsewhere" checkbox instead
    // of an address label.
    (function () {
        const DATA = JSON.parse(document.getElementById('fulfillment-data').textContent);
        const KEY = 'bulk-' + generationId + '-';
        const DEFAULTS = {
            layout: DATA.default_layout, margin: 0.1, id_line: 'num_id_count', uppercase: true, cut_lines: true,
            return_address: '', tracking_limit: 5.0, value_limit: 50, big_order: 15, hide_printed: false, topload: 5,
            lane: 'all', custom: { page_width: 4, page_height: 6, cols: 2, rows: 2, rotate: true },
        };
        const load = (k, fallback) => { try { return JSON.parse(localStorage.getItem(k)) ?? fallback; } catch (e) { return fallback; } };
        // Label settings carry over between batches; marks are per batch.
        let settings = { ...DEFAULTS, ...load('bulk-settings', {}) };
        settings.custom = { ...DEFAULTS.custom, ...(settings.custom || {}) };
        let printed = load(KEY + 'printed', {});    // order -> when its label was printed (here or elsewhere)
        let packed = load(KEY + 'packed', {});      // "order|itemId" -> true
        let missing = load(KEY + 'missing', {});    // "order|itemId" -> copies not found
        let pulledOther = load(KEY + 'pulled-products', {});  // product item id -> true
        let excluded = new Set(load(KEY + 'excluded', []));
        // Label numbers are frozen once labels are printed, so changing a setting
        // afterwards can't renumber envelopes away from their printed labels.
        let locked = load(KEY + 'numbers', null);   // order -> label number
        let previewUrl = null, previewTimer = null, previewSeq = 0;

        const $ = (id) => document.getElementById(id);
        const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
        const money = (n) => '$' + Number(n || 0).toFixed(2);
        const save = () => {
            localStorage.setItem('bulk-settings', JSON.stringify(settings));
            localStorage.setItem(KEY + 'printed', JSON.stringify(printed));
            localStorage.setItem(KEY + 'packed', JSON.stringify(packed));
            localStorage.setItem(KEY + 'missing', JSON.stringify(missing));
            localStorage.setItem(KEY + 'pulled-products', JSON.stringify(pulledOther));
            localStorage.setItem(KEY + 'excluded', JSON.stringify([...excluded]));
            localStorage.setItem(KEY + 'numbers', JSON.stringify(locked));
        };

        // ---- Items: Magic cards (Scryfall data) and other products (pull sheet only) ----
        // Card ids are their pull sheet row index (kept so earlier marks still match);
        // other products are "p<index>".
        const ITEMS = {};
        DATA.cards.forEach((c) => ITEMS[c.index] = { ...c, id: String(c.index), kind: 'card', product_line: 'Magic' });
        (DATA.products || []).forEach((p) => ITEMS['p' + p.index] = {
            ...p, id: 'p' + p.index, kind: 'product', foil: /foil/i.test(p.condition), badges: [], price: null,
        });
        const itemsByOrder = {};
        Object.values(ITEMS).forEach((it) => Object.entries(it.orders).forEach(([o, q]) => (itemsByOrder[o] ||= []).push({ item: it, qty: q })));
        const ordersByNumber = Object.fromEntries(DATA.orders.map((o) => [o.order_number, o]));
        const lineKey = (order, item) => order + '|' + item.id;
        const missingOf = (order, item) => missing[lineKey(order, item)] || 0;
        const linesOf = (o) => itemsByOrder[o.order_number] || [];
        const missingTotal = (o) => linesOf(o).reduce((n, x) => n + missingOf(o.order_number, x.item), 0);
        // What actually ships: the order's items less any that couldn't be found.
        const itemCount = (o) => (o.item_count || linesOf(o).reduce((n, x) => n + x.qty, 0)) - missingTotal(o);
        const productLabel = (line) => line === 'Magic Sealed' ? 'Sealed Magic' : line;

        // ---- Lanes ---------------------------------------------------------------------
        const needsTracking = (o) => !!o.tracking_number
            || (o.has_shipping && (o.shipping_fee > settings.tracking_limit || o.product_value >= settings.value_limit));
        function reasons(o) {
            const r = [];
            if (!o.has_shipping) r.push('Not in the shipping export');
            if (!o.in_pull_sheet) r.push('Not in the pull sheet');
            if (o.tracking_number) r.push('Already has tracking');
            if (o.has_shipping && o.shipping_fee > settings.tracking_limit) r.push('Paid ' + money(o.shipping_fee) + ' shipping');
            if (o.has_shipping && o.product_value >= settings.value_limit) r.push('Worth ' + money(o.product_value));
            Object.entries(o.tcgs).forEach(([line, qty]) => { if (line !== 'Magic') r.push('Has ' + productLabel(line) + ' ×' + qty); });
            if (o.has_shipping && (o.country || 'US').toUpperCase() !== 'US') r.push('Ships to ' + o.country);
            return r;
        }
        const nothingToShip = (o) => o.in_pull_sheet && linesOf(o).length > 0 && itemCount(o) <= 0;
        // An address label from this app: has an address, no tracking needed, something to send.
        const printable = (o) => o.has_shipping && !needsTracking(o) && !nothingToShip(o);

        let all = [], laneOf = {}, why = {}, labelNo = {};
        function classify() {
            laneOf = {}; why = {}; labelNo = {};
            DATA.orders.forEach((o) => { why[o.order_number] = reasons(o); laneOf[o.order_number] = why[o.order_number].length ? 'other' : 'bulk'; });
            // Bulk first, then Other; each oldest first (DATA.orders is in date order).
            all = [...DATA.orders.filter((o) => laneOf[o.order_number] === 'bulk'), ...DATA.orders.filter((o) => laneOf[o.order_number] === 'other')];
            if (!locked) {
                all.forEach((o, i) => labelNo[o.order_number] = i + 1);
                return;
            }
            // Locked: printed numbers stay put; orders without one go after them.
            let next = Math.max(0, ...Object.values(locked));
            all.forEach((o) => labelNo[o.order_number] = locked[o.order_number] || ++next);
            all.sort((a, b) => labelNo[a.order_number] - labelNo[b.order_number]);
        }
        const inLane = (o) => settings.lane === 'all' || laneOf[o.order_number] === settings.lane;
        const shown = () => all.filter(inLane);
        const flags = (o) => {
            const f = [];
            if (itemCount(o) > settings.big_order) f.push(itemCount(o) + ' items');
            if (missingTotal(o)) f.push(missingTotal(o) + ' missing');
            if (nothingToShip(o)) f.push('Nothing left to ship: refund');
            return f;
        };
        const selected = () => shown().filter((o) => printable(o) && !excluded.has(o.order_number) && !(settings.hide_printed && printed[o.order_number]));
        const laneChip = (o) => laneOf[o.order_number] === 'bulk'
            ? '<span class="ff-chip lane-bulk">Bulk</span>' : '<span class="ff-chip lane-other">Other</span>';

        // ---- Steps and lane filter ---------------------------------------------------
        function showStep(step) {
            document.querySelectorAll('.step-tab[data-step]').forEach((t) => t.classList.toggle('active', t.dataset.step === step));
            document.querySelectorAll('.step-panel').forEach((el) => el.hidden = el.id !== 'step-' + step);
            $('lane-filter').hidden = step === 'pull';
            localStorage.setItem(KEY + 'step', JSON.stringify(step));
            if (step === 'labels') schedulePreview();
        }
        document.querySelectorAll('.step-tab[data-step]').forEach((t) => t.addEventListener('click', () => showStep(t.dataset.step)));
        document.querySelectorAll('.lane-btn').forEach((b) => b.addEventListener('click', () => {
            settings.lane = b.dataset.lane;
            save();
            render();
        }));

        // ---- Pull step: where each Magic card goes, and the other products --------------
        function decorateChips() {
            document.querySelectorAll('.card-item .order-chip[data-order]').forEach((chip) => {
                const o = chip.dataset.order;
                const rec = ordersByNumber[o] || {};
                const bulk = laneOf[o] === 'bulk';
                chip.classList.toggle('bulk-label', bulk);
                chip.classList.toggle('set-aside', !bulk);
                chip.innerHTML = '<span class="label-no">#' + labelNo[o] + '</span>' + esc(rec.short_id || o) + ' ×' + esc(chip.dataset.qty);
                chip.title = '#' + labelNo[o] + (bulk ? ' · Bulk' : ' · Other: ' + (why[o] || []).join(', ')) + ' · ' + o;
            });
        }
        function decorateMissing() {
            document.querySelectorAll('.card-item[data-index]').forEach((row) => {
                const item = ITEMS[row.dataset.index];
                if (!item) return;
                const n = Object.keys(item.orders).reduce((sum, o) => sum + missingOf(o, item), 0);
                row.classList.toggle('has-missing', n > 0);
                let badge = row.querySelector('.missing-badge');
                if (n > 0 && !badge) {
                    badge = document.createElement('span');
                    badge.className = 'missing-badge';
                    row.querySelector('.card-name').appendChild(badge);
                }
                if (badge) n > 0 ? badge.textContent = n + ' missing' : badge.remove();
            });
        }
        function orderChip(order, qty) {
            const rec = ordersByNumber[order] || {};
            const bulk = laneOf[order] === 'bulk';
            return '<span class="order-chip ' + (bulk ? 'bulk-label' : 'set-aside') + '" title="' + esc('#' + labelNo[order] + ' · ' + order) + '">'
                + '<span class="label-no">#' + labelNo[order] + '</span>' + esc(rec.short_id || order) + ' ×' + qty + '</span>';
        }
        function renderOtherPull() {
            const products = Object.values(ITEMS).filter((it) => it.kind === 'product');
            $('ff-other-pull').hidden = products.length === 0;
            if (!products.length) return;
            const groups = {};
            products.forEach((p) => ((groups[productLabel(p.product_line)] ||= {})[p.set_name || 'No set'] ||= []).push(p));
            const done = products.filter((p) => pulledOther[p.id]).length;
            $('ff-other-progress').textContent = done + ' / ' + products.length + ' pulled';
            $('ff-other-groups').innerHTML = Object.keys(groups).sort().map((game) => '<div class="other-game"><h3>' + esc(game) + '</h3>'
                + Object.keys(groups[game]).sort().map((set) => '<div class="other-set"><div class="other-set-name">' + esc(set) + '</div>'
                    + groups[game][set].sort((a, b) => a.name.localeCompare(b.name)).map((p) => {
                        const total = Object.values(p.orders).reduce((n, q) => n + q, 0);
                        const miss = Object.keys(p.orders).reduce((n, o) => n + missingOf(o, p), 0);
                        return '<div class="other-item' + (pulledOther[p.id] ? ' pulled' : '') + (miss ? ' has-missing' : '') + '" data-pull-product="' + esc(p.id) + '" data-item-row="' + esc(p.id) + '">'
                            + '<span class="qty">' + total + 'x</span><span class="other-body"><span class="name">' + esc(p.name) + '</span>'
                            + (miss ? '<span class="missing-badge">' + miss + ' missing</span>' : '')
                            + '<div class="meta">' + esc([p.number && '#' + p.number, p.rarity, p.condition].filter(Boolean).join(' · ')) + '</div>'
                            + '<div class="card-orders">' + Object.entries(p.orders).map(([o, q]) => orderChip(o, q)).join('') + '</div></span></div>';
                    }).join('') + '</div>').join('') + '</div>').join('');
        }
        document.addEventListener('click', (e) => {
            const row = e.target.closest('[data-pull-product]');
            if (!row) return;
            const id = row.dataset.pullProduct;
            pulledOther[id] ? delete pulledOther[id] : pulledOther[id] = true;
            save();
            renderOtherPull();
        });

        // ---- Missing items: right-click an item (Pull or Pack) --------------------------
        const menu = $('ff-menu');
        let menuItem = null, menuOrders = [];
        function openMenu(item, orders, x, y) {
            menuItem = item;
            menuOrders = orders;
            renderMenu();
            menu.hidden = false;
            const w = menu.offsetWidth, h = menu.offsetHeight;
            menu.style.left = Math.max(8, Math.min(x, window.innerWidth - w - 8)) + 'px';
            menu.style.top = Math.max(8, Math.min(y, window.innerHeight - h - 8)) + 'px';
        }
        function closeMenu() { menu.hidden = true; menuItem = null; }
        function orderName(o) {
            const rec = ordersByNumber[o] || {};
            return '#' + labelNo[o] + ' · ' + esc(rec.short_id || o)
                + (rec.first_name ? ' <span class="meta">' + esc(rec.first_name + ' ' + rec.last_name) + '</span>' : '');
        }
        const itemMeta = (it) => [it.kind === 'product' ? productLabel(it.product_line) : '', it.set_name, it.number && '#' + it.number,
            it.condition + (it.kind === 'card' && it.foil ? ' Foil' : '')].filter(Boolean).join(' · ');
        function renderMenu() {
            const it = menuItem;
            menu.innerHTML = '<h3>' + esc(it.name) + '</h3><div class="meta">' + esc(itemMeta(it))
                + '</div><div class="meta ff-menu-q">How many couldn\'t you find?</div>'
                + menuOrders.map((o) => {
                    const qty = it.orders[o], m = missingOf(o, it);
                    return '<div class="ff-menu-row"><span class="who">' + orderName(o) + '</span>'
                        + '<button data-miss="' + esc(o) + '" data-step="-1"' + (m <= 0 ? ' disabled' : '') + ' aria-label="One fewer missing">−</button>'
                        + '<span class="count' + (m ? ' some' : '') + '">' + m + ' of ' + qty + ' missing</span>'
                        + '<button data-miss="' + esc(o) + '" data-step="1"' + (m >= qty ? ' disabled' : '') + ' aria-label="One more missing">+</button></div>';
                }).join('')
                + '<div class="ff-menu-actions"><button class="nav-btn" data-miss-all="1">All missing</button>'
                + '<button class="nav-btn" data-miss-all="0">None missing</button><button class="nav-btn primary" data-miss-close>Done</button></div>';
        }
        function setMissing(o, item, n) {
            const key = lineKey(o, item);
            if (n > 0) missing[key] = n; else delete missing[key];
        }
        menu.addEventListener('click', (e) => {
            e.stopPropagation();
            const b = e.target.closest('button');
            if (!b || !menuItem) return;
            if (b.dataset.miss) {
                const o = b.dataset.miss;
                setMissing(o, menuItem, Math.max(0, Math.min(menuItem.orders[o], missingOf(o, menuItem) + Number(b.dataset.step))));
            } else if (b.dataset.missAll) {
                menuOrders.forEach((o) => setMissing(o, menuItem, b.dataset.missAll === '1' ? menuItem.orders[o] : 0));
            } else if ('missClose' in b.dataset) {
                closeMenu();
                return;
            }
            save();
            render();
            renderMenu();
        });
        document.addEventListener('contextmenu', (e) => {
            const row = e.target.closest('.card-item[data-index], [data-item-row]');
            const line = e.target.closest('[data-item-line]');
            if (!row && !line) return;
            e.preventDefault();
            if (line) {
                const [o, id] = line.dataset.itemLine.split('|');
                openMenu(ITEMS[id], [o], e.clientX, e.clientY);
            } else {
                const item = ITEMS[row.dataset.index ?? row.dataset.itemRow];
                if (item) openMenu(item, Object.keys(item.orders), e.clientX, e.clientY);
            }
        });
        document.addEventListener('click', (e) => { if (!menu.hidden && !menu.contains(e.target)) closeMenu(); }, true);
        document.addEventListener('keydown', (e) => { if (e.key === 'Escape') closeMenu(); });
        window.addEventListener('scroll', () => { if (!menu.hidden) closeMenu(); }, { passive: true });

        // ---- Labels step --------------------------------------------------------------
        const fields = ['layout', 'margin', 'id_line', 'uppercase', 'cut_lines', 'return_address',
            'tracking_limit', 'value_limit', 'big_order'];
        const read = (el) => el.type === 'checkbox' ? el.checked : el.type === 'number' ? Number(el.value) : el.value;
        const write = (el, v) => { if (el.type === 'checkbox') el.checked = !!v; else el.value = v ?? ''; };

        $('ff-layout').innerHTML = DATA.layouts.map((l) => '<option value="' + l.key + '">' + esc(l.name) + '</option>').join('')
            + '<option value="custom">Custom size…</option>';
        function syncInputs() {
            fields.forEach((f) => write($('ff-' + f), settings[f]));
            Object.keys(DEFAULTS.custom).forEach((k) => write($('ff-c-' + k), settings.custom[k]));
            write($('ff-hide-printed'), settings.hide_printed);
            write($('ff-topload'), settings.topload);
        }
        function currentLayout() {
            if (settings.layout !== 'custom') return DATA.layouts.find((l) => l.key === settings.layout) || DATA.layouts[0];
            const c = settings.custom;
            const gw = c.rotate ? c.page_height : c.page_width, gh = c.rotate ? c.page_width : c.page_height;
            return { per_sheet: c.cols * c.rows, label_width: gw / c.cols, label_height: gh / c.rows };
        }
        function onSettings() {
            fields.forEach((f) => settings[f] = read($('ff-' + f)));
            Object.keys(DEFAULTS.custom).forEach((k) => settings.custom[k] = read($('ff-c-' + k)));
            settings.hide_printed = $('ff-hide-printed').checked;
            save();
            render();
        }
        document.querySelectorAll('#step-labels .ff-settings input, #step-labels .ff-settings select, #step-labels .ff-settings textarea, #ff-hide-printed')
            .forEach((el) => el.addEventListener(el.tagName === 'TEXTAREA' || el.type === 'number' ? 'input' : 'change', onSettings));
        $('ff-reset-settings').addEventListener('click', () => {
            settings = { ...DEFAULTS, custom: { ...DEFAULTS.custom }, lane: settings.lane };
            syncInputs();
            onSettings();
        });

        function labelCell(o) {
            const n = o.order_number;
            if (needsTracking(o)) {
                return '<label class="ff-check elsewhere"><input type="checkbox" data-elsewhere="' + esc(n) + '"' + (printed[n] ? ' checked' : '')
                    + '> Tracked label printed elsewhere</label>';
            }
            if (!o.has_shipping) return '<span class="ff-chip warn">No address</span>';
            if (nothingToShip(o)) return '<span class="ff-chip bad">Nothing to ship</span>';
            return printed[n] ? '<span class="ff-chip done">Printed</span>' : '<span class="ff-chip ok">Label</span>';
        }
        function renderLabels() {
            const l = currentLayout();
            $('ff-custom').hidden = settings.layout !== 'custom';
            $('ff-layout-hint').textContent = l.per_sheet + ' per sheet, each ' + (+l.label_width.toFixed(2)) + ' x ' + (+l.label_height.toFixed(2)) + ' in';
            const visible = shown().filter((o) => !(settings.hide_printed && printed[o.order_number]));
            $('ff-label-rows').innerHTML = visible.map((o) => {
                const n = o.order_number;
                const can = printable(o);
                const on = can && !excluded.has(n);
                const chips = [laneChip(o), labelCell(o),
                    ...(why[n] || []).map((r) => '<span class="ff-chip reason">' + esc(r) + '</span>'),
                    ...flags(o).map((f) => '<span class="ff-chip warn">' + esc(f) + '</span>')].join('');
                return '<tr class="' + (on || !can ? '' : 'off') + '">'
                    + '<td>' + (can ? '<input type="checkbox" data-include="' + esc(n) + '"' + (on ? ' checked' : '') + ' aria-label="Print label for ' + esc(o.short_id) + '">' : '') + '</td>'
                    + '<td class="num"><span class="label-no-big' + (laneOf[n] === 'other' ? ' other' : '') + '">#' + labelNo[n] + '</span></td>'
                    + '<td>' + esc(o.short_id) + '<div class="sub">' + esc(o.order_date || '') + '</div></td>'
                    + '<td>' + (o.has_shipping ? esc(o.first_name) + ' ' + esc(o.last_name) + '<div class="sub">' + esc([o.city, o.state].filter(Boolean).join(', ')) + ' ' + esc(o.postal_code) + '</div>' : '<span class="sub">No shipping info</span>') + '</td>'
                    + '<td class="num">' + itemCount(o) + '</td><td class="num">' + (o.has_shipping ? money(o.product_value) : '') + '</td>'
                    + '<td class="num">' + (o.has_shipping ? money(o.shipping_fee) : '') + '</td><td>' + chips + '</td></tr>';
            }).join('');
            const empty = $('ff-label-empty');
            empty.hidden = visible.length > 0;
            empty.textContent = shown().length ? 'Every label in this view is marked printed.' : 'No orders in this view.';
            const sel = selected();
            const sheets = Math.ceil(sel.length / (l.per_sheet || 1));
            $('ff-sheet-count').textContent = sel.length ? sel.length + ' label' + (sel.length === 1 ? '' : 's') + ' · ' + sheets + ' sheet' + (sheets === 1 ? '' : 's') : '';
            $('ff-mark-printed').disabled = sel.length === 0;
            if (!$('step-labels').hidden) schedulePreview();
        }
        document.addEventListener('change', (e) => {
            const d = e.target.dataset || {};
            if (d.include) {
                e.target.checked ? excluded.delete(d.include) : excluded.add(d.include);
            } else if (d.elsewhere) {
                if (e.target.checked) printed[d.elsewhere] = new Date().toISOString(); else delete printed[d.elsewhere];
                if (e.target.checked) locked = { ...(locked || {}), ...labelNo };
            } else {
                return;
            }
            save();
            render();
        });
        $('ff-select-all').addEventListener('click', () => { shown().forEach((o) => excluded.delete(o.order_number)); save(); render(); });
        $('ff-select-none').addEventListener('click', () => { shown().forEach((o) => excluded.add(o.order_number)); save(); render(); });
        $('ff-mark-printed').addEventListener('click', () => {
            selected().forEach((o) => printed[o.order_number] = new Date().toISOString());
            locked = { ...(locked || {}), ...labelNo };
            save();
            render();
        });
        $('ff-clear-printed').addEventListener('click', () => {
            if (!confirm('Clear the printed mark from every order in this batch? Label numbers will be worked out again.')) return;
            printed = {};
            locked = null;
            save();
            render();
        });

        function schedulePreview() {
            clearTimeout(previewTimer);
            previewTimer = setTimeout(updatePreview, 300);
        }
        async function updatePreview() {
            const sel = selected();
            const seq = ++previewSeq;
            if (!sel.length) { showPreview(null, 'Select orders to see their labels.'); return; }
            const shipments = sel.map((o) => ({ ...o, label_number: labelNo[o.order_number], item_count: itemCount(o) }));
            try {
                const res = await fetch('/api/labels.pdf', {
                    method: 'POST', headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ shipments, settings }),
                });
                if (!res.ok) throw new Error((await res.json().catch(() => ({}))).detail || "Couldn't build the labels.");
                const blob = await res.blob();
                if (seq === previewSeq) showPreview(URL.createObjectURL(blob));
            } catch (err) {
                const offline = location.protocol === 'file:';
                if (seq === previewSeq) showPreview(null, offline
                    ? 'Labels are built by the web app. Run "make serve" and upload the files there to print labels.'
                    : err.message);
            }
        }
        function showPreview(url, message) {
            if (previewUrl) URL.revokeObjectURL(previewUrl);
            previewUrl = url;
            $('ff-preview').hidden = !url;
            $('ff-preview-empty').hidden = !!url;
            $('ff-open-pdf').disabled = !url;
            if (url) $('ff-preview').src = url + '#view=Fit';
            else $('ff-preview-empty').textContent = message;
        }
        $('ff-open-pdf').addEventListener('click', () => previewUrl && window.open(previewUrl, '_blank'));

        // ---- Pack step -------------------------------------------------------------------
        const isTopLoad = (it) => it.price != null && it.price >= settings.topload;
        function itemLine(o, x) {
            const it = x.item;
            const key = lineKey(o.order_number, it);
            const m = missingOf(o.order_number, it);
            const gone = m >= x.qty;
            const note = m ? '<span class="missing-note">' + (gone ? 'missing: refund' : m + ' of ' + x.qty + ' missing') + '</span>' : '';
            // Same colored badges as the pull sheet rows, so treatments and foils stand out.
            const badges = (it.kind === 'product' ? '<span class="card-variant variant-product">' + esc(productLabel(it.product_line)) + '</span>' : '')
                + (it.badges || []).map((b) => '<span class="card-variant ' + esc(b.cls) + '">' + esc(b.label) + '</span>').join('')
                + (it.kind === 'card' && it.foil ? '<span class="card-foil"> ★ FOIL</span>' : '')
                + (it.language ? '<span class="card-language"> [' + esc(it.language) + ']</span>' : '');
            const top = isTopLoad(it) && !gone;
            const price = it.price != null
                ? '<span class="env-price' + (top ? ' top' : '') + '" title="TCGplayer market price per copy (Scryfall, updated daily)">'
                    + (top ? 'Top-load · ' : '') + '~' + money(it.price) + '</span>'
                : '';
            const meta = [it.set_name, it.number && '#' + it.number, it.condition].filter(Boolean).join(' · ');
            return '<li class="env-card' + (gone ? ' gone static' : packed[key] ? ' in' : '') + (top ? ' valuable' : '') + '"'
                + ' data-item-line="' + esc(key) + '"' + (gone ? '' : ' data-pack="' + esc(key) + '"') + (it.image ? ' data-image="' + esc(it.image) + '"' : '') + '>'
                + '<span class="qty">' + (gone ? x.qty : x.qty - m) + 'x</span><span class="env-card-body"><span>' + esc(it.name) + '</span>' + badges + note
                + '<div class="meta">' + esc(meta) + '</div></span>' + price + '</li>';
        }
        // Fully missing lines don't hold an envelope back; partly missing ones still need packing.
        const lineDone = (o, x) => missingOf(o.order_number, x.item) >= x.qty || packed[lineKey(o.order_number, x.item)];
        const isPacked = (o) => linesOf(o).length > 0 && linesOf(o).every((x) => lineDone(o, x));

        function refundText(o, lines) {
            const who = o.has_shipping ? o.first_name + ' ' + o.last_name : 'Buyer';
            return 'Order ' + o.order_number + ' (' + who + '): missing\n'
                + lines.map((x) => '- ' + missingOf(o.order_number, x.item) + 'x ' + x.item.name
                    + (x.item.variant ? ' (' + x.item.variant + ')' : '') + ', ' + itemMeta(x.item)).join('\n');
        }
        function renderRefunds() {
            const refunds = all.map((o) => ({ o, lines: linesOf(o).filter((x) => missingOf(o.order_number, x.item)) })).filter((r) => r.lines.length);
            const box = $('ff-refunds');
            box.hidden = refunds.length === 0;
            box.innerHTML = '<h3>Refunds to send (' + refunds.length + ' order' + (refunds.length === 1 ? '' : 's') + ')</h3>'
                + refunds.map(({ o, lines }) => '<div class="refund"><div class="who"><b>#' + labelNo[o.order_number] + '</b> · '
                    + '<b>' + esc(o.has_shipping ? o.first_name + ' ' + o.last_name : o.short_id) + '</b> <span class="meta">' + esc(o.order_number) + '</span>'
                    + '<ul>' + lines.map((x) => '<li>' + missingOf(o.order_number, x.item) + 'x ' + esc(x.item.name)
                        + (x.item.variant ? ' (' + esc(x.item.variant) + ')' : '') + ' <span class="meta">' + esc(itemMeta(x.item)) + '</span></li>').join('') + '</ul></div>'
                    + '<div><button class="nav-btn" data-copy-refund="' + esc(o.order_number) + '">Copy details</button></div></div>').join('');
            box.querySelectorAll('[data-copy-refund]').forEach((b) => b.addEventListener('click', async () => {
                const r = refunds.find((x) => x.o.order_number === b.dataset.copyRefund);
                try { await navigator.clipboard.writeText(refundText(r.o, r.lines)); b.textContent = 'Copied'; }
                catch (err) { b.textContent = 'Copy failed'; }
                setTimeout(() => { b.textContent = 'Copy details'; }, 1500);
            }));
        }
        function renderPack() {
            const list = shown();
            $('ff-pack-progress').textContent = list.filter(isPacked).length + ' / ' + list.length + ' envelopes packed';
            const hide = $('ff-hide-packed').checked;
            $('ff-envelopes').innerHTML = list.filter((o) => !(hide && isPacked(o))).map((o) => {
                const n = o.order_number;
                const lines = linesOf(o);
                const inCount = lines.filter((x) => lineDone(o, x)).length;
                const miss = missingTotal(o);
                const p = isPacked(o);
                const tops = lines.filter((x) => isTopLoad(x.item) && missingOf(n, x.item) < x.qty)
                    .reduce((t, x) => t + x.qty - missingOf(n, x.item), 0);
                const name = o.has_shipping ? esc(o.first_name) + ' ' + esc(o.last_name) : 'No shipping info';
                const chips = [laneChip(o),
                    needsTracking(o) ? '<span class="ff-chip track">Needs tracking</span>' : '',
                    ...(why[n] || []).filter((r) => !/shipping$|^Worth|tracking/.test(r)).map((r) => '<span class="ff-chip reason">' + esc(r) + '</span>'),
                    miss ? '<span class="ff-chip bad">' + miss + ' missing: refund the buyer</span>' : ''].join('');
                return '<div class="envelope' + (p ? ' packed' : '') + (laneOf[n] === 'other' ? ' lane-other-env' : '') + '">'
                    + '<div class="env-head"><span class="env-no' + (laneOf[n] === 'other' ? ' other' : '') + '">#' + labelNo[n] + '</span>'
                    + '<span class="env-title"><b>' + name + '</b> <span class="meta">' + esc(o.short_id) + ' · ' + itemCount(o) + 'x'
                    + (o.has_shipping ? ' · <span class="env-value">' + money(o.product_value) + '</span>' : '') + '</span>'
                    + (tops ? '<div class="env-tops">' + tops + ' to top-load</div>' : '') + '</span>'
                    + (lines.length ? '<label class="env-status env-done"><input type="checkbox" data-pack-order="' + esc(n) + '"' + (p ? ' checked' : '')
                        + ' aria-label="Envelope #' + labelNo[n] + ' packed"> ' + (p ? 'Packed' : inCount + ' / ' + lines.length + ' in') + '</label>' : '')
                    + '</div><div class="env-reasons">' + chips + '</div>'
                    + (lines.length ? '<ul class="env-cards">' + lines.map((x) => itemLine(o, x)).join('') + '</ul>'
                        : '<div class="env-other">Nothing for this order on the pull sheet</div>') + '</div>';
            }).join('') || '<p class="ff-empty">' + (list.length ? 'Every envelope in this view is packed.' : 'No orders in this view.') + '</p>';
        }
        document.addEventListener('click', (e) => {
            const line = e.target.closest('[data-pack]');
            if (!line) return;
            const key = line.dataset.pack;
            packed[key] ? delete packed[key] : packed[key] = true;
            save();
            render();
        });
        // The envelope checkbox packs (or unpacks) every item in it at once.
        document.addEventListener('change', (e) => {
            const order = e.target.dataset && e.target.dataset.packOrder;
            if (!order) return;
            (itemsByOrder[order] || []).forEach((x) => {
                const key = lineKey(order, x.item);
                if (e.target.checked && missingOf(order, x.item) < x.qty) packed[key] = true;
                else delete packed[key];
            });
            save();
            render();
        });
        $('ff-hide-packed').addEventListener('change', renderPack);
        $('ff-topload').addEventListener('input', () => {
            settings.topload = Number($('ff-topload').value) || 0;
            save();
            render();
        });
        $('ff-reset-pack').addEventListener('click', () => {
            if (!confirm('Clear packing progress for every envelope?')) return;
            packed = {};
            save();
            render();
        });

        function render() {
            classify();
            decorateChips();
            decorateMissing();
            renderOtherPull();
            renderLabels();
            renderPack();
            renderRefunds();
            const bulkCount = all.filter((o) => laneOf[o.order_number] === 'bulk').length;
            const counts = { all: all.length, bulk: bulkCount, other: all.length - bulkCount };
            document.querySelectorAll('.lane-btn').forEach((b) => b.classList.toggle('active', b.dataset.lane === settings.lane));
            Object.entries(counts).forEach(([lane, c]) => $('lane-count-' + lane).textContent = '(' + c + ')');
            const list = shown();
            $('labels-count').textContent = '(' + list.filter(printable).length + ')';
            $('pack-count').textContent = '(' + list.filter(isPacked).length + '/' + list.length + ')';
            const range = (lane) => {
                const nums = all.filter((o) => laneOf[o.order_number] === lane).map((o) => labelNo[o.order_number]);
                return nums.length ? ' (#' + Math.min(...nums) + '–#' + Math.max(...nums) + ')' : '';
            };
            $('bulk-banner').textContent = all.length + ' orders: ' + bulkCount + ' bulk' + range('bulk') + ', ' + counts.other + ' other' + range('other')
                + ' · oldest first' + (locked ? ' · numbers locked since printing' : '');
        }

        syncInputs();
        render();
        const saved = load(KEY + 'step', 'pull');
        showStep(document.getElementById('step-' + saved) ? saved : 'pull');
    })();
"""
