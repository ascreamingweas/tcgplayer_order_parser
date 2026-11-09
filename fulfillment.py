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


def build_orders(pull_sheet_orders: list[dict], shipments: list[Shipment]) -> list[dict]:
    """One record per order in either file, in the pull sheet's order.

    Orders only in the shipping export (nothing to pull) come last.
    """
    by_number = {s.order_number: s for s in shipments}
    records = []
    seen = set()
    for o in pull_sheet_orders:
        seen.add(o["order_number"])
        records.append(_order_record(o, by_number.get(o["order_number"])))
    for s in shipments:
        if s.order_number not in seen:
            records.append(_order_record(None, s))
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


def card_records(cards_in_page_order: list) -> list[dict]:
    """Card rows as the report renders them; index matches each row's data-index."""
    return [{
        "index": i,
        "name": c.card_name,
        "set_name": c.set_name,
        "number": c.collector_number,
        "condition": c.condition,
        "foil": c.is_foil,
        "variant": c.variant,
        "image": c.image_url,
        "orders": c.orders or {},
    } for i, c in enumerate(cards_in_page_order)]


def report_data(orders: list[dict], cards: list[dict]) -> str:
    """JSON for the report, safe to drop inside a <script> tag."""
    data = {
        "orders": orders,
        "cards": cards,
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
        <button class="step-tab aside" data-step="aside">Set aside <span class="step-count" id="aside-count"></span></button>
    </nav>
    <div class="bulk-banner" id="bulk-banner">{order_count} orders in this batch</div>
"""


def step_panels() -> str:
    """Labels, Pack, and Set aside steps. Their content is filled in by FULFILLMENT_JS."""
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
                <label class="ff-field">Set aside if shipping paid is over
                    <span class="ff-money">$<input id="ff-tracking_limit" type="number" step="0.01" min="0"></span>
                    <span class="ff-hint">Above this the buyer paid for tracking.</span>
                </label>
                <label class="ff-field">Set aside if the order is worth at least
                    <span class="ff-money">$<input id="ff-value_limit" type="number" step="1" min="0"></span>
                </label>
                <label class="ff-field">Flag orders with more than
                    <span class="ff-money"><input id="ff-big_order" type="number" min="1"> cards</span>
                    <span class="ff-hint">Still bulk; just might not fit a stamped envelope.</span>
                </label>
                <p class="ff-hint">Orders with sealed product or other games, orders shipping outside the US, and orders missing from either file are always set aside.</p>
                <button class="nav-btn" id="ff-reset-settings">Reset settings</button>
            </aside>

            <div class="ff-card">
                <div class="ff-toolbar">
                    <h2>Bulk orders</h2>
                    <span class="ff-spacer"></span>
                    <label class="ff-check"><input type="checkbox" id="ff-hide-printed"> Hide printed</label>
                    <button class="nav-btn" id="ff-select-all">Select all</button>
                    <button class="nav-btn" id="ff-select-none">Select none</button>
                </div>
                <table class="ff-table">
                    <thead><tr><th></th><th class="num">Label</th><th>Order</th><th>Ship to</th><th class="num">Cards</th><th class="num">Value</th><th class="num">Shipping</th><th>Status</th></tr></thead>
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
            <label class="ff-check"><input type="checkbox" id="ff-hide-packed"> Hide packed</label>
            <button class="nav-btn danger" id="ff-reset-pack">Reset packing</button>
        </div>
        <p class="ff-hint">Match each card to its label number. Click a card when it's in the envelope; an envelope is packed when all its cards are in. Right-click a card to mark it missing.</p>
        <div class="ff-refunds" id="ff-refunds" hidden></div>
        <div class="envelopes" id="ff-envelopes"></div>
    </section>

    <section class="step-panel" id="step-aside" hidden>
        <h2>Set aside: fulfill these another way</h2>
        <p class="ff-hint">Their Magic singles are on the pull sheet with everything else. Keep them out of the bulk envelopes and ship them with tracking or with their other items.</p>
        <div class="envelopes" id="ff-aside"></div>
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
        .bulk-banner { text-align: center; color: #aaa; margin-bottom: 18px; }
        .step-panel[hidden] { display: none; }
        .step-panel h2 { font-size: 1.1em; margin: 0; color: #fff; }

        .order-chip.bulk-label { border-color: #4caf50; }
        .order-chip .label-no { font-weight: 700; color: #8fdc93; margin-right: 4px; }
        .order-chip.set-aside { border-style: dashed; color: #f0b35a; }

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
        .envelope.set-aside { border: 1px dashed #6b5326; }
        .env-head { display: flex; align-items: baseline; gap: 10px; margin-bottom: 6px; }
        .env-no { font-size: 1.6em; font-weight: 800; color: #8fdc93; min-width: 2ch; }
        .env-title { flex: 1; min-width: 0; }
        .env-title b { color: #fff; }
        .env-status { font-size: 0.85em; color: #8a93a8; white-space: nowrap; }
        .envelope.packed .env-status { color: #8fdc93; }
        .env-cards { list-style: none; margin: 6px 0 0; padding: 0; }
        .env-card { display: flex; gap: 8px; padding: 5px 6px; border-radius: 6px; cursor: pointer; align-items: baseline; }
        .env-card:hover { background: rgba(255,255,255,0.05); }
        .env-card.in { text-decoration: line-through; opacity: 0.5; }
        .env-card .qty { font-weight: 700; min-width: 2.5ch; }
        .env-card .meta { color: #8a93a8; font-size: 0.85em; }
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
    // ---- Bulk fulfillment: Labels / Pack / Set aside steps ----------------------
    (function () {
        const DATA = JSON.parse(document.getElementById('fulfillment-data').textContent);
        const KEY = 'bulk-' + generationId + '-';
        const DEFAULTS = {
            layout: DATA.default_layout, margin: 0.1, id_line: 'num_id_count', uppercase: true, cut_lines: true,
            return_address: '', tracking_limit: 5.0, value_limit: 50, big_order: 15, hide_printed: false,
            custom: { page_width: 4, page_height: 6, cols: 2, rows: 2, rotate: true },
        };
        const load = (k, fallback) => { try { return JSON.parse(localStorage.getItem(k)) ?? fallback; } catch (e) { return fallback; } };
        // Label settings carry over between batches; marks are per batch.
        let settings = { ...DEFAULTS, ...load('bulk-settings', {}) };
        settings.custom = { ...DEFAULTS.custom, ...(settings.custom || {}) };
        let printed = load(KEY + 'printed', {});
        let packed = load(KEY + 'packed', {});      // "order|cardIndex" -> true
        let missing = load(KEY + 'missing', {});    // "order|cardIndex" -> copies not found
        let excluded = new Set(load(KEY + 'excluded', []));
        let previewUrl = null, previewTimer = null, previewSeq = 0;

        const $ = (id) => document.getElementById(id);
        const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
        const money = (n) => '$' + Number(n || 0).toFixed(2);
        const save = () => {
            localStorage.setItem('bulk-settings', JSON.stringify(settings));
            localStorage.setItem(KEY + 'printed', JSON.stringify(printed));
            localStorage.setItem(KEY + 'packed', JSON.stringify(packed));
            localStorage.setItem(KEY + 'missing', JSON.stringify(missing));
            localStorage.setItem(KEY + 'excluded', JSON.stringify([...excluded]));
        };

        const cardsByOrder = {};
        DATA.cards.forEach((c) => Object.entries(c.orders).forEach(([o, q]) => (cardsByOrder[o] ||= []).push({ card: c, qty: q })));
        const missingOf = (order, index) => missing[order + '|' + index] || 0;
        const missingTotal = (o) => (cardsByOrder[o.order_number] || []).reduce((n, x) => n + missingOf(o.order_number, x.card.index), 0);
        // What actually ships: the order's items less any cards that couldn't be found.
        const cardCount = (o) => (o.item_count || (cardsByOrder[o.order_number] || []).reduce((n, x) => n + x.qty, 0)) - missingTotal(o);

        // ---- Which orders are bulk ------------------------------------------------
        function reasons(o) {
            const r = [];
            if (!o.has_shipping) r.push('Not in the shipping export');
            if (!o.in_pull_sheet) r.push('Not in the pull sheet');
            if (o.tracking_number) r.push('Already has tracking');
            if (o.has_shipping && o.shipping_fee > settings.tracking_limit) r.push('Paid ' + money(o.shipping_fee) + ' shipping');
            if (o.has_shipping && o.product_value >= settings.value_limit) r.push('Worth ' + money(o.product_value));
            Object.entries(o.tcgs).forEach(([line, qty]) => {
                if (line !== 'Magic') r.push('Has ' + (line === 'Magic Sealed' ? 'sealed Magic' : line) + ' ×' + qty);
            });
            if (o.has_shipping && (o.country || 'US').toUpperCase() !== 'US') r.push('Ships to ' + o.country);
            if (o.in_pull_sheet && o.magic_cards > 0 && cardCount(o) <= 0) r.push('Every card missing: refund it');
            return r;
        }
        let bulk = [], aside = [], labelNo = {};
        function classify() {
            bulk = []; aside = []; labelNo = {};
            DATA.orders.forEach((o) => {
                const r = reasons(o);
                if (r.length) aside.push({ ...o, reasons: r });
                else { bulk.push(o); labelNo[o.order_number] = bulk.length; }
            });
        }
        const flags = (o) => {
            const f = [];
            if (cardCount(o) > settings.big_order) f.push(cardCount(o) + ' cards');
            if (missingTotal(o)) f.push(missingTotal(o) + ' missing');
            return f;
        };
        const selected = () => bulk.filter((o) => !excluded.has(o.order_number) && !(settings.hide_printed && printed[o.order_number]));

        // ---- Steps ------------------------------------------------------------------
        function showStep(step) {
            document.querySelectorAll('.step-tab[data-step]').forEach((t) => t.classList.toggle('active', t.dataset.step === step));
            document.querySelectorAll('.step-panel').forEach((el) => el.hidden = el.id !== 'step-' + step);
            localStorage.setItem(KEY + 'step', JSON.stringify(step));
            if (step === 'labels') schedulePreview();
        }
        document.querySelectorAll('.step-tab[data-step]').forEach((t) => t.addEventListener('click', () => showStep(t.dataset.step)));

        // ---- Pull step: show where each card goes ---------------------------------
        function decorateChips() {
            document.querySelectorAll('.card-item .order-chip[data-order]').forEach((chip) => {
                const o = chip.dataset.order;
                const qty = chip.dataset.qty;
                const short = (DATA.orders.find((x) => x.order_number === o) || {}).short_id || o;
                const no = labelNo[o];
                chip.classList.toggle('bulk-label', !!no);
                chip.classList.toggle('set-aside', !no);
                chip.innerHTML = (no ? '<span class="label-no">#' + no + '</span>' : '') + esc(short) + ' ×' + esc(qty);
                const a = aside.find((x) => x.order_number === o);
                chip.title = no ? 'Label #' + no + ' · ' + o : 'Set aside: ' + (a ? a.reasons.join(', ') : '') + ' · ' + o;
            });
        }

        function decorateMissing() {
            document.querySelectorAll('.card-item[data-index]').forEach((item) => {
                const card = DATA.cards[+item.dataset.index];
                if (!card) return;
                const n = Object.keys(card.orders).reduce((sum, o) => sum + missingOf(o, card.index), 0);
                item.classList.toggle('has-missing', n > 0);
                let badge = item.querySelector('.missing-badge');
                if (n > 0 && !badge) {
                    badge = document.createElement('span');
                    badge.className = 'missing-badge';
                    item.querySelector('.card-name').appendChild(badge);
                }
                if (badge) n > 0 ? badge.textContent = n + ' missing' : badge.remove();
            });
        }

        // ---- Missing cards: right-click a card (Pull, Pack, or Set aside) ----------
        const menu = $('ff-menu');
        let menuCard = null, menuOrders = [];
        function openMenu(card, orders, x, y) {
            menuCard = card;
            menuOrders = orders;
            renderMenu();
            menu.hidden = false;
            const w = menu.offsetWidth, h = menu.offsetHeight;
            menu.style.left = Math.max(8, Math.min(x, window.innerWidth - w - 8)) + 'px';
            menu.style.top = Math.max(8, Math.min(y, window.innerHeight - h - 8)) + 'px';
        }
        function closeMenu() { menu.hidden = true; menuCard = null; }
        function orderName(o) {
            const rec = DATA.orders.find((r) => r.order_number === o) || {};
            const no = labelNo[o];
            return (no ? '#' + no + ' · ' : 'Set aside · ') + esc(rec.short_id || o)
                + (rec.first_name ? ' <span class="meta">' + esc(rec.first_name + ' ' + rec.last_name) + '</span>' : '');
        }
        function renderMenu() {
            const c = menuCard;
            menu.innerHTML = '<h3>' + esc(c.name) + '</h3><div class="meta">' + esc(c.set_name + ' #' + c.number + ' · ' + c.condition + (c.foil ? ' Foil' : ''))
                + '</div><div class="meta ff-menu-q">How many copies couldn\'t you find?</div>'
                + menuOrders.map((o) => {
                    const qty = c.orders[o], m = missingOf(o, c.index);
                    return '<div class="ff-menu-row"><span class="who">' + orderName(o) + '</span>'
                        + '<button data-miss="' + esc(o) + '" data-step="-1"' + (m <= 0 ? ' disabled' : '') + ' aria-label="One fewer missing">−</button>'
                        + '<span class="count' + (m ? ' some' : '') + '">' + m + ' of ' + qty + ' missing</span>'
                        + '<button data-miss="' + esc(o) + '" data-step="1"' + (m >= qty ? ' disabled' : '') + ' aria-label="One more missing">+</button></div>';
                }).join('')
                + '<div class="ff-menu-actions"><button class="nav-btn" data-miss-all="1">All missing</button>'
                + '<button class="nav-btn" data-miss-all="0">None missing</button><button class="nav-btn primary" data-miss-close>Done</button></div>';
        }
        function setMissing(o, index, n) {
            const key = o + '|' + index;
            if (n > 0) missing[key] = n; else delete missing[key];
            save();
            render();
            if (menuCard) renderMenu();
        }
        menu.addEventListener('click', (e) => {
            e.stopPropagation();
            const b = e.target.closest('button');
            if (!b || !menuCard) return;
            if (b.dataset.miss) {
                const o = b.dataset.miss;
                const n = Math.max(0, Math.min(menuCard.orders[o], missingOf(o, menuCard.index) + Number(b.dataset.step)));
                setMissing(o, menuCard.index, n);
            } else if (b.dataset.missAll) {
                const all = b.dataset.missAll === '1';
                menuOrders.forEach((o) => { const key = o + '|' + menuCard.index; if (all) missing[key] = menuCard.orders[o]; else delete missing[key]; });
                save();
                render();
                renderMenu();
            } else if ('missClose' in b.dataset) {
                closeMenu();
            }
        });
        document.addEventListener('contextmenu', (e) => {
            const row = e.target.closest('.card-item[data-index]');
            const line = e.target.closest('[data-card-line]');
            if (!row && !line) return;
            e.preventDefault();
            if (row) {
                const card = DATA.cards[+row.dataset.index];
                if (card) openMenu(card, Object.keys(card.orders), e.clientX, e.clientY);
            } else {
                const [o, idx] = line.dataset.cardLine.split('|');
                openMenu(DATA.cards[+idx], [o], e.clientX, e.clientY);
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
            settings = { ...DEFAULTS, custom: { ...DEFAULTS.custom } };
            syncInputs();
            onSettings();
        });

        function renderLabels() {
            const l = currentLayout();
            $('ff-custom').hidden = settings.layout !== 'custom';
            $('ff-layout-hint').textContent = l.per_sheet + ' per sheet, each ' + (+l.label_width.toFixed(2)) + ' x ' + (+l.label_height.toFixed(2)) + ' in';
            const visible = bulk.filter((o) => !(settings.hide_printed && printed[o.order_number]));
            $('ff-label-rows').innerHTML = visible.map((o) => {
                const on = !excluded.has(o.order_number);
                const chips = [printed[o.order_number] ? '<span class="ff-chip done">Printed</span>' : '<span class="ff-chip ok">Label</span>',
                    ...flags(o).map((f) => '<span class="ff-chip warn">' + esc(f) + '</span>')].join('');
                return '<tr class="' + (on ? '' : 'off') + '">'
                    + '<td><input type="checkbox" data-include="' + esc(o.order_number) + '"' + (on ? ' checked' : '') + ' aria-label="Print label for ' + esc(o.short_id) + '"></td>'
                    + '<td class="num"><span class="label-no-big">#' + labelNo[o.order_number] + '</span></td>'
                    + '<td>' + esc(o.short_id) + '<div class="sub">' + esc(o.order_date) + '</div></td>'
                    + '<td>' + esc(o.first_name) + ' ' + esc(o.last_name) + '<div class="sub">' + esc([o.city, o.state].filter(Boolean).join(', ')) + ' ' + esc(o.postal_code) + '</div></td>'
                    + '<td class="num">' + cardCount(o) + '</td><td class="num">' + money(o.product_value) + '</td>'
                    + '<td class="num">' + money(o.shipping_fee) + '</td><td>' + chips + '</td></tr>';
            }).join('');
            const empty = $('ff-label-empty');
            empty.hidden = visible.length > 0;
            empty.textContent = bulk.length ? 'Every bulk label is marked printed.' : 'No orders qualify for bulk with these settings.';
            const sel = selected();
            const sheets = Math.ceil(sel.length / (l.per_sheet || 1));
            $('ff-sheet-count').textContent = sel.length ? sel.length + ' label' + (sel.length === 1 ? '' : 's') + ' · ' + sheets + ' sheet' + (sheets === 1 ? '' : 's') : '';
            $('ff-mark-printed').disabled = sel.length === 0;
            if (!$('step-labels').hidden) schedulePreview();
        }
        document.addEventListener('change', (e) => {
            const id = e.target.dataset && e.target.dataset.include;
            if (!id) return;
            e.target.checked ? excluded.delete(id) : excluded.add(id);
            save();
            render();
        });
        $('ff-select-all').addEventListener('click', () => { excluded = new Set(); save(); render(); });
        $('ff-select-none').addEventListener('click', () => { excluded = new Set(bulk.map((o) => o.order_number)); save(); render(); });
        $('ff-mark-printed').addEventListener('click', () => {
            selected().forEach((o) => printed[o.order_number] = new Date().toISOString());
            save();
            render();
        });
        $('ff-clear-printed').addEventListener('click', () => {
            if (!confirm('Clear the printed mark from every order in this batch?')) return;
            printed = {};
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
            const shipments = sel.map((o) => ({ ...o, label_number: labelNo[o.order_number], item_count: cardCount(o) }));
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

        // ---- Pack + Set aside steps ------------------------------------------------
        function cardLine(o, x, clickable) {
            const c = x.card;
            const key = o.order_number + '|' + c.index;
            const m = missingOf(o.order_number, c.index);
            const gone = m >= x.qty;
            const bits = [c.set_name + ' #' + c.number, c.condition + (c.foil ? ' Foil' : '')];
            const note = m ? '<span class="missing-note">' + (gone ? 'missing: refund' : m + ' of ' + x.qty + ' missing') + '</span>' : '';
            clickable = clickable && !gone;
            return '<li class="env-card' + (gone ? ' gone static' : clickable ? (packed[key] ? ' in' : '') : ' static') + '"'
                + ' data-card-line="' + esc(key) + '"'
                + (clickable ? ' data-pack="' + esc(key) + '"' : '') + (c.image ? ' data-image="' + esc(c.image) + '"' : '') + '>'
                + '<span class="qty">' + (gone ? x.qty : x.qty - m) + 'x</span><span><span>' + esc(c.name) + '</span>' + (c.variant ? ' <span class="meta">(' + esc(c.variant) + ')</span>' : '') + note
                + '<div class="meta">' + esc(bits.join(' · ')) + '</div></span></li>';
        }
        // Fully missing lines don't hold an envelope back; partly missing ones still need packing.
        const lineDone = (o, x) => missingOf(o.order_number, x.card.index) >= x.qty || packed[o.order_number + '|' + x.card.index];
        function isPacked(o) {
            const lines = cardsByOrder[o.order_number] || [];
            return lines.length > 0 && lines.every((x) => lineDone(o, x));
        }

        function refundText(o, lines) {
            const who = o.has_shipping ? o.first_name + ' ' + o.last_name : 'Buyer';
            return 'Order ' + o.order_number + ' (' + who + '): missing\n'
                + lines.map((x) => '- ' + missingOf(o.order_number, x.card.index) + 'x ' + x.card.name
                    + (x.card.variant ? ' (' + x.card.variant + ')' : '') + ', ' + x.card.set_name + ' #' + x.card.number
                    + ', ' + x.card.condition + (x.card.foil ? ' Foil' : '')).join('\n');
        }
        function renderRefunds() {
            const refunds = DATA.orders.map((o) => ({ o, lines: (cardsByOrder[o.order_number] || []).filter((x) => missingOf(o.order_number, x.card.index)) }))
                .filter((r) => r.lines.length);
            const box = $('ff-refunds');
            box.hidden = refunds.length === 0;
            box.innerHTML = '<h3>Refunds to send (' + refunds.length + ' order' + (refunds.length === 1 ? '' : 's') + ')</h3>'
                + refunds.map(({ o, lines }) => '<div class="refund"><div class="who">'
                    + (labelNo[o.order_number] ? '<b>#' + labelNo[o.order_number] + '</b> · ' : '<span class="meta">Set aside</span> · ')
                    + '<b>' + esc(o.has_shipping ? o.first_name + ' ' + o.last_name : o.short_id) + '</b> <span class="meta">' + esc(o.order_number) + '</span>'
                    + '<ul>' + lines.map((x) => '<li>' + missingOf(o.order_number, x.card.index) + 'x ' + esc(x.card.name)
                        + (x.card.variant ? ' (' + esc(x.card.variant) + ')' : '') + ' <span class="meta">' + esc(x.card.set_name + ' #' + x.card.number + ' · ' + x.card.condition + (x.card.foil ? ' Foil' : '')) + '</span></li>').join('') + '</ul></div>'
                    + '<div><button class="nav-btn" data-copy-refund="' + esc(o.order_number) + '">Copy details</button></div></div>').join('');
            box.querySelectorAll('[data-copy-refund]').forEach((b) => b.addEventListener('click', async () => {
                const r = refunds.find((x) => x.o.order_number === b.dataset.copyRefund);
                try { await navigator.clipboard.writeText(refundText(r.o, r.lines)); b.textContent = 'Copied'; }
                catch (err) { b.textContent = 'Copy failed'; }
                setTimeout(() => { b.textContent = 'Copy details'; }, 1500);
            }));
        }
        function renderPack() {
            const done = bulk.filter(isPacked).length;
            $('ff-pack-progress').textContent = done + ' / ' + bulk.length + ' envelopes packed';
            const hide = $('ff-hide-packed').checked;
            $('ff-envelopes').innerHTML = bulk.filter((o) => !(hide && isPacked(o))).map((o) => {
                const lines = cardsByOrder[o.order_number] || [];
                const inCount = lines.filter((x) => lineDone(o, x)).length;
                const miss = missingTotal(o);
                const p = isPacked(o);
                return '<div class="envelope' + (p ? ' packed' : '') + '">'
                    + '<div class="env-head"><span class="env-no">#' + labelNo[o.order_number] + '</span>'
                    + '<span class="env-title"><b>' + esc(o.first_name) + ' ' + esc(o.last_name) + '</b> <span class="meta">' + esc(o.short_id) + ' · ' + cardCount(o) + 'x</span></span>'
                    + '<span class="env-status">' + (p ? '✓ Packed' : inCount + ' / ' + lines.length + ' in') + '</span></div>'
                    + (miss ? '<div class="env-reasons"><span class="ff-chip bad">' + miss + ' missing: refund the buyer</span></div>' : '')
                    + '<ul class="env-cards">' + lines.map((x) => cardLine(o, x, true)).join('') + '</ul></div>';
            }).join('') || '<p class="ff-empty">' + (bulk.length ? 'Every envelope is packed.' : 'No bulk orders with these settings.') + '</p>';

            $('ff-aside').innerHTML = aside.map((o) => {
                const lines = cardsByOrder[o.order_number] || [];
                const name = o.has_shipping ? esc(o.first_name) + ' ' + esc(o.last_name) : '<span class="meta">No shipping info</span>';
                return '<div class="envelope set-aside">'
                    + '<div class="env-head"><span class="env-title"><b>' + name + '</b> <span class="meta">' + esc(o.short_id) + '</span></span></div>'
                    + '<div class="env-reasons">' + o.reasons.map((r) => '<span class="ff-chip warn">' + esc(r) + '</span>').join('') + '</div>'
                    + (lines.length ? '<ul class="env-cards">' + lines.map((x) => cardLine(o, x, false)).join('') + '</ul>'
                        : '<div class="env-other">No Magic singles on the pull sheet</div>')
                    + '</div>';
            }).join('') || '<p class="ff-empty">Nothing to set aside.</p>';
        }
        document.addEventListener('click', (e) => {
            const line = e.target.closest('[data-pack]');
            if (!line) return;
            const key = line.dataset.pack;
            packed[key] ? delete packed[key] : packed[key] = true;
            save();
            render();
        });
        $('ff-hide-packed').addEventListener('change', renderPack);
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
            renderLabels();
            renderPack();
            renderRefunds();
            $('labels-count').textContent = '(' + bulk.length + ')';
            $('pack-count').textContent = '(' + bulk.filter(isPacked).length + '/' + bulk.length + ')';
            $('aside-count').textContent = '(' + aside.length + ')';
            $('bulk-banner').textContent = DATA.orders.length + ' orders: ' + bulk.length + ' bulk (labels ' + (bulk.length ? '#1–#' + bulk.length : 'none')
                + '), ' + aside.length + ' set aside';
        }

        syncInputs();
        render();
        const saved = load(KEY + 'step', 'pull');
        showStep(document.getElementById('step-' + saved) ? saved : 'pull');
    })();
"""
