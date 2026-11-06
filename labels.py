"""Lay out address labels for bulk (untracked) orders onto printer sheets and render a PDF."""

import io
from dataclasses import dataclass

from reportlab.lib.units import inch
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.pdfgen import canvas

FONT = "Helvetica"
FONT_BOLD = "Helvetica-Bold"


@dataclass
class Layout:
    """A printer sheet split into a grid of labels.

    page_width/page_height are the sheet as the printer feeds it. With rotate,
    the grid is laid across the sheet sideways (text runs along the long edge).
    """
    key: str
    name: str
    page_width: float
    page_height: float
    cols: int
    rows: int
    rotate: bool = False
    margin: float = 0.1

    @property
    def grid_width(self) -> float:
        return self.page_height if self.rotate else self.page_width

    @property
    def grid_height(self) -> float:
        return self.page_width if self.rotate else self.page_height

    @property
    def per_sheet(self) -> int:
        return self.cols * self.rows

    @property
    def label_width(self) -> float:
        return self.grid_width / self.cols

    @property
    def label_height(self) -> float:
        return self.grid_height / self.rows

    def describe(self) -> dict:
        return {
            "key": self.key, "name": self.name,
            "page_width": self.page_width, "page_height": self.page_height,
            "cols": self.cols, "rows": self.rows, "rotate": self.rotate, "margin": self.margin,
            "per_sheet": self.per_sheet,
            "label_width": round(self.label_width, 3), "label_height": round(self.label_height, 3),
        }


PRESETS = [
    Layout("4x6-4", "4x6 sheet, 4 labels (3 x 2 in), cut", 4, 6, 2, 2, rotate=True),
    Layout("4x6-2", "4x6 sheet, 2 labels (4 x 3 in), cut", 4, 6, 1, 2),
    Layout("4x6-6", "4x6 sheet, 6 labels (2 x 2 in), cut", 4, 6, 3, 2, rotate=True),
    Layout("4x6-1", "4x6 sheet, 1 label (6 x 4 in)", 4, 6, 1, 1, rotate=True, margin=0.25),
    Layout("roll-2.25x1.25", "Address roll, 2.25 x 1.25 in", 2.25, 1.25, 1, 1, margin=0.06),
    Layout("roll-3.5x1.125", "Address roll, 3.5 x 1.125 in", 3.5, 1.125, 1, 1, margin=0.06),
]
PRESETS_BY_KEY = {p.key: p for p in PRESETS}
DEFAULT_LAYOUT = "4x6-4"


def layout_from_settings(settings: dict) -> Layout:
    key = settings.get("layout") or DEFAULT_LAYOUT
    if key != "custom":
        if key not in PRESETS_BY_KEY:
            raise ValueError(f"Unknown layout '{key}'")
        base = PRESETS_BY_KEY[key]
        margin = settings.get("margin")
        if margin is None:
            return base
        return Layout(base.key, base.name, base.page_width, base.page_height,
                      base.cols, base.rows, base.rotate, _bounded(margin, 0, 1, "margin"))
    c = settings.get("custom") or {}
    return Layout(
        "custom", "Custom",
        _bounded(c.get("page_width", 4), 0.5, 12, "sheet width"),
        _bounded(c.get("page_height", 6), 0.5, 12, "sheet height"),
        int(_bounded(c.get("cols", 1), 1, 10, "columns")),
        int(_bounded(c.get("rows", 1), 1, 10, "rows")),
        bool(c.get("rotate", False)),
        _bounded(settings.get("margin", 0.1) if settings.get("margin") is not None else 0.1, 0, 1, "margin"),
    )


def _bounded(value, low: float, high: float, label: str) -> float:
    try:
        value = float(value)
    except (TypeError, ValueError):
        raise ValueError(f"{label} must be a number")
    if not low <= value <= high:
        raise ValueError(f"{label} must be between {low} and {high}")
    return value


def address_lines(shipment: dict, uppercase: bool) -> list[str]:
    name = f"{shipment.get('first_name', '')} {shipment.get('last_name', '')}".strip()
    city_line = f"{shipment.get('city', '')}, {shipment.get('state', '')} {shipment.get('postal_code', '')}".strip(" ,")
    lines = [name, shipment.get("address1", ""), shipment.get("address2", ""), city_line]
    country = (shipment.get("country") or "US").strip()
    if country.upper() not in ("US", "USA", "UNITED STATES"):
        lines.append(country)
    lines = [l for l in (x.strip() for x in lines) if l]
    return [l.upper() for l in lines] if uppercase else lines


def card_total(shipment: dict) -> int:
    if shipment.get("item_count"):
        return shipment["item_count"]
    return sum(c.get("quantity", 1) for c in shipment.get("cards") or [])


def id_line(shipment: dict, mode: str) -> str:
    """The small matching line at the bottom of a label, e.g. "#4 · F1175B · 6x"."""
    if mode == "none":
        return ""
    short_id = shipment.get("short_id") or shipment.get("order_id", "")
    if mode == "id":
        return short_id
    # A bare count ("6x") so the outside of the envelope doesn't say it holds cards.
    text = f"{short_id} · {card_total(shipment)}x"
    number = shipment.get("label_number")
    if mode == "num_id_count" and number:
        text = f"#{number} · {text}"
    return text


def _fit_size(lines: list[str], font: str, max_width: float, max_height: float,
              largest: float, smallest: float = 4.0, leading: float = 1.18) -> float:
    size = largest
    while size > smallest:
        widest = max((stringWidth(l, font, size) for l in lines), default=0)
        if widest <= max_width and len(lines) * size * leading <= max_height:
            return size
        size -= 0.25
    return smallest


def _clip(text: str, font: str, size: float, max_width: float) -> str:
    if stringWidth(text, font, size) <= max_width:
        return text
    while text and stringWidth(text + "…", font, size) > max_width:
        text = text[:-1]
    return text + "…"


def _draw_label(c: canvas.Canvas, shipment: dict, x: float, y: float, w: float, h: float,
                margin: float, settings: dict) -> None:
    """Draw one label with its lower-left corner at (x, y); all units are points."""
    uppercase = settings.get("uppercase", True)
    ret_lines = [l.strip() for l in (settings.get("return_address") or "").splitlines() if l.strip()]
    if uppercase:
        ret_lines = [l.upper() for l in ret_lines]
    tag = id_line(shipment, settings.get("id_line", "id_count"))

    left, right, bottom, top = x + margin, x + w - margin, y + margin, y + h - margin
    inner_w, inner_h = right - left, top - bottom

    if tag:
        tag_size = max(4.5, min(8.0, inner_h * 0.075))
        c.setFont(FONT, tag_size)
        c.drawRightString(right, bottom, _clip(tag, FONT, tag_size, inner_w))
        bottom += tag_size * 1.5

    if ret_lines:
        ret_size = _fit_size(ret_lines, FONT, inner_w * 0.6, inner_h * 0.3, largest=7.5)
        c.setFont(FONT, ret_size)
        ty = top - ret_size
        for line in ret_lines:
            c.drawString(left, ty, line)
            ty -= ret_size * 1.18
        top = ty - ret_size * 0.6

    lines = address_lines(shipment, uppercase)
    box_h = top - bottom
    size = _fit_size(lines, FONT_BOLD, inner_w, box_h, largest=min(14.0, inner_h * 0.16))
    leading = size * 1.18
    block_h = len(lines) * leading
    block_w = max(stringWidth(l, FONT_BOLD, size) for l in lines)
    # Center the address block in the space left over, starting at least at the left margin.
    bx = left + max(0.0, (inner_w - block_w) / 2)
    by = bottom + (box_h - block_h) / 2 + block_h - size
    c.setFont(FONT_BOLD, size)
    for line in lines:
        c.drawString(bx, by, _clip(line, FONT_BOLD, size, inner_w))
        by -= leading


def render_pdf(shipments: list[dict], settings: dict) -> bytes:
    layout = layout_from_settings(settings)
    page_w, page_h = layout.page_width * inch, layout.page_height * inch
    grid_w, grid_h = layout.grid_width * inch, layout.grid_height * inch
    cell_w, cell_h = layout.label_width * inch, layout.label_height * inch
    margin = layout.margin * inch
    cut_lines = settings.get("cut_lines", True) and layout.per_sheet > 1

    if not shipments:
        raise ValueError("No labels to print")
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=(page_w, page_h))
    c.setTitle("Shipping labels")
    per = layout.per_sheet
    for start in range(0, len(shipments), per):
        if layout.rotate:
            # Turn the grid sideways: grid x runs up the sheet, grid y runs right to left.
            c.translate(page_w, 0)
            c.rotate(90)
        for slot, shipment in enumerate(shipments[start:start + per]):
            col, row = slot % layout.cols, slot // layout.cols
            x = col * cell_w
            y = grid_h - (row + 1) * cell_h
            _draw_label(c, shipment, x, y, cell_w, cell_h, margin, settings)
        if cut_lines:
            c.setLineWidth(0.4)
            c.setDash(2, 3)
            for col in range(1, layout.cols):
                c.line(col * cell_w, 0, col * cell_w, grid_h)
            for row in range(1, layout.rows):
                c.line(0, row * cell_h, grid_w, row * cell_h)
            c.setDash()
        c.showPage()
    c.save()
    return buf.getvalue()
