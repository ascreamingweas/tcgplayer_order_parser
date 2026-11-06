import io
import json
import re

import pytest
from fastapi.testclient import TestClient
from pypdf import PdfReader

import app as app_module
import fulfillment
import mtg_packing_slip_organizer as organizer

SHIPPING = """Order #,FirstName,LastName,Address1,Address2,City,State,PostalCode,Country,Order Date,Product Weight,Shipping Method,Item Count,Value Of Products,Shipping Fee Paid,Tracking #,Carrier
"AAAA0000-111111-0000A","Ada","Lovelace","1 Main St","Apt 2","Springfield","IL","62701","US","2026-11-05","0.14","Standard (7-10 days)","3","5.96","1.49","",""
"AAAA0000-222222-0000B","Alan","Turing","2 Oak Ave","","Portland","OR","97201-1234","US","2026-11-05","0.07","Standard (7-10 days)","2","8.00","1.49","",""
"AAAA0000-333333-0000C","Grace","Hopper","3 Pine Rd","","Austin","TX","73301","US","2026-11-05","0.07","Standard (7-10 days)","1","2.00","1.49","",""
"AAAA0000-444444-0000D","Edsger","Dijkstra","4 Elm St","","Boston","MA","02101","US","2026-11-05","0.07","Standard (7-10 days)","1","3.00","1.49","",""
"""

PULL = """Product Line,Product Name,Condition,Number,Set,Rarity,Quantity,Main Photo URL,Set Release Date,SkuId,Order Quantity
"Magic","Shared Card","Near Mint","1","Test Set","R","2","","10/02/2026 00:00:00","1","AAAA0000-333333-0000C:1 | AAAA0000-111111-0000A:1"
"Magic","Other Card (Borderless)","Near Mint Foil","2","Test Set","U","2","","10/02/2026 00:00:00","2","AAAA0000-111111-0000A:2"
"Magic","Mixed Card","Near Mint","5","Test Set","C","1","","10/02/2026 00:00:00","5","AAAA0000-222222-0000B:1"
"Lorcana TCG","Lore Card","Near Mint","3/204","Lore Set","Rare","1","","10/02/2026 00:00:00","3","AAAA0000-222222-0000B:1"
"Magic","Test Set Play Booster","Unopened","","Test Set","","1","","10/02/2026 00:00:00","4","AAAA0000-444444-0000D:1"
Orders Contained in Pull Sheet:,AAAA0000-111111-0000A|AAAA0000-222222-0000B|AAAA0000-333333-0000C|AAAA0000-444444-0000D
"""


@pytest.fixture(autouse=True)
def no_scryfall(monkeypatch):
    monkeypatch.setattr(app_module, "fetch_colors_from_scryfall", lambda cards, on_progress=None: cards)


client = TestClient(app_module.app)


def test_csv_kind():
    assert fulfillment.csv_kind(PULL) == "pull_sheet"
    assert fulfillment.csv_kind("﻿" + SHIPPING) == "shipping_export"
    assert fulfillment.csv_kind("a,b\n") == "unknown"


def test_sealed_magic_is_not_a_card_and_marks_the_order():
    cards, orders = organizer.parse_pull_sheet(PULL)
    assert "Test Set Play Booster" not in [c.card_name for c in cards]
    dijkstra = next(o for o in orders if o["short_id"] == "444444")
    assert dijkstra["tcgs"] == {"Magic Sealed": 1}
    assert dijkstra["magic_cards"] == 0


def test_build_orders_joins_both_files():
    _, summaries = organizer.parse_pull_sheet(PULL)
    shipments = fulfillment.parse_shipping_export(SHIPPING)
    extra = fulfillment.Shipment(**{**shipments[0].__dict__, "order_number": "AAAA0000-999999-0000Z"})
    records = fulfillment.build_orders(summaries, shipments + [extra])
    assert [r["short_id"] for r in records] == ["111111", "222222", "333333", "444444", "999999"]
    ada = records[0]
    assert ada["has_shipping"] and ada["in_pull_sheet"] and ada["city"] == "Springfield"
    assert records[1]["tcgs"] == {"Magic": 1, "Lorcana TCG": 1}
    assert records[-1]["in_pull_sheet"] is False


def _run_job(files):
    res = client.post("/api/parse", files=[("files", (n, b.encode(), "text/csv")) for n, b in files])
    assert res.status_code == 200, res.text
    job_id = res.json()["job_id"]
    stream = client.get(f"/api/parse/{job_id}/progress").text
    url = re.search(r'"report_url": "([^"]+)"', stream).group(1)
    return client.get(url).text


def test_report_with_shipping_export_has_bulk_steps_and_data():
    html = _run_job([("ship.csv", SHIPPING), ("pull.csv", PULL)])
    for marker in ['id="steps"', 'id="step-pull"', 'id="step-labels"', 'id="step-pack"', 'id="step-aside"']:
        assert marker in html
    data = json.loads(re.search(r'<script type="application/json" id="fulfillment-data">(.*?)</script>', html, re.S).group(1))
    assert len(data["orders"]) == 4
    # Each card record's index points at the row with that data-index.
    for card in data["cards"]:
        row = re.search(rf'data-index="{card["index"]}"[^>]*>.*?class="card-name">(.*?)</div>', html, re.S).group(1)
        assert card["name"] in row
    assert {l["key"] for l in data["layouts"]} >= {"4x6-4", "roll-2.25x1.25"}


def test_report_from_pull_sheet_alone_has_no_bulk_steps():
    html = _run_job([("pull.csv", PULL)])
    assert 'id="steps"' not in html and "fulfillment-data" not in html


def test_parse_rejects_bad_csv_combinations():
    two_pulls = client.post("/api/parse", files=[("files", ("a.csv", PULL.encode())), ("files", ("b.csv", PULL.encode()))])
    assert two_pulls.status_code == 400
    ship_only = client.post("/api/parse", files=[("files", ("s.csv", SHIPPING.encode()))])
    assert ship_only.status_code == 400 and "pull sheet" in ship_only.json()["detail"]


def test_labels_pdf_prints_label_numbers():
    shipments = [{**s.__dict__, "short_id": s.order_number.split("-")[1], "label_number": i + 1}
                 for i, s in enumerate(fulfillment.parse_shipping_export(SHIPPING))]
    res = client.post("/api/labels.pdf", json={"shipments": shipments, "settings": {"id_line": "num_id_count"}})
    assert res.status_code == 200
    pdf = PdfReader(io.BytesIO(res.content))
    assert len(pdf.pages) == 1
    text = pdf.pages[0].extract_text()
    assert "ADA LOVELACE" in text and "#1 · 111111 · 3x" in text
    assert client.post("/api/labels.pdf", json={"shipments": [], "settings": {}}).status_code == 400
