import gzip
import json

import pytest

import mtg_packing_slip_organizer as organizer
import scryfall_db

CARDS = [
    {"name": "Firebending Student", "set": "tla", "set_name": "Avatar: The Last Airbender", "collector_number": "139",
     "colors": ["R"], "cmc": 2.0, "type_line": "Creature — Human Monk", "rarity": "rare", "released_at": "2025-11-21",
     "image_uris": {"normal": "https://img/139.jpg", "small": "x"}, "oracle_text": "dropped"},
    {"name": "Firebending Student", "set": "tla", "set_name": "Avatar: The Last Airbender", "collector_number": "393",
     "colors": ["R"], "cmc": 2.0, "type_line": "Creature — Human Monk", "rarity": "rare", "released_at": "2025-11-21",
     "frame_effects": ["extendedart"], "promo_types": ["buyabox"], "image_uris": {"normal": "https://img/393.jpg"}},
    {"name": "Glamdring, Foe-hammer // Gleam of Death", "set": "ltr", "collector_number": "204", "cmc": 2.0,
     "released_at": "2023-06-23", "card_faces": [{"name": "Glamdring, Foe-hammer", "type_line": "Artifact — Equipment",
                                                  "colors": [], "image_uris": {"normal": "https://img/g.jpg"}}]},
    {"name": "Ruin Crab", "set": "znr", "collector_number": "75", "colors": ["U"], "released_at": "2020-09-25"},
    {"name": "Ruin Crab", "set": "ytest", "collector_number": "1", "colors": ["U"], "released_at": "2030-01-01", "digital": True},
]


@pytest.fixture
def db(tmp_path, monkeypatch):
    path = tmp_path / "scryfall.sqlite"
    monkeypatch.setattr(scryfall_db, "DB_PATH", path)
    bulk = tmp_path / "cards.jsonl.gz"
    with gzip.open(bulk, "wt") as f:
        f.write("\n".join(json.dumps(c) for c in CARDS) + "\n")
    assert scryfall_db.build(bulk, "2026-11-08T00:00:00Z") == len(CARDS)
    return bulk


def test_lookups(db):
    assert scryfall_db.by_set_number("TLA", "393")["image_uris"] == {"normal": "https://img/393.jpg"}
    assert scryfall_db.by_name_number("Firebending Student", "393")["collector_number"] == "393"
    assert "oracle_text" not in scryfall_db.by_set_number("tla", "139")
    # Front face name finds the full split/adventure card.
    assert scryfall_db.by_name("glamdring, foe-hammer")["card_faces"][0]["type_line"] == "Artifact — Equipment"
    # Name-only: paper printings win over a newer digital one.
    assert scryfall_db.by_name("Ruin Crab")["set"] == "znr"
    assert scryfall_db.by_set_number("tla", "999") is None


def test_rebuild_keeps_recorded_misses(db):
    scryfall_db.record_miss("Brand New Card", "Future Set", "1", found_live=True)
    scryfall_db.build(db, "2026-11-09T00:00:00Z")
    status = scryfall_db.status()
    assert status["scryfall_updated_at"] == "2026-11-09T00:00:00Z"
    assert [m["card_name"] for m in status["recent_misses"]] == ["Brand New Card"]


def test_search_uses_local_copy_without_calling_scryfall(db, monkeypatch):
    def no_network(*a, **k):
        raise AssertionError("live Scryfall called")

    monkeypatch.setattr(organizer, "scryfall_get", no_network)
    monkeypatch.setattr(organizer, "get_lookup_set_code", lambda set_name, card_name: None)
    card = organizer.search_scryfall("Firebending Student", "Buy-A-Box Promos", "393")
    assert card["image_uris"]["normal"] == "https://img/393.jpg"
    assert organizer.get_card_color(card) == "Red"


def test_local_miss_falls_back_to_live_and_is_recorded(db, monkeypatch):
    live = {"name": "Brand New Card", "set": "new", "collector_number": "7", "colors": ["G"]}
    monkeypatch.setattr(organizer, "get_lookup_set_code", lambda set_name, card_name: "new")
    monkeypatch.setattr(organizer, "search_scryfall_live", lambda *a: live)
    assert organizer.search_scryfall("Brand New Card", "New Set", "7") == live
    miss = scryfall_db.status()["recent_misses"][0]
    assert (miss["card_name"], miss["found_live"]) == ("Brand New Card", 1)


def test_set_and_name_match_when_numbers_are_written_differently(tmp_path, monkeypatch):
    path = tmp_path / "s.sqlite"
    monkeypatch.setattr(scryfall_db, "DB_PATH", path)
    bulk = tmp_path / "b.jsonl.gz"
    rows = [{"name": "Wizard Class", "set": "plst", "collector_number": "AFR-81", "colors": ["U"]},
            {"name": "Wizard Class", "set": "plst", "collector_number": "BLC-112", "colors": ["U"]}]
    with gzip.open(bulk, "wt") as f:
        f.write("\n".join(json.dumps(r) for r in rows))
    scryfall_db.build(bulk, "x")
    assert scryfall_db.by_set_number("plst", "081/281") is None
    assert scryfall_db.by_set_name("plst", "Wizard Class", "081/281")["collector_number"] == "AFR-81"
    assert scryfall_db.by_set_name("plst", "Wizard Class", "112/300")["collector_number"] == "BLC-112"


def test_name_number_prefers_the_promo_kind_in_the_set_name(tmp_path, monkeypatch):
    monkeypatch.setattr(scryfall_db, "DB_PATH", tmp_path / "s.sqlite")
    bulk = tmp_path / "b.jsonl.gz"
    rows = [
        {"name": "Jadar", "set": "dbl", "collector_number": "108", "released_at": "2022-01-28"},
        {"name": "Jadar", "set": "mid", "collector_number": "108", "released_at": "2021-09-24"},
        {"name": "Jadar", "set": "pmid", "collector_number": "108p", "released_at": "2021-09-24", "promo_types": ["promopack"]},
        {"name": "Jadar", "set": "pmid", "collector_number": "108s", "released_at": "2021-09-24", "promo_types": ["prerelease"]},
    ]
    with gzip.open(bulk, "wt") as f:
        f.write("\n".join(json.dumps(r) for r in rows))
    scryfall_db.build(bulk, "x")
    pick = lambda set_name: scryfall_db.by_name_number("Jadar", "108", set_name)["collector_number"]
    assert pick("Promo Pack: Innistrad: Midnight Hunt") == "108p"
    assert pick("Prerelease Cards") == "108s"
    assert pick("Some Unknown Set") == "108"
    assert scryfall_db.by_name_number("Jadar", "999") is None
