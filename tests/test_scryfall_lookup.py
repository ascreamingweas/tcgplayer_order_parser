import io
import json
import urllib.error

import pytest

import mtg_packing_slip_organizer as organizer


class FakeResponse(io.BytesIO):
    def __enter__(self):
        return self

    def __exit__(self, *exc):
        return False


def http_error(code):
    return urllib.error.HTTPError("https://api.scryfall.com/x", code, "err", {}, None)


@pytest.fixture(autouse=True)
def no_sleep(monkeypatch):
    monkeypatch.setattr(organizer.time, "sleep", lambda s: None)


def test_scryfall_get_retries_when_rate_limited(monkeypatch):
    calls = []

    def fake_urlopen(req, timeout):
        calls.append(req.full_url)
        if len(calls) < 3:
            raise http_error(429)
        return FakeResponse(json.dumps({"name": "Firebending Student"}).encode())

    monkeypatch.setattr(organizer.urllib.request, "urlopen", fake_urlopen)
    assert organizer.scryfall_get("https://api.scryfall.com/x") == {"name": "Firebending Student"}
    assert len(calls) == 3


def test_scryfall_get_does_not_retry_not_found(monkeypatch):
    calls = []

    def fake_urlopen(req, timeout):
        calls.append(1)
        raise http_error(404)

    monkeypatch.setattr(organizer.urllib.request, "urlopen", fake_urlopen)
    with pytest.raises(urllib.error.HTTPError):
        organizer.scryfall_get("https://api.scryfall.com/x")
    assert len(calls) == 1


def test_unmapped_set_looks_up_name_and_collector_number(monkeypatch):
    seen = []
    printing = {"name": "Firebending Student", "set": "tla", "collector_number": "393", "colors": ["R"]}

    def fake_get(url, timeout=10):
        seen.append(url)
        return {"data": [printing]}

    monkeypatch.setattr(organizer, "get_lookup_set_code", lambda set_name, card_name: None)
    monkeypatch.setattr(organizer, "scryfall_get", fake_get)
    assert organizer.search_scryfall("Firebending Student", "Buy-A-Box Promos", "393") == printing
    assert len(seen) == 1 and "cn%3A%22393%22" in seen[0] and "Firebending+Student" in seen[0]
