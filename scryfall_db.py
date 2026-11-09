"""Local copy of Scryfall's card data, refreshed from its daily bulk download.

Scryfall publishes every English printing as one gzipped JSON-lines file
("Default Cards", ~80 MB) and asks apps that look up many cards to use it instead
of calling the API card by card. We keep the fields the organizer reads in a
SQLite file and look cards up there first; the live API stays as the fallback
for anything newer than the last download (see search_scryfall).

Lookups the local copy couldn't answer are recorded in the `misses` table so they
can be reviewed (status()) and, later, checked by a poller.
"""

import gzip
import json
import os
import sqlite3
import threading
import time
import urllib.request
from datetime import datetime, timezone
from pathlib import Path
from typing import Optional

BULK_INFO_URL = "https://api.scryfall.com/bulk-data/default-cards"
HEADERS = {"User-Agent": "MTG-Packing-Slip-Organizer/1.0", "Accept": "application/json"}
DB_PATH = Path(os.environ.get("SCRYFALL_DB", Path(__file__).parent / "data" / "scryfall.sqlite"))
REFRESH_EVERY = 24 * 60 * 60  # Scryfall regenerates the file about every 12 hours

# Card fields the organizer reads (color, mana value, type, rarity, image, name,
# exact printing). Everything else in the bulk file is dropped to keep the DB small.
CARD_FIELDS = ("name", "set", "set_name", "collector_number", "colors", "cmc", "type_line", "rarity",
               "layout", "promo_types", "frame_effects", "released_at", "lang", "digital")
FACE_FIELDS = ("name", "type_line", "colors")

_lock = threading.Lock()
_refreshing = threading.Event()


def _connect(path: Path = None) -> sqlite3.Connection:
    conn = sqlite3.connect(path or DB_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn


def _trim(card: dict) -> dict:
    out = {k: card[k] for k in CARD_FIELDS if k in card}
    if card.get("image_uris", {}).get("normal"):
        out["image_uris"] = {"normal": card["image_uris"]["normal"]}
    if card.get("card_faces"):
        faces = []
        for face in card["card_faces"]:
            f = {k: face[k] for k in FACE_FIELDS if k in face}
            if face.get("image_uris", {}).get("normal"):
                f["image_uris"] = {"normal": face["image_uris"]["normal"]}
            faces.append(f)
        out["card_faces"] = faces
    return out


def _create_schema(conn: sqlite3.Connection) -> None:
    conn.executescript("""
        CREATE TABLE IF NOT EXISTS meta (key TEXT PRIMARY KEY, value TEXT);
        CREATE TABLE IF NOT EXISTS cards (
            set_code TEXT NOT NULL,
            collector_number TEXT NOT NULL,
            name_key TEXT NOT NULL,       -- lowercased full name ("a // b" for split cards)
            front_key TEXT NOT NULL,      -- lowercased front face name
            released_at TEXT,
            paper INTEGER NOT NULL,
            data TEXT NOT NULL            -- trimmed Scryfall card JSON
        );
        CREATE TABLE IF NOT EXISTS misses (
            card_name TEXT NOT NULL,
            set_name TEXT,
            collector_number TEXT,
            found_live INTEGER NOT NULL,
            seen_at TEXT NOT NULL,
            PRIMARY KEY (card_name, set_name, collector_number)
        );
    """)


def _index(conn: sqlite3.Connection) -> None:
    conn.executescript("""
        CREATE INDEX IF NOT EXISTS cards_set_number ON cards (set_code, collector_number);
        CREATE INDEX IF NOT EXISTS cards_name ON cards (name_key);
        CREATE INDEX IF NOT EXISTS cards_front ON cards (front_key);
    """)


def build(jsonl_gz: Path, updated_at: str, path: Path = None) -> int:
    """Build a fresh DB from a downloaded bulk file. Returns the number of cards.

    Writes to a temporary file and swaps it in, so lookups keep working (on the old
    copy) while a refresh runs. Recorded misses are carried over.
    """
    path = Path(path or DB_PATH)
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(".building")
    tmp.unlink(missing_ok=True)
    conn = _connect(tmp)
    _create_schema(conn)
    count = 0
    with gzip.open(jsonl_gz, "rt", encoding="utf-8") as f:
        batch = []
        for line in f:
            line = line.strip().rstrip(",")
            if not line or line in ("[", "]"):
                continue
            card = json.loads(line)
            name = card.get("name", "")
            batch.append((
                card.get("set", ""), str(card.get("collector_number", "")), name.lower(),
                name.split(" // ")[0].lower(), card.get("released_at"),
                0 if card.get("digital") else 1, json.dumps(_trim(card), separators=(",", ":")),
            ))
            if len(batch) >= 5000:
                conn.executemany("INSERT INTO cards VALUES (?,?,?,?,?,?,?)", batch)
                count += len(batch)
                batch = []
        conn.executemany("INSERT INTO cards VALUES (?,?,?,?,?,?,?)", batch)
        count += len(batch)
    _index(conn)
    conn.executemany("INSERT OR REPLACE INTO meta VALUES (?, ?)", [
        ("updated_at", updated_at), ("built_at", datetime.now(timezone.utc).isoformat()), ("card_count", str(count)),
    ])
    if path.exists():
        conn.execute("ATTACH DATABASE ? AS old", (str(path),))
        try:
            conn.execute("INSERT OR IGNORE INTO misses SELECT * FROM old.misses")
        except sqlite3.Error:
            pass
        conn.commit()
        conn.execute("DETACH DATABASE old")
    conn.commit()
    conn.close()
    with _lock:
        os.replace(tmp, path)
    return count


def _meta(key: str) -> Optional[str]:
    if not DB_PATH.exists():
        return None
    try:
        with _lock:
            conn = _connect()
            row = conn.execute("SELECT value FROM meta WHERE key = ?", (key,)).fetchone()
            conn.close()
        return row["value"] if row else None
    except sqlite3.Error:
        return None


def needs_refresh() -> bool:
    built = _meta("built_at")
    if not built:
        return True
    age = datetime.now(timezone.utc) - datetime.fromisoformat(built)
    return age.total_seconds() > REFRESH_EVERY


def refresh(force: bool = False, log=print) -> bool:
    """Download Scryfall's latest Default Cards file and rebuild the DB if it changed.

    Returns True when the DB was rebuilt.
    """
    if _refreshing.is_set():
        return False
    _refreshing.set()
    try:
        req = urllib.request.Request(BULK_INFO_URL, headers=HEADERS)
        with urllib.request.urlopen(req, timeout=30) as response:
            info = json.loads(response.read().decode())
        updated_at = info["updated_at"]
        if not force and _meta("updated_at") == updated_at:
            # Same file as last time; just note that we checked.
            _set_meta("built_at", datetime.now(timezone.utc).isoformat())
            log(f"Scryfall data already current ({updated_at})")
            return False
        url = info.get("jsonl_download_uri") or info["download_uri"]
        log(f"Downloading Scryfall card data ({info.get('compressed_size', 0) // 1_000_000} MB)...")
        DB_PATH.parent.mkdir(parents=True, exist_ok=True)
        download = DB_PATH.with_suffix(".download.gz")
        req = urllib.request.Request(url, headers={"User-Agent": HEADERS["User-Agent"]})
        with urllib.request.urlopen(req, timeout=120) as response, open(download, "wb") as out:
            while chunk := response.read(1 << 20):
                out.write(chunk)
        if not url.endswith(".gz"):
            raise ValueError(f"Expected a gzipped bulk file, got {url}")
        start = time.time()
        count = build(download, updated_at)
        download.unlink(missing_ok=True)
        log(f"Scryfall data ready: {count} printings ({updated_at}, built in {time.time() - start:.0f}s)")
        return True
    finally:
        _refreshing.clear()


def _set_meta(key: str, value: str) -> None:
    with _lock:
        conn = _connect()
        conn.execute("INSERT OR REPLACE INTO meta VALUES (?, ?)", (key, value))
        conn.commit()
        conn.close()


def start_background_refresh(log=print) -> threading.Thread:
    """Refresh now if the copy is missing or stale, then keep checking daily."""
    def loop():
        while True:
            try:
                if needs_refresh():
                    refresh(log=log)
            except Exception as e:
                log(f"Scryfall data refresh failed (live lookups still work): {e}")
            time.sleep(60 * 60)

    thread = threading.Thread(target=loop, daemon=True, name="scryfall-refresh")
    thread.start()
    return thread


def available() -> bool:
    return DB_PATH.exists()


def _query(sql: str, params: tuple) -> Optional[dict]:
    if not DB_PATH.exists():
        return None
    try:
        with _lock:
            conn = _connect()
            row = conn.execute(sql, params).fetchone()
            conn.close()
    except sqlite3.Error:
        return None
    return json.loads(row["data"]) if row else None


# Prefer paper printings, then the newest, like Scryfall's own name lookup.
_ORDER = "ORDER BY paper DESC, released_at DESC"


def by_set_number(set_code: str, collector_number: str) -> Optional[dict]:
    return _query(f"SELECT data FROM cards WHERE set_code = ? AND collector_number = ? {_ORDER} LIMIT 1",
                  (set_code.lower(), str(collector_number)))


def _number_key(collector_number: str) -> str:
    """Comparable collector number: TCGplayer's "081/281" and Scryfall's "AFR-81" -> "81"."""
    return str(collector_number).split("/")[0].split("-")[-1].lstrip("0").lower()


def by_set_name(set_code: str, card_name: str, collector_number: str = "") -> Optional[dict]:
    """A printing of this card in this set, for when the collector numbers are written
    differently (The List: TCGplayer "081/281", Scryfall "AFR-81")."""
    if not DB_PATH.exists():
        return None
    key = card_name.split(" // ")[0].strip().lower()
    try:
        with _lock:
            conn = _connect()
            rows = conn.execute(
                f"SELECT collector_number, data FROM cards WHERE set_code = ? AND (name_key = ? OR front_key = ?) {_ORDER}",
                (set_code.lower(), card_name.strip().lower(), key)).fetchall()
            conn.close()
    except sqlite3.Error:
        return None
    if not rows:
        return None
    wanted = _number_key(collector_number) if collector_number else None
    best = next((r for r in rows if wanted and _number_key(r["collector_number"]) == wanted), rows[0])
    return json.loads(best["data"])


# TCGplayer catch-all promo sets -> the Scryfall promo type that marks those printings.
PROMO_HINTS = (("promo pack", "promopack"), ("prerelease", "prerelease"), ("buy-a-box", "buyabox"),
               ("bundle", "bundle"), ("store championship", "storechampionship"))


def by_name_number(card_name: str, collector_number: str, set_name: str = "") -> Optional[dict]:
    """A printing by card name and collector number, for TCGplayer sets Scryfall doesn't
    have. "Promo Pack: Innistrad: Midnight Hunt" #108 is Scryfall's pmid #108p, so
    numbers match loosely and the set name's promo kind breaks ties."""
    if not DB_PATH.exists():
        return None
    key = card_name.split(" // ")[0].strip().lower()
    try:
        with _lock:
            conn = _connect()
            rows = conn.execute(
                f"SELECT collector_number, data FROM cards WHERE name_key = ? OR front_key = ? {_ORDER}",
                (card_name.strip().lower(), key)).fetchall()
            conn.close()
    except sqlite3.Error:
        return None
    wanted = str(collector_number).lower()
    wanted_key = _number_key(collector_number)
    hint = next((promo for words, promo in PROMO_HINTS if words in (set_name or "").lower()), None)

    def rank(row):
        data = json.loads(row["data"])
        number = str(row["collector_number"]).lower()
        loose = number.rstrip("abcdefghijklmnopqrstuvwxyz★s") if number[:1].isdigit() else number
        if number == wanted:
            number_rank = 0
        elif _number_key(loose) == wanted_key:
            number_rank = 1
        else:
            return None
        return (0 if hint and hint in (data.get("promo_types") or []) else 1, number_rank), data

    ranked = [r for r in (rank(row) for row in rows) if r]
    if not ranked:
        return None
    # min() keeps the first of equals, i.e. the paper/newest order from the query.
    return min(ranked, key=lambda r: r[0])[1]


def by_name(card_name: str) -> Optional[dict]:
    key = card_name.split(" // ")[0].strip().lower()
    return _query(f"SELECT data FROM cards WHERE name_key = ? OR front_key = ? {_ORDER} LIMIT 1",
                  (card_name.strip().lower(), key))


def record_miss(card_name: str, set_name: Optional[str], collector_number: Optional[str], found_live: bool) -> None:
    if not DB_PATH.exists():
        return
    try:
        with _lock:
            conn = _connect()
            conn.execute("INSERT OR REPLACE INTO misses VALUES (?, ?, ?, ?, ?)",
                         (card_name, set_name or "", collector_number or "", int(found_live),
                          datetime.now(timezone.utc).isoformat()))
            conn.commit()
            conn.close()
    except sqlite3.Error:
        pass


def status(recent_misses: int = 25) -> dict:
    if not DB_PATH.exists():
        return {"available": False, "refreshing": _refreshing.is_set()}
    with _lock:
        conn = _connect()
        misses = [dict(r) for r in conn.execute(
            "SELECT * FROM misses ORDER BY seen_at DESC LIMIT ?", (recent_misses,))]
        conn.close()
    return {
        "available": True,
        "refreshing": _refreshing.is_set(),
        "scryfall_updated_at": _meta("updated_at"),
        "checked_at": _meta("built_at"),
        "card_count": int(_meta("card_count") or 0),
        "recent_misses": misses,
    }


if __name__ == "__main__":
    import sys

    command = sys.argv[1] if len(sys.argv) > 1 else "status"
    if command == "refresh":
        refresh(force="--force" in sys.argv)
    elif command == "status":
        print(json.dumps(status(), indent=2))
    else:
        print("Usage: python scryfall_db.py [refresh [--force] | status]")
        sys.exit(1)
