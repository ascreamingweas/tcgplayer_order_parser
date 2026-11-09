import os
import sys
import tempfile
from pathlib import Path

# Tests never touch the real local Scryfall copy (scryfall_db reads this at import).
os.environ["SCRYFALL_DB"] = str(Path(tempfile.mkdtemp()) / "scryfall.sqlite")

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
