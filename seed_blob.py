"""
seed_blob.py — one-off: upload existing local JSON data into the configured
Azure Blob container (CONTAINER). Run once after setting BLOB_CONNECTION_STRING
to migrate quote_data.json (job cache) and quotes.json (saved-quote index).

    python seed_blob.py
"""

from pathlib import Path
from dotenv import load_dotenv

load_dotenv()
import storage  # noqa: E402  (must load env first)

LOCAL = Path(__file__).resolve().parent.parent / "static"

print("Storage backend:", storage.backend())
if not storage.blob_enabled():
    print("BLOB_CONNECTION_STRING not set — nothing to do (already local).")
else:
    for name in ("quote_data.json", "quotes.json"):
        p = LOCAL / name
        if p.exists():
            storage.put_text(name, p.read_text(encoding="utf-8"))
            print(f"  uploaded {name} ({p.stat().st_size:,} bytes)")
        else:
            print(f"  skip (missing locally): {name}")
    print("Done.")
