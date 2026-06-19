"""
storage.py
Tiny persistence layer for the quote tool's JSON data
(quote_data.json — job cache, quotes.json — saved-quote index).

If BLOB_CONNECTION_STRING is set it stores the files in Azure Blob Storage
(container = CONTAINER env, default "acme-online-quote"). Otherwise it falls
back to local files under QUOTE_DATA_DIR / ../static — so local dev is unchanged.

This is the right model for Azure App Service: the data survives restarts and
re-deploys and is shared across instances, instead of living on the ephemeral
app filesystem.
"""

import os
import json
import time
import threading
from pathlib import Path
from typing import Any, Optional

try:
    from dotenv import load_dotenv
    load_dotenv()
except Exception:
    pass

# Env is read lazily (at call time) so it works regardless of when this module
# is imported relative to load_dotenv().
def _conn() -> str:
    return (os.getenv("BLOB_CONNECTION_STRING") or "").strip()

def _container_name() -> str:
    return (os.getenv("CONTAINER") or "acme-online-quote").strip()

def _local_dir() -> Path:
    return Path(os.getenv("QUOTE_DATA_DIR") or (Path(__file__).resolve().parent.parent / "static"))

_cache: dict = {}          # name -> (expires_ts, text)
_lock = threading.Lock()
_svc = None


def blob_enabled() -> bool:
    return bool(_conn())


def backend() -> str:
    return f"blob:{_container_name()}" if blob_enabled() else f"local:{_local_dir()}"


def _container():
    global _svc
    from azure.storage.blob import BlobServiceClient
    if _svc is None:
        _svc = BlobServiceClient.from_connection_string(_conn())
    cc = _svc.get_container_client(_container_name())
    try:
        cc.create_container()
    except Exception:
        pass  # already exists
    return cc


def get_text(name: str, default: str = "", ttl: int = 0) -> str:
    """Read a blob/file as text. ttl>0 caches the result in-process."""
    now = time.time()
    if ttl > 0:
        with _lock:
            hit = _cache.get(name)
            if hit and hit[0] > now:
                return hit[1]
    text = default
    got = False
    try:
        if blob_enabled():
            data = _container().get_blob_client(name).download_blob().readall()
            text = data.decode("utf-8")
            got = True
        else:
            p = _local_dir() / name
            if p.exists():
                text = p.read_text(encoding="utf-8")
                got = True
    except Exception:
        got = False  # blob cold-start / missing / network — do NOT cache the miss
    # Only cache successful reads, so a transient failure isn't pinned for `ttl`.
    if ttl > 0 and got:
        with _lock:
            _cache[name] = (now + ttl, text)
    return text


def put_text(name: str, text: str) -> None:
    if blob_enabled():
        _container().get_blob_client(name).upload_blob(
            text.encode("utf-8"), overwrite=True)
    else:
        d = _local_dir()
        try:
            d.mkdir(parents=True, exist_ok=True)
        except OSError:
            pass
        (d / name).write_text(text, encoding="utf-8")
    with _lock:
        _cache.pop(name, None)


_bytes_cache: dict = {}     # name -> (expires_ts, bytes)


def get_bytes(name: str, ttl: int = 600) -> Optional[bytes]:
    """Read a blob/file as raw bytes (for logos etc). ttl>0 caches in-process.
    Returns None if missing or on a transient failure (the miss is NOT cached)."""
    now = time.time()
    if ttl > 0:
        with _lock:
            hit = _bytes_cache.get(name)
            if hit and hit[0] > now:
                return hit[1]
    data: Optional[bytes] = None
    try:
        if blob_enabled():
            data = _container().get_blob_client(name).download_blob().readall()
        else:
            p = _local_dir() / name
            if p.exists():
                data = p.read_bytes()
    except Exception:
        data = None
    if ttl > 0 and data is not None:
        with _lock:
            _bytes_cache[name] = (now + ttl, data)
    return data


def put_bytes(name: str, data: bytes, content_type: str = "application/octet-stream") -> None:
    if blob_enabled():
        from azure.storage.blob import ContentSettings
        _container().get_blob_client(name).upload_blob(
            data, overwrite=True, content_settings=ContentSettings(content_type=content_type))
    else:
        d = _local_dir()
        try:
            d.mkdir(parents=True, exist_ok=True)
        except OSError:
            pass
        (d / name).write_bytes(data)
    with _lock:
        _bytes_cache.pop(name, None)


def get_json(name: str, default: Any = None, ttl: int = 0) -> Any:
    raw = get_text(name, "", ttl)
    if not raw:
        return default
    try:
        return json.loads(raw)
    except Exception:
        return default


def put_json(name: str, obj: Any) -> None:
    put_text(name, json.dumps(obj, ensure_ascii=False, indent=2))


def invalidate(name: Optional[str] = None) -> None:
    with _lock:
        if name:
            _cache.pop(name, None)
        else:
            _cache.clear()
