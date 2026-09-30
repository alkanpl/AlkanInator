"""Klient e-katalogu Kontakt-Simon (https://ekatalog.kontakt-simon.com.pl/api/v2/).

Uzycie z innych skryptow:

    from simon_api import SimonApiClient, load_api_key, flatten_category_tree
    client = SimonApiClient(load_api_key())
    products = client.fetch_products(["TR1/111", "DW1.01/11"])   # dict symbol -> JSON
    tree = client.fetch_category_tree()
    paths = flatten_category_tree(tree)                          # dict article_id -> [[nazwy kategorii], ...]

Klucz API: zmienna srodowiskowa SIMON_API_KEY albo lokalny plik
``Simon API/Program.php`` w katalogu glownym repo (``define("API_KEY", "...")``).
Plik z kluczem nie jest w gicie i nie wolno go kopiowac do innych plikow.

Odpowiedzi API sa cache'owane w ``cache/simon_api/`` (jeden JSON na artykul,
klucz = symbol producenta), zeby powtorne uruchomienia nie odpytywaly serwera.
"""

from __future__ import annotations

import json
import os
import re
import ssl
import time
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path
from typing import Iterable

ROOT = Path(__file__).resolve().parents[1]
API_BASE = "https://ekatalog.kontakt-simon.com.pl/api/v2/"
DEFAULT_CACHE_DIR = ROOT / "cache" / "simon_api"
API_KEY_FILES = [
    ROOT / "Simon API" / "Program.php",
    ROOT / "Simon API" / "Final.php",
    ROOT / "archived_input_files" / "Simon API" / "Program.php",
]
CHUNK_SIZE = 25


class SimonApiError(RuntimeError):
    pass


def load_api_key(explicit: str | None = None) -> str:
    if explicit:
        return explicit.strip()
    env = os.environ.get("SIMON_API_KEY", "").strip()
    if env:
        return env
    for path in API_KEY_FILES:
        if not path.exists():
            continue
        text = path.read_text(encoding="utf-8", errors="ignore")
        match = re.search(r'define\s*\(\s*["\']API_KEY["\']\s*,\s*["\']([^"\']+)["\']', text, flags=re.I)
        if match:
            return match.group(1).strip()
    raise SimonApiError(
        "Nie znaleziono klucza API Simon. Ustaw SIMON_API_KEY albo polozyc plik "
        "'Simon API/Program.php' z define(\"API_KEY\", ...) w katalogu repo."
    )


def normalize_symbol(symbol: str) -> str:
    """Symbol producenta bez sufiksu sklepu (/KON, /KAN) i bez bialych znakow."""
    value = (symbol or "").strip().upper()
    value = re.sub(r"/(KON|KAN)$", "", value)
    return value


def api_identifier(symbol: str) -> str:
    return normalize_symbol(symbol).replace("/", "_")


def _items_from_payload(payload: object) -> list[dict]:
    if isinstance(payload, dict) and payload.get("symbol"):
        return [payload]
    if isinstance(payload, dict):
        items = payload.get("collection", {}).get("items", {})
        if isinstance(items, dict):
            return list(items.values())
        if isinstance(items, list):
            return items
    if isinstance(payload, list):
        return payload
    return []


def product_params(product: dict) -> dict[str, str]:
    """Parametry artykulu jako slownik nazwa -> wartosc (pierwsze wystapienie wygrywa)."""
    result: dict[str, str] = {}
    for param in product.get("params") or []:
        if not isinstance(param, dict):
            continue
        name = str(param.get("name") or "").strip()
        value = param.get("value")
        if name and value is not None and name not in result:
            result[name] = str(value).strip()
    return result


def flatten_category_tree(tree: dict) -> dict[int, list[list[str]]]:
    """Mapa id artykulu -> lista sciezek kategorii (nazwy od korzenia do liscia)."""
    result: dict[int, list[list[str]]] = {}

    def walk(node: dict, names: list[str]) -> None:
        current = names + [str(node.get("name") or "").strip()]
        for article_id in node.get("articleIdList") or []:
            try:
                key = int(article_id)
            except (TypeError, ValueError):
                continue
            result.setdefault(key, []).append(current)
        subs = node.get("subCategories") or {}
        children = subs.values() if isinstance(subs, dict) else subs
        for child in children:
            if isinstance(child, dict):
                walk(child, current)

    walk(tree, [])
    return result


class SimonApiClient:
    def __init__(
        self,
        api_key: str,
        cache_dir: Path | None = None,
        sleep_seconds: float = 0.4,
        timeout: int = 60,
        verbose: bool = False,
    ) -> None:
        self.api_key = api_key
        self.cache_dir = Path(cache_dir) if cache_dir else DEFAULT_CACHE_DIR
        self.articles_dir = self.cache_dir / "articles"
        self.articles_dir.mkdir(parents=True, exist_ok=True)
        self.sleep_seconds = sleep_seconds
        self.timeout = timeout
        self.verbose = verbose
        self.ssl_context = ssl.create_default_context()
        self.requests_made = 0

    # -- HTTP ---------------------------------------------------------------
    def _get_json(self, path: str) -> object:
        url = API_BASE + path
        request = urllib.request.Request(
            url, headers={"X-ApiKey": self.api_key, "Accept": "application/json"}
        )
        delay = 2.0
        for attempt in range(6):
            try:
                with urllib.request.urlopen(request, timeout=self.timeout, context=self.ssl_context) as response:
                    self.requests_made += 1
                    raw = response.read().decode("utf-8")
                if self.sleep_seconds:
                    time.sleep(self.sleep_seconds)
                return json.loads(raw)
            except urllib.error.HTTPError as exc:
                body = exc.read().decode("utf-8", errors="ignore")
                if exc.code == 404:
                    return None
                if exc.code == 401:
                    raise SimonApiError(f"API Simon odrzucilo klucz (401): {body[:200]}") from exc
                if exc.code in (429, 500, 502, 503, 504) and attempt < 5:
                    if self.verbose:
                        print(f"  HTTP {exc.code}, ponawiam za {delay:.0f}s")
                    time.sleep(delay)
                    delay *= 2
                    continue
                raise SimonApiError(f"HTTP {exc.code} dla {url}: {body[:200]}") from exc
            except (urllib.error.URLError, TimeoutError) as exc:
                if attempt < 5:
                    time.sleep(delay)
                    delay *= 2
                    continue
                raise SimonApiError(f"Blad polaczenia z {url}: {exc}") from exc
        raise SimonApiError(f"Nie udalo sie pobrac {url}")

    # -- cache --------------------------------------------------------------
    def _cache_path(self, symbol: str) -> Path:
        safe = re.sub(r"[^A-Za-z0-9._-]", "_", normalize_symbol(symbol))
        return self.articles_dir / f"{safe}.json"

    def cached_product(self, symbol: str) -> dict | None:
        path = self._cache_path(symbol)
        if not path.exists():
            return None
        try:
            return json.loads(path.read_text(encoding="utf-8"))
        except json.JSONDecodeError:
            return None

    def _store(self, symbol: str, product: dict | None) -> None:
        payload = product if product is not None else {"symbol": normalize_symbol(symbol), "_missing": True}
        self._cache_path(symbol).write_text(json.dumps(payload, ensure_ascii=False, indent=1), encoding="utf-8")

    # -- public -------------------------------------------------------------
    def fetch_products(self, symbols: Iterable[str], refresh: bool = False) -> dict[str, dict]:
        """Zwraca slownik znormalizowany symbol -> JSON artykulu.

        Brak w API jest zapamietywany jako ``{"symbol": ..., "_missing": true}``.
        """
        wanted: list[str] = []
        seen: set[str] = set()
        for symbol in symbols:
            key = normalize_symbol(symbol)
            if key and key not in seen:
                seen.add(key)
                wanted.append(key)

        result: dict[str, dict] = {}
        to_fetch: list[str] = []
        for key in wanted:
            cached = None if refresh else self.cached_product(key)
            if cached is not None:
                result[key] = cached
            else:
                to_fetch.append(key)

        for offset in range(0, len(to_fetch), CHUNK_SIZE):
            chunk = to_fetch[offset : offset + CHUNK_SIZE]
            identifiers = ",".join(urllib.parse.quote(api_identifier(s), safe="._-") for s in chunk)
            if self.verbose:
                print(f"  API Simon: {offset + len(chunk)}/{len(to_fetch)} symboli")
            payload = self._get_json(f"catalog/article/{identifiers}.json")
            found: dict[str, dict] = {}
            for item in _items_from_payload(payload):
                sym = normalize_symbol(str(item.get("symbol") or ""))
                if sym:
                    found[sym] = item
            for key in chunk:
                product = found.get(key)
                if product is None and len(chunk) > 1:
                    # API potrafi pominac pojedynczy symbol w paczce - sprawdz osobno.
                    single = self._get_json(f"catalog/article/{urllib.parse.quote(api_identifier(key), safe='._-')}.json")
                    items = _items_from_payload(single)
                    product = items[0] if items else None
                self._store(key, product)
                result[key] = product if product is not None else {"symbol": key, "_missing": True}
        return result

    def fetch_category_tree(self, refresh: bool = False) -> dict:
        path = self.cache_dir / "category_tree.json"
        if path.exists() and not refresh:
            return json.loads(path.read_text(encoding="utf-8"))
        payload = self._get_json("catalog/category/tree.json")
        if not isinstance(payload, dict):
            raise SimonApiError("API Simon nie zwrocilo drzewa kategorii")
        path.write_text(json.dumps(payload, ensure_ascii=False, indent=1), encoding="utf-8")
        return payload


def is_missing(product: dict | None) -> bool:
    return product is None or bool(product.get("_missing"))
