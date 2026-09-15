from __future__ import annotations

import json
import random
import re
import time
from collections import deque
from dataclasses import dataclass
from typing import Any, Callable

import requests


BASELINKER_ENDPOINT = "https://api.baselinker.com/connector.php"


class BaseLinkerError(RuntimeError):
    def __init__(self, message: str, *, code: str = "", status_code: int | None = None) -> None:
        super().__init__(message)
        self.code = code
        self.status_code = status_code


class BaseLinkerResponseError(BaseLinkerError):
    pass


@dataclass
class RateLimiter:
    requests_per_minute: int = 100

    def __post_init__(self) -> None:
        self._timestamps: deque[float] = deque()

    def wait(self) -> None:
        if self.requests_per_minute <= 0:
            return
        now = time.monotonic()
        while self._timestamps and now - self._timestamps[0] >= 60.0:
            self._timestamps.popleft()
        if len(self._timestamps) >= self.requests_per_minute:
            delay = max(0.0, 60.0 - (now - self._timestamps[0]))
            if delay:
                time.sleep(delay)
            now = time.monotonic()
            while self._timestamps and now - self._timestamps[0] >= 60.0:
                self._timestamps.popleft()
        self._timestamps.append(time.monotonic())


def _blocked_delay(message: str) -> float | None:
    match = re.search(r"(?:until|do)\D*(\d{10})", message, flags=re.IGNORECASE)
    if not match:
        return None
    return max(0.0, float(match.group(1)) - time.time() + 2.0)


def call_with_retry(
    fn: Callable[[], Any],
    *,
    limiter: RateLimiter,
    max_retries: int = 5,
    backoff_base_seconds: float = 1.5,
    log: Callable[[str], None] | None = None,
) -> Any:
    logger = log or (lambda _: None)
    for attempt in range(max(0, max_retries) + 1):
        limiter.wait()
        try:
            return fn()
        except (requests.RequestException, BaseLinkerError) as exc:
            retryable = isinstance(exc, requests.RequestException)
            if isinstance(exc, BaseLinkerError):
                retryable = exc.status_code in {408, 429, 500, 502, 503, 504} or any(
                    marker in exc.code.upper() for marker in ("LIMIT", "TIMEOUT", "BLOCKED")
                )
            if not retryable or attempt >= max_retries:
                raise
            blocked = _blocked_delay(str(exc))
            delay = blocked if blocked is not None else backoff_base_seconds * (2**attempt) + random.uniform(0, 0.25)
            logger(f"BaseLinker: ponawiam za {delay:.1f} s (próba {attempt + 2}/{max_retries + 1}).")
            time.sleep(delay)
    raise AssertionError("Nieosiągalny koniec mechanizmu retry")


class BaseLinkerClient:
    """Wąski klient oficjalnego endpointu BaseLinkera, bez logowania sekretów."""

    def __init__(
        self,
        token: str,
        *,
        timeout_seconds: float = 60.0,
        session: requests.Session | None = None,
    ) -> None:
        if not token.strip():
            raise ValueError("Brak tokenu BaseLinkera.")
        self._token = token.strip()
        self.timeout_seconds = float(timeout_seconds)
        self.session = session or requests.Session()
        self._owns_session = session is None

    def close(self) -> None:
        if self._owns_session:
            self.session.close()

    def __enter__(self) -> BaseLinkerClient:
        return self

    def __exit__(self, *_: object) -> None:
        self.close()

    def call(self, method: str, parameters: dict[str, Any] | None = None) -> dict[str, Any]:
        response = self.session.post(
            BASELINKER_ENDPOINT,
            headers={"X-BLToken": self._token},
            data={
                "method": method,
                "parameters": json.dumps(parameters or {}, ensure_ascii=False, separators=(",", ":")),
            },
            timeout=self.timeout_seconds,
        )
        if response.status_code != 200:
            raise BaseLinkerResponseError(
                f"HTTP {response.status_code} podczas {method}.",
                code=f"HTTP_{response.status_code}",
                status_code=response.status_code,
            )
        try:
            payload = response.json()
        except ValueError as exc:
            raise BaseLinkerResponseError(f"Niepoprawna odpowiedź JSON dla {method}.") from exc
        if not isinstance(payload, dict):
            raise BaseLinkerResponseError(f"Nieoczekiwany format odpowiedzi dla {method}.")
        if str(payload.get("status", "")).upper() != "SUCCESS":
            code = str(payload.get("error_code", "UNKNOWN"))
            message = str(payload.get("error_message", "Nieznany błąd API"))
            raise BaseLinkerResponseError(f"{code}: {message}", code=code)
        return payload

    def get_inventories(self) -> list[dict[str, Any]]:
        payload = self.call("getInventories")
        return self._object_rows(payload.get("inventories", []), "inventory_id")

    @staticmethod
    def _object_rows(value: Any, id_field: str) -> list[dict[str, Any]]:
        if isinstance(value, list):
            return [item for item in value if isinstance(item, dict)]
        if isinstance(value, dict):
            rows: list[dict[str, Any]] = []
            for key, item in value.items():
                if not isinstance(item, dict):
                    continue
                row = dict(item)
                row.setdefault(id_field, key)
                rows.append(row)
            return rows
        return []

    def get_products_list(
        self,
        inventory_id: int,
        *,
        page: int = 1,
        include_variants: bool = False,
        filter_category_id: int | None = None,
        filter_sku: str | None = None,
        filter_ean: str | None = None,
    ) -> list[dict[str, Any]]:
        params: dict[str, Any] = {
            "inventory_id": inventory_id,
            "page": page,
            "include_variants": include_variants,
        }
        if filter_category_id is not None:
            params["filter_category_id"] = filter_category_id
        if filter_sku:
            params["filter_sku"] = filter_sku
        if filter_ean:
            params["filter_ean"] = filter_ean
        value = self.call("getInventoryProductsList", params).get("products", [])
        if isinstance(value, list):
            return [item for item in value if isinstance(item, dict)]
        if isinstance(value, dict):
            rows: list[dict[str, Any]] = []
            for key, item in value.items():
                if not isinstance(item, dict):
                    continue
                row = dict(item)
                row.setdefault("id", key)
                rows.append(row)
            return rows
        return []

    def get_products_data(self, inventory_id: int, product_ids: list[int]) -> dict[str, dict[str, Any]]:
        value = self.call(
            "getInventoryProductsData",
            {"inventory_id": inventory_id, "products": product_ids},
        ).get("products", {})
        return {str(key): item for key, item in value.items() if isinstance(item, dict)} if isinstance(value, dict) else {}

    def get_categories(self, inventory_id: int) -> list[dict[str, Any]]:
        value = self.call("getInventoryCategories", {"inventory_id": inventory_id}).get("categories", [])
        return self._object_rows(value, "category_id")

    def add_category(self, parameters: dict[str, Any]) -> dict[str, Any]:
        return self.call("addInventoryCategory", parameters)

    def get_manufacturers(self) -> list[dict[str, Any]]:
        rows: list[dict[str, Any]] = []
        page = 1
        while True:
            value = self.call("getInventoryManufacturers", {"page": page}).get("manufacturers", [])
            current = self._object_rows(value, "manufacturer_id")
            rows.extend(current)
            if len(current) < 1000:
                break
            page += 1
        return rows

    def add_manufacturer(self, name: str) -> dict[str, Any]:
        return self.call("addInventoryManufacturer", {"manufacturer_name": name})

    def add_product(self, parameters: dict[str, Any]) -> dict[str, Any]:
        return self.call("addInventoryProduct", parameters)
