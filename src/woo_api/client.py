from __future__ import annotations

import json
import time
from typing import Any, Callable, Iterator

import requests


API_PATH = "/wp-json/alkan-inator/v1"
RETRY_STATUS = {408, 429, 500, 502, 503, 504}
# Glowna wersja API wtyczki, z ktora ten klient umie pracowac (naglowek X-Alkan-Inator-Api-Version).
SUPPORTED_API_MAJOR = 2
API_VERSION_HEADER = "X-Alkan-Inator-Api-Version"


class WooBridgeError(RuntimeError):
    def __init__(self, message: str, *, status_code: int | None = None, code: str = "") -> None:
        super().__init__(message)
        self.status_code = status_code
        self.code = code


def check_api_version(version: str) -> None:
    """Klient i wtyczka musza miec te sama glowna wersje API, inaczej bezpieczniki zapisu moglyby nie dzialac."""
    major = version.split(".", 1)[0].strip()
    if not major:
        raise WooBridgeError(
            "Wtyczka AlkanInator Bridge w sklepie jest starsza niż ten AlkanInator (brak wersji API). "
            "Zaktualizuj wtyczkę w sklepie."
        )
    if not major.isdigit() or int(major) != SUPPORTED_API_MAJOR:
        raise WooBridgeError(
            f"Sklep ma API w wersji {version}, a ten AlkanInator obsługuje {SUPPORTED_API_MAJOR}.x. "
            + ("Zaktualizuj AlkanInator (git pull main)." if major.isdigit() and int(major) > SUPPORTED_API_MAJOR else "Zaktualizuj wtyczkę w sklepie.")
        )


class WooBridgeClient:
    """Klient REST wtyczki AlkanInator Bridge; token idzie w nagłówku X-Alkan-Inator-Token."""

    def __init__(
        self,
        base_url: str,
        token: str,
        *,
        verify_ssl: bool = True,
        timeout_seconds: float = 300.0,
        max_retries: int = 3,
        session: requests.Session | None = None,
        log: Callable[[str], None] | None = None,
    ) -> None:
        self.api_url = base_url.rstrip("/") + API_PATH
        self.timeout_seconds = timeout_seconds
        self.max_retries = max_retries
        self.log = log or (lambda _: None)
        self.session = session or requests.Session()
        self.session.headers.update({"X-Alkan-Inator-Token": token, "Accept": "application/json"})
        self.session.verify = verify_ssl
        if not verify_ssl:
            import urllib3

            urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

    def __enter__(self) -> "WooBridgeClient":
        return self

    def __exit__(self, *_: object) -> None:
        self.session.close()

    def _request(
        self, method: str, path: str, *, params: dict[str, Any] | None = None, body: Any = None, retry: bool = True
    ) -> Any:
        url = self.api_url + path
        max_retries = self.max_retries if retry else 0
        for attempt in range(max_retries + 1):
            try:
                response = self.session.request(method, url, params=params, json=body, timeout=self.timeout_seconds)
            except requests.RequestException as exc:
                if attempt >= max_retries:
                    raise WooBridgeError(f"Brak połączenia z {url}: {exc}") from exc
                self._wait(attempt, str(exc))
                continue

            if response.status_code in RETRY_STATUS and attempt < max_retries:
                self._wait(attempt, f"HTTP {response.status_code}")
                continue

            # Some WordPress installs prepend a UTF-8 BOM to every response.
            text = response.content.decode("utf-8-sig", errors="replace")
            try:
                payload = json.loads(text) if text.strip() else None
            except json.JSONDecodeError as exc:
                raise WooBridgeError(
                    f"Odpowiedź {url} nie jest JSON-em (HTTP {response.status_code}): {text[:300]}",
                    status_code=response.status_code,
                ) from exc

            if response.status_code >= 400:
                message = payload.get("message", text[:300]) if isinstance(payload, dict) else text[:300]
                code = payload.get("code", "") if isinstance(payload, dict) else ""
                raise WooBridgeError(f"HTTP {response.status_code}: {message}", status_code=response.status_code, code=code)
            check_api_version(response.headers.get(API_VERSION_HEADER, ""))
            return payload
        raise WooBridgeError(f"Nie udało się wykonać {method} {url}")

    def _wait(self, attempt: int, reason: str) -> None:
        delay = 2.0 * (2**attempt)
        self.log(f"Ponawiam za {delay:.0f} s ({reason})")
        time.sleep(delay)

    def knowledge(self) -> dict[str, Any]:
        return self._request("GET", "/knowledge")

    def products_page(self, page: int = 1, per_page: int = 200, **filters: Any) -> dict[str, Any]:
        body = {key: value for key, value in filters.items() if value not in (None, "", [], ())}
        body.update({"page": page, "per_page": per_page})
        # POST so long SKU lists fit in the body.
        return self._request("POST", "/products", body=body)

    def iter_product_pages(self, per_page: int = 200, **filters: Any) -> Iterator[dict[str, Any]]:
        page = 1
        while True:
            data = self.products_page(page=page, per_page=per_page, **filters)
            yield data
            if page >= int(data.get("total_pages") or 0):
                return
            page += 1

    def sync(
        self,
        rows: list[dict[str, Any]],
        *,
        dry_run: bool = True,
        allow_new_terms: bool = False,
        allow_destructive: bool = False,
        columns: list[str] | None = None,
        changeset: str = "",
    ) -> dict[str, Any]:
        body: dict[str, Any] = {
            "rows": rows,
            "dry_run": dry_run,
            "allow_new_terms": allow_new_terms,
            "allow_destructive": allow_destructive,
        }
        if columns:
            body["columns"] = columns
        if changeset:
            body["changeset"] = changeset
        # A repeated write could download the same images twice, so writes are never retried.
        return self._request("POST", "/products/sync", body=body, retry=dry_run)

    def changesets(self, limit: int = 50) -> list[dict[str, Any]]:
        return self._request("GET", "/changesets", params={"limit": limit})

    def changeset(self, changeset: str) -> list[dict[str, Any]]:
        return self._request("GET", f"/changesets/{changeset}")

    def revert(self, changeset: str, *, dry_run: bool = True) -> dict[str, Any]:
        return self._request("POST", f"/changesets/{changeset}/revert", body={"dry_run": dry_run}, retry=dry_run)
