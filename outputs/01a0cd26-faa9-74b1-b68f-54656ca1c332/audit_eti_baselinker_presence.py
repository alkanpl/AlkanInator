"""Read-only check of ETI SKU presence in available BaseLinker inventories."""

from __future__ import annotations

import json
from pathlib import Path
from urllib.parse import urlencode
from urllib.request import Request, urlopen


ROOT = Path(__file__).resolve().parents[2]
CONFIG = ROOT / "src" / "baselinker_api" / "config.json"
TOKEN = json.loads(CONFIG.read_text(encoding="utf-8-sig"))["token"]
ENDPOINT = "https://api.baselinker.com/connector.php"


def call(method: str, parameters: dict) -> dict:
    body = urlencode({"method": method, "parameters": json.dumps(parameters, separators=(",", ":"))}).encode()
    request = Request(ENDPOINT, data=body, headers={"X-BLToken": TOKEN})
    with urlopen(request, timeout=30) as response:
        result = json.load(response)
    if result.get("status") != "SUCCESS":
        raise RuntimeError(f"{method}: {result.get('error_code')} {result.get('error_message')}")
    return result


def main() -> None:
    inventories = call("getInventories", {})["inventories"]
    for inventory in inventories:
        inventory_id = int(inventory["inventory_id"])
        name = inventory.get("name", "")
        for sku in ("002111512/ETI", "002061261/ETI", "002540203/ETI"):
            result = call("getInventoryProductsList", {
                "inventory_id": inventory_id,
                "page": 1,
                "filter_sku": sku,
                "include_variants": False,
            })
            products = result.get("products", {})
            rows = list(products.values()) if isinstance(products, dict) else products
            exact = [row for row in rows if str(row.get("sku", "")).casefold() == sku.casefold()]
            print(f"inventory_id={inventory_id} name={name} sku={sku} exact_matches={len(exact)}")


if __name__ == "__main__":
    main()
