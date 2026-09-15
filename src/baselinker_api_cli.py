from __future__ import annotations

import argparse
import json
import os
import sys
from pathlib import Path

from baselinker_api.client import BaseLinkerClient
from baselinker_api.export import ExportOptions, export_inventory_csv
from baselinker_api.sync import DEFAULT_FIELDS, SyncOptions, sync_inventory


ALLOWED_FIELDS = DEFAULT_FIELDS | {"identifiers"}
DEFAULT_CONFIG_PATH = Path(__file__).resolve().parent / "baselinker_api" / "config.json"


def _token_from_file(path: Path) -> str:
    raw = path.read_text(encoding="utf-8-sig").strip()
    if not raw:
        return ""
    if raw.startswith("{"):
        try:
            payload = json.loads(raw)
        except json.JSONDecodeError as exc:
            raise ValueError(f"Niepoprawny JSON w pliku tokenu {path}: {exc.msg}") from exc
        if not isinstance(payload, dict):
            raise ValueError(f"Plik konfiguracyjny {path} musi zawierać obiekt JSON.")
        return str(payload.get("token", "")).strip()
    return raw


def _token(args: argparse.Namespace) -> str:
    if args.token_file:
        token_path = Path(args.token_file)
        token = _token_from_file(token_path)
    else:
        token = os.environ.get(args.token_env, "").strip()
        if not token and DEFAULT_CONFIG_PATH.is_file():
            token = _token_from_file(DEFAULT_CONFIG_PATH)
    if not token:
        raise ValueError(
            f"Brak tokenu. Ustaw zmienną środowiskową {args.token_env}, użyj --token-file "
            f"albo zapisz ignorowany config w {DEFAULT_CONFIG_PATH}."
        )
    return token


def _positive_int(value: str) -> int:
    parsed = int(value)
    if parsed <= 0:
        raise argparse.ArgumentTypeError("Wartość musi być większa od zera.")
    return parsed


def _add_connection_args(parser: argparse.ArgumentParser) -> None:
    parser.add_argument("--token-env", default="BASELINKER_TOKEN", help="Nazwa zmiennej środowiskowej z tokenem.")
    parser.add_argument(
        "--token-file",
        default="",
        help="Plik zawierający token albo JSON z polem token; trzymaj go poza repo lub w ignorowanej lokalizacji.",
    )
    parser.add_argument("--timeout", type=float, default=60.0)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Bezpośredni eksport i bezpieczna synchronizacja AlkanInatora z API BaseLinkera."
    )
    subparsers = parser.add_subparsers(dest="command", required=True)

    inventories = subparsers.add_parser("inventories", help="Wyświetla dostępne katalogi BaseLinkera.")
    _add_connection_args(inventories)

    export = subparsers.add_parser("export", help="Eksportuje katalog do kanonicznego CSV AlkanInatora.")
    _add_connection_args(export)
    export.add_argument("--inventory-id", type=_positive_int, required=True)
    export.add_argument("--output", required=True)
    export.add_argument("--category-id", type=int)
    export.add_argument("--manufacturer-id", type=int)
    export.add_argument("--sku", default="")
    export.add_argument("--ean", default="")
    export.add_argument("--include-variants", action="store_true")
    export.add_argument("--language", default="pl")
    export.add_argument("--requests-per-minute", type=int, default=100)
    export.add_argument("--max-retries", type=int, default=5)

    sync = subparsers.add_parser("sync", help="Waliduje lub wysyła plik CSV/XLSX/JSON do BaseLinkera.")
    _add_connection_args(sync)
    sync.add_argument("--inventory-id", type=_positive_int, required=True)
    sync.add_argument("--input", required=True)
    sync.add_argument("--report-dir", default="")
    sync.add_argument(
        "--fields",
        default=",".join(sorted(DEFAULT_FIELDS)),
        help="Lista: name,manufacturer,category,description,features,images,identifiers.",
    )
    sync.add_argument("--match-by", choices=["auto", "product_id", "sku", "ean"], default="auto")
    sync.add_argument("--include-variants", action="store_true")
    sync.add_argument("--replace-features", action="store_true", help="Zastępuje wszystkie cechy zamiast bezpiecznego merge.")
    sync.add_argument("--create-missing-categories", action="store_true")
    sync.add_argument("--create-missing-manufacturers", action="store_true")
    sync.add_argument("--allow-http-images", action="store_true", help="Dopuszcza niezabezpieczone adresy HTTP zdjęć.")
    sync.add_argument("--allow-partial", action="store_true", help="Wysyła poprawne wiersze mimo błędów innych rekordów.")
    sync.add_argument("--language", default="pl")
    sync.add_argument("--requests-per-minute", type=int, default=100)
    sync.add_argument("--max-retries", type=int, default=5)
    sync.add_argument(
        "--apply",
        action="store_true",
        help="Wykonuje zapis. Bez tej flagi polecenie zawsze działa jako dry-run.",
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    try:
        token = _token(args)
        with BaseLinkerClient(token, timeout_seconds=args.timeout) as client:
            if args.command == "inventories":
                inventories = client.get_inventories()
                for inventory in inventories:
                    inventory_id = inventory.get("inventory_id", inventory.get("id", ""))
                    name = inventory.get("name", inventory.get("inventory_name", ""))
                    print(f"{inventory_id}\t{name}")
                return 0

            if args.command == "export":
                options = ExportOptions(
                    inventory_id=args.inventory_id,
                    include_variants=args.include_variants,
                    category_id=args.category_id,
                    manufacturer_id=args.manufacturer_id,
                    sku=args.sku,
                    ean=args.ean,
                    language=args.language,
                    requests_per_minute=args.requests_per_minute,
                    max_retries=args.max_retries,
                )
                count = export_inventory_csv(client, args.output, options, log=print)
                print(f"Gotowe: {Path(args.output).resolve()} ({count} produktów)")
                return 0

            fields = {item.strip().casefold() for item in args.fields.split(",") if item.strip()}
            unknown = fields - ALLOWED_FIELDS
            if unknown:
                parser.error(f"Nieznane pola --fields: {', '.join(sorted(unknown))}")
            update_identifiers = "identifiers" in fields
            fields.discard("identifiers")
            options = SyncOptions(
                inventory_id=args.inventory_id,
                apply=args.apply,
                fields=fields,
                match_by=args.match_by,
                include_variants=args.include_variants,
                merge_features=not args.replace_features,
                update_identifiers=update_identifiers,
                create_missing_categories=args.create_missing_categories,
                create_missing_manufacturers=args.create_missing_manufacturers,
                allow_partial=args.allow_partial,
                allow_http_images=args.allow_http_images,
                language=args.language,
                requests_per_minute=args.requests_per_minute,
                max_retries=args.max_retries,
            )
            result = sync_inventory(
                client,
                args.input,
                options,
                report_dir=args.report_dir or None,
                log=print,
            )
            counts: dict[str, int] = {}
            for row in result.audit_rows:
                counts[row.status] = counts.get(row.status, 0) + 1
            print(f"Tryb: {'APPLY' if args.apply else 'DRY_RUN'}")
            print(f"Raport: {result.report_dir.resolve()}")
            print("Statusy: " + ", ".join(f"{key}={value}" for key, value in sorted(counts.items())))
            return 2 if result.blocked or any(key in counts for key in ("INVALID", "ERROR", "BLOCKED_BATCH")) else 0
    except KeyboardInterrupt:
        print("Przerwano przez użytkownika.", file=sys.stderr)
        return 130
    except Exception as exc:
        print(f"Błąd: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
