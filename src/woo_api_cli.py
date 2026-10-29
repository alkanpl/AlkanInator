"""Połączenie AlkanInatora ze sklepem WooCommerce (wtyczka AlkanInator Bridge).

    py src/woo_api_cli.py knowledge
    py src/woo_api_cli.py pull --category "Gniazdka Kontakt Simon" --output output/simon_woo.xlsx
    py src/woo_api_cli.py push --input output/simon_woo.xlsx                 (dry-run + raport)
    py src/woo_api_cli.py push --input output/simon_woo.xlsx --apply         (zapis)
    py src/woo_api_cli.py changesets
    py src/woo_api_cli.py revert ai-20261021-153000-simon_woo --apply
    py src/woo_api_cli.py --profile prod push --input plik.xlsx --apply     (produkcja: potwierdzenie w terminalu)

Połączenie: src/woo_api/config.json - jeden sklep ({"url": ..., "token": ..., "verify_ssl": ...},
traktowany jako profil "local") albo kilka:
    {"default": "local", "profiles": {"local": {...}, "prod": {"url": ..., "token": ..., "production": true}}}
Zmienne ALKAN_WOO_URL i ALKAN_WOO_TOKEN nadpisują adres i token.

Zapis na produkcję (--apply, revert --apply) wymaga wpisania liczby produktów w terminalu;
bez terminala (agent AI) jest odrzucany.
"""

from __future__ import annotations

import argparse
import json
import os
import sys
from datetime import datetime
from pathlib import Path
from urllib.parse import urlparse

from woo_api.client import WooBridgeClient, WooBridgeError
from woo_api.knowledge import write_knowledge
from woo_api.pull import pull_products, read_sku_file
from woo_api.push import PushOptions, run_push, run_revert


ROOT = Path(__file__).resolve().parents[1]
DEFAULT_CONFIG_PATH = Path(__file__).resolve().parent / "woo_api" / "config.json"
DEFAULT_REPORT_DIR = ROOT / "reports" / "woo_api"


def load_connection(args: argparse.Namespace) -> dict[str, object]:
    config: dict[str, object] = {}
    config_path = Path(args.config) if args.config else DEFAULT_CONFIG_PATH
    if config_path.is_file():
        try:
            config = json.loads(config_path.read_text(encoding="utf-8-sig"))
        except json.JSONDecodeError as exc:
            raise ValueError(f"Niepoprawny JSON w {config_path}: {exc.msg}") from exc
    elif args.config:
        raise ValueError(f"Nie ma pliku konfiguracji {config_path}")

    requested = getattr(args, "profile", "") or ""
    profiles = config.get("profiles")
    if isinstance(profiles, dict):
        name = requested or str(config.get("default", "local"))
        if not isinstance(profiles.get(name), dict):
            raise ValueError(f"W {config_path} nie ma profilu „{name}” (są: {', '.join(profiles) or 'brak'}).")
        profile: dict[str, object] = profiles[name]
    else:
        # Plaski format sprzed profili: jeden sklep, profil "local".
        if requested and requested != "local":
            raise ValueError(f"{config_path} opisuje jeden sklep (profil local); profil „{requested}” wymaga sekcji \"profiles\".")
        name, profile, profiles = "local", config, {}

    url = args.url or os.environ.get("ALKAN_WOO_URL", "") or str(profile.get("url", ""))
    token = os.environ.get("ALKAN_WOO_TOKEN", "") or str(profile.get("token", ""))
    verify_ssl = bool(profile.get("verify_ssl", True)) and not args.insecure
    if not url or not token:
        raise ValueError(
            f"Brak adresu sklepu albo tokenu. Uzupełnij {DEFAULT_CONFIG_PATH} "
            '({"url": "https://alkanlocal.test:8443", "token": "...", "verify_ssl": false}) '
            "albo ustaw ALKAN_WOO_URL i ALKAN_WOO_TOKEN. Token generuje się w WooCommerce > AlkanInator API."
        )

    # Adres produkcji podany przez --url albo ALKAN_WOO_URL nadal liczy sie jako produkcja.
    production_hosts = {
        urlparse(str(item.get("url", ""))).netloc.lower()
        for key, item in profiles.items()
        if isinstance(item, dict) and item.get("production", key == "prod")
    }
    production = bool(profile.get("production", name == "prod")) or urlparse(url).netloc.lower() in production_hosts
    return {"url": url, "token": token, "verify_ssl": verify_ssl, "profile": name, "production": production}


def _has_terminal() -> bool:
    """Czy na wejsciu jest konsola, w ktorej pisze czlowiek."""
    if sys.stdin is None or not sys.stdin.isatty():
        return False
    if os.name == "nt":
        # Na Windows isatty() jest prawdziwe takze dla wejscia z NUL; konsole rozpoznaje dopiero GetConsoleMode.
        import ctypes
        import msvcrt

        mode = ctypes.c_uint32()
        return bool(ctypes.windll.kernel32.GetConsoleMode(msvcrt.get_osfhandle(sys.stdin.fileno()), ctypes.byref(mode)))
    return True


def confirm_production(count: int, action: str) -> bool:
    """Zapis na produkcje zatwierdza czlowiek w terminalu; agent AI nie ma terminala, wiec dostaje odmowe."""
    if not _has_terminal():
        print(
            "Zapis na PRODUKCJĘ wymaga potwierdzenia w terminalu - uruchom to polecenie samodzielnie w konsoli "
            "(PowerShell albo cmd; w Git Bash: winpty py ...). "
            "Agent AI może przygotować plik i dry-run, ale nie zapisuje na produkcję.",
            file=sys.stderr,
        )
        return False
    try:
        answer = input(f"PRODUKCJA: {action} - produktów: {count}. Wpisz {count}, aby potwierdzić: ")
    except EOFError:
        return False
    return answer.strip() == str(count)


def _list(value: str) -> list[str]:
    return [item.strip() for item in value.split(",") if item.strip()]


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Wymiana danych AlkanInatora ze sklepem WooCommerce.")
    parser.add_argument("--config", default="", help=f"Plik JSON z url/token/verify_ssl (domyślnie {DEFAULT_CONFIG_PATH}).")
    parser.add_argument("--profile", default="", help='Profil z sekcji "profiles" w config.json, np. prod (domyślnie "default").')
    parser.add_argument("--url", default="", help="Adres sklepu, np. https://alkanlocal.test:8443.")
    parser.add_argument("--insecure", action="store_true", help="Nie sprawdza certyfikatu SSL (lokalny Laragon).")
    parser.add_argument("--timeout", type=float, default=300.0)
    subparsers = parser.add_subparsers(dest="command", required=True)

    knowledge = subparsers.add_parser("knowledge", help="Pobiera wiedzę o sklepie do dictionaries/ i cache/.")
    knowledge.add_argument("--dictionaries-dir", default=str(ROOT / "dictionaries"))
    knowledge.add_argument("--cache", default=str(ROOT / "cache" / "woo_api" / "knowledge.json"))

    pull = subparsers.add_parser("pull", help="Pobiera produkty do XLSX w układzie WP All Export.")
    pull.add_argument("--output", required=True)
    pull.add_argument("--category", default="", help="Kategorie (slug, ID albo ścieżka A>B) po przecinku; z podkategoriami.")
    pull.add_argument("--sku", default="", help="SKU po przecinku.")
    pull.add_argument("--sku-file", default="", help="Plik z SKU: XLSX z kolumną SKU albo TXT/CSV po jednym w linii.")
    pull.add_argument("--ids", default="", help="ID produktów po przecinku.")
    pull.add_argument("--status", default="", help="publish,draft,pending,private (domyślnie wszystkie).")
    pull.add_argument("--search", default="", help="Fraza w tytule/opisie.")
    pull.add_argument("--modified-after", default="", help="Tylko zmienione po dacie, np. 2026-10-01.")
    pull.add_argument("--all", action="store_true", help="Cały katalog (wymagane, gdy nie podano żadnego filtra).")
    pull.add_argument("--per-page", type=int, default=200)

    push = subparsers.add_parser("push", help="Porównuje plik XLSX ze sklepem (dry-run) i opcjonalnie zapisuje zmiany.")
    push.add_argument("--input", required=True)
    push.add_argument("--sheet", default="")
    push.add_argument("--columns", default="", help="Zapisuj tylko te kolumny (po przecinku); domyślnie wszystkie zapisywalne z pliku.")
    push.add_argument("--allow-new-terms", action="store_true", help="Pozwala tworzyć nowe tagi i wartości atrybutów.")
    push.add_argument(
        "--allow-destructive",
        action="store_true",
        help="Pozwala wyczyścić Content, Short Description albo Image URL (pusta komórka) - tylko na wyraźne życzenie.",
    )
    push.add_argument("--skip-errors", action="store_true", help="Zapisuje poprawne produkty mimo błędów w innych wierszach.")
    push.add_argument("--batch-size", type=int, default=100)
    push.add_argument("--report-dir", default=str(DEFAULT_REPORT_DIR))
    push.add_argument("--apply", action="store_true", help="Wykonuje zapis. Bez tej flagi zawsze tylko dry-run.")
    push.add_argument(
        "--skip-image-check",
        action="store_true",
        help="Nie sprawdza nowych zdjęć z Image URL (JPG 1000x1000 na białym tle) - tylko na wyraźne życzenie.",
    )

    changesets = subparsers.add_parser("changesets", help="Lista ostatnich zapisów w sklepie.")
    changesets.add_argument("--limit", type=int, default=20)

    revert = subparsers.add_parser("revert", help="Cofa zapis (domyślnie dry-run).")
    revert.add_argument("changeset")
    revert.add_argument("--report-dir", default=str(DEFAULT_REPORT_DIR))
    revert.add_argument("--apply", action="store_true")
    return parser


def _print_summary(label: str, result: dict) -> None:
    summary = result.get("summary", {})
    print(
        f"{label}: wierszy {summary.get('rows', 0)}, ze zmianami {summary.get('changed', 0)}, "
        f"zapisanych {summary.get('applied', 0)}, zmian pól {summary.get('changes', 0)}, "
        f"z błędami {summary.get('errors', 0)}, ostrzeżeń {summary.get('warnings', 0)}"
    )
    if result.get("new_terms"):
        print(f"Nowe terminy: {len(result['new_terms'])} (arkusz 'Nowe terminy')")
    if result.get("ignored_columns"):
        print("Pominięte nieznane kolumny: " + ", ".join(result["ignored_columns"]))


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    try:
        connection = load_connection(args)
        production = bool(connection["production"])
        print(f"Sklep: {connection['url']} (profil {connection['profile']}{', PRODUKCJA' if production else ''})")
        with WooBridgeClient(
            str(connection["url"]),
            str(connection["token"]),
            verify_ssl=bool(connection["verify_ssl"]),
            timeout_seconds=args.timeout,
            log=print,
        ) as client:
            if args.command == "knowledge":
                knowledge = client.knowledge()
                paths = write_knowledge(knowledge, Path(args.dictionaries_dir), Path(args.cache))
                print(
                    f"Kategorii {len(knowledge.get('categories', []))}, atrybutów {len(knowledge.get('attributes', []))}, "
                    f"tagów {len(knowledge.get('tags', []))}, klas wysyłkowych {len(knowledge.get('shipping_classes', []))}"
                )
                for path in paths.values():
                    print(f"-> {path}")
                return 0

            if args.command == "pull":
                skus = _list(args.sku) + (read_sku_file(args.sku_file) if args.sku_file else [])
                filters = {
                    "category": _list(args.category),
                    "sku": list(dict.fromkeys(skus)),
                    "include": _list(args.ids),
                    "status": _list(args.status),
                    "search": args.search,
                    "modified_after": args.modified_after,
                }
                if not any(filters.values()) and not args.all:
                    parser.error("Podaj filtr (--category, --sku, --sku-file, --ids, --search, --modified-after) albo --all.")
                count, warnings = pull_products(client, args.output, filters=filters, per_page=args.per_page)
                for warning in warnings:
                    print(f"UWAGA: {warning}")
                print(f"Gotowe: {Path(args.output).resolve()} ({count} produktów)")
                return 0

            if args.command == "push":
                options = PushOptions(
                    apply=args.apply,
                    columns=_list(args.columns),
                    allow_new_terms=args.allow_new_terms,
                    allow_destructive=args.allow_destructive,
                    skip_errors=args.skip_errors,
                    batch_size=args.batch_size,
                    sheet=args.sheet,
                    check_images=not args.skip_image_check,
                    confirm=(lambda count: confirm_production(count, "zapis zmian")) if production else None,
                )
                result = run_push(client, args.input, options, report_dir=args.report_dir)
                _print_summary("Dry-run", result.dry_run)
                if result.applied is not None:
                    _print_summary("Zapis", result.applied)
                    print(f"Zapis: {result.changeset}  (cofnięcie: py src/woo_api_cli.py revert {result.changeset} --apply)")
                for path in result.report_paths:
                    print(f"Raport: {path.resolve()}")
                if result.blocked:
                    return 2
                failed = result.applied if result.applied is not None else result.dry_run
                return 2 if failed.get("summary", {}).get("errors") else 0

            if args.command == "changesets":
                for item in client.changesets(args.limit):
                    reverted = f", cofnięte {item['reverted']}" if item.get("reverted") else ""
                    print(
                        f"{item['changeset']}\t{item['created_at']}\t{item.get('token_label') or item.get('user', '')}\t"
                        f"produkty {item['products']}, zmiany {item['changes']}{reverted}\t{', '.join(item.get('fields', []))}"
                    )
                return 0

            if args.command == "revert":
                result, report = run_revert(
                    client,
                    args.changeset,
                    apply=args.apply,
                    report_dir=args.report_dir,
                    now=datetime.now(),
                    confirm=(lambda count: confirm_production(count, f"cofnięcie {args.changeset}")) if production else None,
                )
                _print_summary("Cofnięcie (zapis)" if args.apply else "Cofnięcie (dry-run)", result)
                print(f"Raport: {report.resolve()}")
                return 2 if result.get("summary", {}).get("errors") else 0
    except KeyboardInterrupt:
        print("Przerwano przez użytkownika.", file=sys.stderr)
        return 130
    except (WooBridgeError, ValueError, OSError) as exc:
        print(f"Błąd: {exc}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
