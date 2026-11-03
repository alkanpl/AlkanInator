"""Lokalna aplikacja "Wgraj zmiany do sklepu" - w przegladarce, bez terminala.

    py src/woo_preview_app.py            (albo dwuklik "Wgraj zmiany do sklepu.bat")

Przebieg: plik XLSX -> dry-run w sklepie -> karty produktow przed/po w ukladzie motywu ->
"Zatwierdz" -> zapis tylko tego, co bylo w podgladzie (gdy sklep zmienil sie w miedzyczasie,
zapis staje) -> identyfikator zapisu i "Cofnij". Serwer slucha tylko na 127.0.0.1, a API wymaga
jednorazowego klucza z adresu strony. Zapis i cofniecie na PRODUKCJI zatwierdza czlowiek w okienku
Windows - odpowiednik potwierdzenia w terminalu CLI, ktorego agent AI tez nie wykona.
"""

from __future__ import annotations

import argparse
import json
import mimetypes
import os
import re
import secrets
import threading
import time
import traceback
import unicodedata
import webbrowser
from collections import Counter
from dataclasses import dataclass, field
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any, Callable
from urllib.parse import parse_qs, unquote, urlparse

from woo_api.client import WooBridgeClient, WooBridgeError
from woo_api.fingerprint import read_fingerprints
from woo_api.push import PushOptions, fetch_image, merge_results, run_push, run_revert
from woo_api.workbook import ROW_NUMBER_KEY, read_products_xlsx, write_products_xlsx
from woo_api_cli import DEFAULT_CONFIG_PATH, DEFAULT_REPORT_DIR, ROOT, load_connection
from woo_preview.cards import (
    PreviewContext,
    ProductView,
    VariantGroup,
    build_context,
    build_views,
    change_html,
    changes_signature,
    SHIPPING_COLUMNS,
    grouped_tags,
    groups_from_api,
    load_variant_groups,
    related_refs,
    render_card,
)


STATIC_DIR = Path(__file__).resolve().parent / "static"
WORK_DIR = ROOT / "cache" / "woo_preview"
# Grupy WPC Linked Variation podaje wtyczka bridge 2.1 (/variant-groups); dla starszej zostaje zrzut z AlkanLocal.
VARIANT_GROUPS_PATH = ROOT / "dictionaries" / "alkanlocal_wpclv_groups.tsv"
RELATED_BATCH = 200
TAG_BATCH = 50
DEFAULT_THEME_DIR = Path(r"C:\laragon\www\AlkanLocal\wp-content\themes\alkan")
THEME_TYPES = {".css", ".woff", ".woff2", ".ttf", ".eot", ".svg", ".png", ".jpg", ".jpeg", ".gif", ".webp"}
MAX_UPLOAD = 100 * 1024 * 1024
# Strona pinguje co 30 s; po zamknieciu karty serwer konczy sie sam (chyba ze trwa zapis).
IDLE_SHUTDOWN_SECONDS = 15 * 60
STALE_PREVIEW = (
    "Sklep zmienił się od podglądu - zmiany do zapisu nie są już takie same jak na ekranie. "
    "Kliknij „Odśwież podgląd” i sprawdź je jeszcze raz."
)


class BusyError(Exception):
    pass


def plural(count: int, one: str, few: str, many: str) -> str:
    if count == 1:
        return one
    if count % 10 in (2, 3, 4) and count % 100 not in (12, 13, 14):
        return few
    return many


def safe_file_name(name: str) -> str:
    stem = Path(name or "plik.xlsx").name
    stem = stem.translate(str.maketrans("łŁ", "lL"))
    stem = unicodedata.normalize("NFKD", stem).encode("ascii", "ignore").decode("ascii")
    stem = re.sub(r"[^A-Za-z0-9._-]+", "_", stem).strip("._") or "plik"
    return stem if stem.lower().endswith(".xlsx") else stem + ".xlsx"


def list_profiles(config_path: Path = DEFAULT_CONFIG_PATH) -> list[dict[str, Any]]:
    """Profile sklepow z config.json - tylko nazwa, host i flaga produkcji, nigdy token."""
    try:
        config = json.loads(config_path.read_text(encoding="utf-8-sig"))
    except (OSError, json.JSONDecodeError):
        return []
    if isinstance(config.get("profiles"), dict):
        profiles, default = config["profiles"], str(config.get("default", "local"))
    else:
        profiles, default = {"local": config}, "local"
    return [
        {
            "name": name,
            "host": urlparse(str(item.get("url", ""))).netloc,
            "production": bool(item.get("production", name == "prod")),
            "default": name == default,
        }
        for name, item in profiles.items()
        if isinstance(item, dict)
    ]


def windows_confirm(text: str, title: str) -> bool:
    """Okienko Tak/Nie na pulpicie; domyslnie zaznaczone "Nie". Poza Windows zawsze odmowa."""
    if os.name != "nt":
        return False
    import ctypes

    flags = 0x4 | 0x30 | 0x100 | 0x1000 | 0x10000 | 0x40000  # YESNO, WARNING, DEFBUTTON2, SYSTEMMODAL, SETFOREGROUND, TOPMOST
    return ctypes.windll.user32.MessageBoxW(None, text, title, flags) == 6  # IDYES


class RecordingClient:
    """Przepuszcza wywolania do klienta sklepu i zapamietuje wyniki dry-runu, zeby porownac je z podgladem."""

    def __init__(self, client: Any) -> None:
        self.client = client
        self.dry_parts: list[dict[str, Any]] = []

    def sync(self, rows: list[dict[str, Any]], **kwargs: Any) -> dict[str, Any]:
        result = self.client.sync(rows, **kwargs)
        if kwargs.get("dry_run", True):
            self.dry_parts.append(result)
        return result

    def __getattr__(self, name: str) -> Any:
        return getattr(self.client, name)


@dataclass
class Job:
    kind: str = ""
    state: str = "idle"  # idle | running | done | error
    log: list[str] = field(default_factory=list)
    message: str = ""


@dataclass
class Session:
    profile: str
    production: bool
    host: str
    file_name: str
    upload: Path
    columns: list[str]
    all_columns: list[str]
    headers: list[str]
    rows: list[dict[str, Any]]
    dry: dict[str, Any]
    views: list[ProductView]
    products_json: list[dict[str, Any]]
    signature: frozenset[tuple[str, str, str, str]]
    reports: list[str]
    context: PreviewContext | None = None
    related_fetched: int = 0
    variant_groups: dict[str, Any] = field(default_factory=dict)
    # Metody dostawy liczone przy otwarciu karty (indeks produktu -> wynik /shipping-preview).
    shipping: dict[int, dict[str, Any]] = field(default_factory=dict)
    shipping_note: str = ""
    all_rows: bool = False
    version: int = 0
    changeset: str = ""
    applied: dict[str, Any] | None = None
    reverted: dict[str, Any] | None = None

    @property
    def to_apply(self) -> int:
        return sum(view.status == "changed" for view in self.views)


def product_json(view: ProductView) -> dict[str, Any]:
    return {
        "i": view.index,
        "id": view.product_id,
        "sku": view.sku,
        "title": view.title,
        "status": view.status,
        "columns": [change["column"] for change in view.changes],
        "changes": [{"column": change["column"], "html": change_html(change["column"], change["old"], change["new"])} for change in view.changes],
        "errors": view.errors,
        "warnings": view.warnings,
    }


def file_shop_host(rows: list[dict[str, Any]]) -> str:
    """Sklep, z ktorego pobrano plik (najczestszy host w kolumnie Permalink); pusty, gdy kolumny brak."""
    hosts = Counter(urlparse(str(row.get("Permalink") or "")).hostname or "" for row in rows)
    hosts.pop("", None)
    return hosts.most_common(1)[0][0] if hosts else ""


def _batches(items: list[str], size: int) -> list[list[str]]:
    return [items[start : start + size] for start in range(0, len(items), size)]


def _group_count(groups: dict[str, list[VariantGroup]]) -> int:
    return len({id(group) for items in groups.values() for group in items})


def _problems(result: dict[str, Any] | None) -> list[str]:
    if not result:
        return []
    return [
        f"{product.get('sku') or product.get('id')}: {error}"
        for product in result.get("products", [])
        for error in product.get("errors", [])
    ]


class App:
    def __init__(
        self,
        *,
        theme_dir: Path | None = DEFAULT_THEME_DIR,
        config_path: Path = DEFAULT_CONFIG_PATH,
        work_dir: Path = WORK_DIR,
        report_dir: Path = DEFAULT_REPORT_DIR,
        variant_groups_path: Path = VARIANT_GROUPS_PATH,
        client_factory: Callable[[str, Callable[[str], None]], tuple[Any, dict[str, Any]]] | None = None,
        confirm_production: Callable[[str, str], bool] = windows_confirm,
        now: Callable[[], datetime] = datetime.now,
    ) -> None:
        self.key = secrets.token_urlsafe(24)
        self.port = 0
        self.lock = threading.Lock()
        self.shipping_lock = threading.Lock()
        self.job = Job()
        self.session: Session | None = None
        self.theme_dir = theme_dir if theme_dir and (theme_dir / "style.css").is_file() else None
        self.config_path = config_path
        self.work_dir = work_dir
        self.report_dir = report_dir
        self.variant_groups = load_variant_groups(variant_groups_path)
        self.variant_groups_info = {
            "count": _group_count(self.variant_groups),
            "file": variant_groups_path.name,
            "date": datetime.fromtimestamp(variant_groups_path.stat().st_mtime).strftime("%Y-%m-%d") if variant_groups_path.is_file() else "",
        }
        self.client_factory = client_factory or self._connect
        self.confirm_production = confirm_production
        self.now = now
        self.last_seen = time.monotonic()
        self._threads: list[threading.Thread] = []

    # --- polaczenie i zadania w tle -------------------------------------------------------

    def _connect(self, profile: str, log: Callable[[str], None]) -> tuple[Any, dict[str, Any]]:
        args = argparse.Namespace(config=str(self.config_path), profile=profile, url="", insecure=False)
        connection = load_connection(args)
        client = WooBridgeClient(
            str(connection["url"]), str(connection["token"]), verify_ssl=bool(connection["verify_ssl"]), log=log
        )
        return client, connection

    def log(self, message: str) -> None:
        with self.lock:
            self.job.log.append(message)
            del self.job.log[:-500]

    def _start(self, kind: str, target: Callable[[], None]) -> None:
        with self.lock:
            if self.job.state == "running":
                raise BusyError("Trwa poprzednie zadanie - poczekaj, aż się skończy.")
            self.job = Job(kind=kind, state="running")

        def run() -> None:
            try:
                target()
                with self.lock:
                    if self.job.state == "running":
                        self.job.state = "done"
            except (WooBridgeError, ValueError, OSError) as exc:
                with self.lock:
                    self.job.state, self.job.message = "error", str(exc)
            except Exception as exc:  # noqa: BLE001 - blad ma trafic na ekran, a nie zabic watek po cichu
                write_app_log(self.work_dir, traceback.format_exc())
                with self.lock:
                    self.job.state, self.job.message = "error", f"Nieoczekiwany błąd: {exc}"

        thread = threading.Thread(target=run, name=f"woo-preview-{kind}", daemon=True)
        self._threads.append(thread)
        thread.start()

    def _fail(self, message: str) -> None:
        with self.lock:
            self.job.state, self.job.message = "error", message

    def wait(self, timeout: float = 30.0) -> None:
        """Czeka na zadania w tle (testy)."""
        for thread in list(self._threads):
            thread.join(timeout)

    # --- podglad --------------------------------------------------------------------------

    def start_preview(self, profile: str, file_name: str, data: bytes, all_rows: bool = False) -> None:
        if not data.startswith(b"PK"):
            raise ValueError("To nie jest plik XLSX - wybierz plik z eksportu sklepu.")
        if profile not in {item["name"] for item in list_profiles(self.config_path)}:
            raise ValueError(f"Nie ma sklepu „{profile}” w konfiguracji.")
        folder = self.work_dir / self.now().strftime("%Y%m%d-%H%M%S-%f")
        folder.mkdir(parents=True, exist_ok=True)
        # Kopia pliku - oryginal uzytkownika nigdy nie jest zmieniany.
        upload = folder / safe_file_name(file_name)
        upload.write_bytes(data)
        self._start("preview", lambda: self._run_preview(profile, upload, Path(file_name).name or upload.name, [], None, all_rows))

    def start_recalc(self, columns: list[str]) -> None:
        session = self._require_session()
        if session.applied is not None:
            raise ValueError("Ten plik jest już zapisany - wczytaj go ponownie, żeby zrobić nowy podgląd.")
        unknown = [column for column in columns if column not in session.all_columns]
        if unknown:
            raise ValueError("Nieznane kolumny: " + ", ".join(unknown))
        chosen = [] if set(columns) == set(session.all_columns) else columns
        self._start(
            "preview",
            lambda: self._run_preview(session.profile, session.upload, session.file_name, chosen, session.all_columns, session.all_rows),
        )

    def _run_preview(
        self, profile: str, upload: Path, file_name: str, columns: list[str], all_columns: list[str] | None, all_rows: bool = False
    ) -> None:
        client, connection = self.client_factory(profile, self.log)
        production = bool(connection.get("production"))
        host = urlparse(str(connection.get("url", ""))).netloc
        self.log(f"Sklep: {host} ({'PRODUKCJA' if production else 'profil ' + profile})")
        self.log(f"Plik: {file_name}")
        headers, rows = read_products_xlsx(upload)
        # Eksport z innego sklepu ma inne ID i adresy zdjec: dry-run pokazalby zmiane prawie kazdego
        # produktu i pobieral tysiace zdjec, a zapis podmienilby je w calym sklepie.
        prints = read_fingerprints(upload)
        source_host = file_shop_host(rows) or (urlparse("//" + prints.site).hostname or "" if prints else "")
        shop_host = urlparse(str(connection.get("url", ""))).hostname or ""
        if source_host and shop_host and source_host.lower() != shop_host.lower():
            raise ValueError(
                f"Ten plik pobrano ze sklepu {source_host} (kolumna Permalink), a wybrany sklep to {shop_host}. "
                "Wybierz właściwy sklep albo pobierz plik z tego sklepu."
            )
        with client:
            result = run_push(
                client,
                upload,
                PushOptions(apply=False, columns=columns, all_rows=all_rows),
                report_dir=self.report_dir,
                log=self.log,
                now=self.now(),
                image_fetcher=self._image_fetcher(client),
            )
            views = build_views(rows, result.dry_run)
            groups, groups_info = self._load_groups(client, host)
            extra_rows = self._fetch_related(client, rows, views, groups if groups_info["live"] else None)
        if all_columns is None:
            changed = {column for view in views for column in view.changed_columns}
            all_columns = [column for column in headers if column in changed]
        session = Session(
            profile=profile,
            production=production,
            host=host,
            file_name=file_name,
            upload=upload,
            columns=columns or list(all_columns),
            all_columns=list(all_columns),
            headers=headers,
            rows=rows,
            dry=result.dry_run,
            views=views,
            products_json=[product_json(view) for view in views],
            signature=changes_signature(result.dry_run),
            reports=[str(path.resolve()) for path in result.report_paths],
            version=(self.session.version + 1) if self.session else 1,
            context=build_context(rows, views, extra_rows, groups),
            related_fetched=len(extra_rows),
            variant_groups=groups_info,
            all_rows=all_rows,
        )
        with self.lock:
            self.session = session
        self.log(f"Podgląd gotowy: {session.to_apply} {plural(session.to_apply, 'produkt', 'produkty', 'produktów')} do zapisu.")

    def _image_fetcher(self, client: Any) -> Callable[[str], bytes]:
        """Pobieranie nowych zdjec do kontroli JPG 1000x1000 - z postepem, bo run_push loguje dopiero po wszystkich."""
        verify = getattr(getattr(client, "session", None), "verify", True)
        count = [0]

        def fetch(url: str) -> bytes:
            count[0] += 1
            if count[0] == 1 or count[0] % 10 == 0:
                self.log(f"Sprawdzam nowe zdjęcia (JPG 1000x1000 na białym tle): {count[0]}…")
            return fetch_image(url, verify=verify)

        return fetch

    def _load_groups(self, client: Any, host: str) -> tuple[dict[str, list[VariantGroup]], dict[str, Any]]:
        """Grupy wariantow prosto ze sklepu; wtyczka sprzed 2.1 (404) -> zrzut z repo."""
        if hasattr(client, "variant_groups"):
            try:
                groups = groups_from_api(client.variant_groups())
                return groups, {"live": True, "source": f"sklep {host}", "count": _group_count(groups)}
            except WooBridgeError as exc:
                if exc.status_code != 404:
                    raise
                self.log("Sklep nie podaje grup wariantów (wtyczka bridge sprzed 2.1) - używam zrzutu z repo.")
        info = self.variant_groups_info
        return self.variant_groups, {"live": False, "source": f"zrzut {info['file']} z dnia {info['date']}", "count": info["count"]}

    def _fetch_related(
        self, client: Any, rows: list[dict[str, Any]], views: list[ProductView], live_groups: dict[str, list[VariantGroup]] | None
    ) -> list[dict[str, Any]]:
        """Produkty spoza pliku potrzebne kartom zmienionych produktow, tylko odczyt.

        Up-Sells/Cross-Sells/kompatybilne: jeden odczyt na 200 SKU. Warianty: produkty z tagami grup
        wariantow, jeden odczyt na 50 tagow - tylko gdy sklep zna filtr tag (wtyczka 2.1; starsza
        pominelaby go i oddala caly katalog).
        """
        if not hasattr(client, "iter_product_pages"):
            return []
        known = {str(row.get("SKU", "")) for row in rows} | {str(row.get("id", "")) for row in rows}
        missing = sorted(ref for ref in related_refs(views) if ref not in known)
        queries: list[tuple[str, list[str], int]] = []
        queries += [("sku", batch, RELATED_BATCH) for batch in _batches([ref for ref in missing if not ref.isdigit()], RELATED_BATCH)]
        queries += [("include", batch, RELATED_BATCH) for batch in _batches([ref for ref in missing if ref.isdigit()], RELATED_BATCH)]
        tags = sorted(grouped_tags(views, live_groups)) if live_groups is not None else []
        queries += [("tag", batch, RELATED_BATCH) for batch in _batches(tags, TAG_BATCH)]
        if not queries:
            return []
        if missing:
            self.log(f"Pobieram {len(missing)} {plural(len(missing), 'powiązany produkt', 'powiązane produkty', 'powiązanych produktów')} spoza pliku (tylko odczyt)…")
        if tags:
            self.log(f"Pobieram warianty z {len(tags)} {plural(len(tags), 'tagu', 'tagów', 'tagów')} grup wariantów (tylko odczyt)…")
        found: list[dict[str, Any]] = []
        try:
            for key, batch, per_page in queries:
                for page in client.iter_product_pages(per_page=per_page, **{key: batch}):
                    for row in page.get("rows", []):
                        if str(row.get("SKU", "")) not in known and str(row.get("id", "")) not in known:
                            known.add(str(row.get("SKU", "")))
                            found.append(row)
        except WooBridgeError as exc:
            # Podglad dziala dalej - brakujace kafelki pokaza tylko SKU.
            self.log(f"Nie udało się pobrać powiązanych produktów: {exc}")
        return found

    # --- zapis i cofniecie ----------------------------------------------------------------

    def _require_session(self) -> Session:
        if self.session is None:
            raise ValueError("Najpierw wczytaj plik i obejrzyj podgląd.")
        return self.session

    def start_apply(self) -> None:
        session = self._require_session()
        if session.applied is not None:
            raise ValueError("Te zmiany są już zapisane.")
        if session.dry.get("summary", {}).get("errors"):
            raise ValueError("Plik ma błędy - popraw go i wczytaj ponownie. Zapis jest wstrzymany.")
        if not session.to_apply:
            raise ValueError("Nie ma nic do zapisania - sklep ma już te wartości.")
        self._start("apply", lambda: self._run_apply(session))

    def _run_apply(self, session: Session) -> None:
        # Do zapisu idzie plik z samymi zmienionymi produktami: dry-run przed zapisem nie mieli
        # calego eksportu jeszcze raz (mniej pracy dla sklepu), a wynik jest ten sam.
        changed_rows = {view.row for view in session.views if view.status == "changed"}
        target = session.upload.parent / "do-zapisu" / session.upload.name
        write_products_xlsx(target, session.headers, [row for row in session.rows if row[ROW_NUMBER_KEY] in changed_rows])

        client, connection = self.client_factory(session.profile, self.log)
        recorder = RecordingClient(client)
        stale = []

        def confirm(count: int) -> bool:
            if changes_signature(merge_results(recorder.dry_parts)) != session.signature:
                stale.append(True)
                return False
            if not session.production:
                return True
            columns = ", ".join(session.columns) or "wszystkie zmienione"
            return self.confirm_production(
                f"Zapisać zmiany w {count} {plural(count, 'produkcie', 'produktach', 'produktach')} na PRODUKCJI?\n\n"
                f"Sklep: {session.host}\nPlik: {session.file_name}\nKolumny: {columns}\n\n"
                "Zapis można później cofnąć przyciskiem „Cofnij ten zapis”.",
                "AlkanInator - zapis na produkcję",
            )

        with client:
            result = run_push(
                recorder,
                target,
                PushOptions(apply=True, columns=[] if session.columns == session.all_columns else session.columns, confirm=confirm),
                report_dir=self.report_dir,
                log=self.log,
                now=self.now(),
            )
        if result.applied is None:
            if stale:
                self._fail(STALE_PREVIEW)
            elif result.blocked:
                last = next((line for line in reversed(self.job.log) if "wstrzymany" in line.lower()), "Zapis wstrzymany.")
                self._fail(last if "brak potwierdzenia" not in last else "Zapis anulowany - nie potwierdzono w okienku.")
            else:
                self._fail("Brak zmian do zapisania - sklep ma już te wartości.")
            return
        with self.lock:
            session.applied = result.applied
            session.changeset = result.changeset
            session.reports = session.reports + [str(path.resolve()) for path in result.report_paths[1:]]
            session.version += 1
        applied = int(result.applied.get("summary", {}).get("applied", 0))
        self.log(f"Zapisano: {applied} {plural(applied, 'produkt', 'produkty', 'produktów')}, zapis {result.changeset}.")

    def start_revert(self) -> None:
        session = self._require_session()
        if not session.changeset:
            raise ValueError("Nie ma zapisu do cofnięcia.")
        if session.reverted is not None:
            raise ValueError("Ten zapis jest już cofnięty.")
        self._start("revert", lambda: self._run_revert(session))

    def _run_revert(self, session: Session) -> None:
        client, _ = self.client_factory(session.profile, self.log)

        def confirm(count: int) -> bool:
            return self.confirm_production(
                f"Cofnąć zapis {session.changeset} na PRODUKCJI?\n\n"
                f"Sklep: {session.host}\nProdukty do przywrócenia: {count}",
                "AlkanInator - cofnięcie na produkcji",
            )

        self.log(f"Cofam zapis {session.changeset}…")
        with client:
            try:
                result, report = run_revert(
                    client,
                    session.changeset,
                    apply=True,
                    report_dir=self.report_dir,
                    now=self.now(),
                    confirm=confirm if session.production else None,
                )
            except ValueError as exc:
                if "brak potwierdzenia" in str(exc):
                    raise ValueError("Cofnięcie anulowane - nie potwierdzono w okienku.") from exc
                raise
        with self.lock:
            session.reverted = {"summary": result.get("summary", {}), "errors": _problems(result), "report": str(report.resolve())}
            session.version += 1
        self.log("Cofnięto zapis.")

    def reset(self) -> None:
        with self.lock:
            if self.job.state == "running":
                raise BusyError("Trwa zadanie - poczekaj, aż się skończy.")
            self.session = None
            self.job = Job()

    # --- odpowiedzi API -------------------------------------------------------------------

    def state(self) -> dict[str, Any]:
        with self.lock:
            job = {"kind": self.job.kind, "state": self.job.state, "log": list(self.job.log[-300:]), "message": self.job.message}
            session = self.session
        data: dict[str, Any] = {"profiles": list_profiles(self.config_path), "theme": self.theme_dir is not None, "job": job, "session": None}
        if session is not None:
            dry = session.dry
            data["session"] = {
                "version": session.version,
                "file": session.file_name,
                "profile": session.profile,
                "production": session.production,
                "host": session.host,
                "columns": session.columns,
                "all_columns": session.all_columns,
                "summary": dry.get("summary", {}),
                "to_apply": session.to_apply,
                "writes_enabled": dry.get("writes_enabled", True),
                "max_products": int(dry.get("max_products") or 0),
                "new_terms": dry.get("new_terms", []),
                "ignored_columns": dry.get("ignored_columns", []),
                "reports": session.reports,
                "related_fetched": session.related_fetched,
                "variant_groups": session.variant_groups,
                "changeset": session.changeset,
                "applied": None
                if session.applied is None
                else {"summary": session.applied.get("summary", {}), "errors": _problems(session.applied)},
                "reverted": session.reverted,
            }
        return data

    def products(self) -> dict[str, Any]:
        session = self.session
        return {"version": session.version if session else 0, "products": session.products_json if session else []}

    def card(self, index: int, side: str) -> str:
        session = self._require_session()
        if not 0 <= index < len(session.views):
            raise ValueError("Nie ma takiego produktu w podglądzie.")
        theme_base = "/theme" if self.theme_dir is not None else None
        view = session.views[index]
        shipping, note = self._shipping_for(session, view)
        return render_card(
            view,
            "before" if side == "before" else "after",
            theme_base=theme_base,
            context=session.context,
            shipping=shipping,
            shipping_note=note,
        )

    def _shipping_for(self, session: Session, view: ProductView) -> tuple[dict[str, Any] | None, str]:
        """Metody dostawy teraz i po zmianie, liczone przez sklep przy pierwszym otwarciu karty (tylko odczyt)."""
        if view.product_id is None:
            return None, "Plik bez kolumny id - sklep nie policzy metod dostawy."
        # Obie ramki (przed/po) pytaja naraz; blokada sprawia, ze sklep liczy raz.
        with self.shipping_lock:
            if session.shipping_note:
                return None, session.shipping_note
            if view.index in session.shipping:
                item = session.shipping[view.index]
                return (None, item["error"]) if item.get("error") else (item, "")
            values = {change["column"]: change["new"] for change in view.changes if change["column"] in SHIPPING_COLUMNS}
            client, _ = self.client_factory(session.profile, self.log)
            if not hasattr(client, "shipping_preview"):
                return None, "Ten sklep nie liczy podglądu wysyłki."
            try:
                with client:
                    result = client.shipping_preview([{"id": view.product_id, "values": values}])
            except WooBridgeError as exc:
                if exc.status_code == 404:
                    session.shipping_note = "Metody dostawy policzy sklep z wtyczką alkan-inator-bridge 2.1 - ten sklep ma starszą."
                    return None, session.shipping_note
                if exc.status_code == 501:
                    session.shipping_note = str(exc)
                    return None, session.shipping_note
                return None, f"Nie udało się policzyć wysyłki: {exc}"
            item = result[0] if result else {"error": "Sklep nie zwrócił wyniku."}
            session.shipping[view.index] = item
            return (None, item["error"]) if item.get("error") else (item, "")

    def touch(self) -> None:
        self.last_seen = time.monotonic()

    def idle(self) -> bool:
        with self.lock:
            running = self.job.state == "running"
        return not running and time.monotonic() - self.last_seen > IDLE_SHUTDOWN_SECONDS


def write_app_log(work_dir: Path, text: str) -> None:
    try:
        work_dir.mkdir(parents=True, exist_ok=True)
        with (work_dir / "app.log").open("a", encoding="utf-8") as handle:
            handle.write(f"[{datetime.now():%Y-%m-%d %H:%M:%S}] {text}\n")
    except OSError:
        pass


class Handler(BaseHTTPRequestHandler):
    server_version = "AlkanInator"
    server: "PreviewServer"

    def log_message(self, format: str, *args: Any) -> None:  # noqa: A002 - podpis z BaseHTTPRequestHandler
        pass  # pyw nie ma konsoli; bledy ida do cache/woo_preview/app.log

    def do_GET(self) -> None:
        self._dispatch("GET")

    def do_POST(self) -> None:
        self._dispatch("POST")

    def _send(self, status: int, content_type: str, body: bytes) -> None:
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.send_header("X-Content-Type-Options", "nosniff")
        self.end_headers()
        self.wfile.write(body)

    def _json(self, data: Any, status: int = 200) -> None:
        self._send(status, "application/json; charset=utf-8", json.dumps(data, ensure_ascii=False).encode("utf-8"))

    def _file(self, base: Path, relative: str, allowed: set[str] | None = None) -> None:
        path = (base / unquote(relative)).resolve()
        if not path.is_relative_to(base.resolve()) or not path.is_file() or (allowed and path.suffix.lower() not in allowed):
            self._send(404, "text/plain; charset=utf-8", b"Nie ma takiego pliku")
            return
        content_type = mimetypes.guess_type(path.name)[0] or "application/octet-stream"
        if content_type.startswith("text/") or content_type == "application/javascript":
            content_type += "; charset=utf-8"
        self._send(200, content_type, path.read_bytes())

    def _body(self) -> bytes:
        length = int(self.headers.get("Content-Length") or 0)
        if length > MAX_UPLOAD:
            raise ValueError("Plik jest za duży (limit 100 MB).")
        return self.rfile.read(length) if length else b""

    def _dispatch(self, method: str) -> None:
        app = self.server.app
        if self.headers.get("Host", "") not in {f"127.0.0.1:{app.port}", f"localhost:{app.port}"}:
            self._send(403, "text/plain; charset=utf-8", b"Zly adres")
            return
        url = urlparse(self.path)
        query = parse_qs(url.query)
        key = self.headers.get("X-Key") or (query.get("k") or [""])[0]
        authorized = secrets.compare_digest(key, app.key)
        app.touch()
        try:
            if method == "GET" and url.path == "/":
                if not authorized:
                    self._send(403, "text/plain; charset=utf-8", "Otwórz aplikację ikoną „Wgraj zmiany do sklepu”.".encode("utf-8"))
                    return
                self._file(STATIC_DIR, "index.html")
                return
            if method == "GET" and url.path.startswith("/static/"):
                self._file(STATIC_DIR, url.path[len("/static/"):], {".js", ".css"})
                return
            if method == "GET" and url.path.startswith("/theme/"):
                if app.theme_dir is None:
                    self._send(404, "text/plain; charset=utf-8", b"Brak motywu")
                    return
                self._file(app.theme_dir, url.path[len("/theme/"):], THEME_TYPES)
                return
            if not authorized:
                self._json({"error": "Brak dostępu - otwórz aplikację od nowa."}, 403)
                return
            self._route(method, url.path, query)
        except BusyError as exc:
            self._json({"error": str(exc)}, 409)
        except (ValueError, WooBridgeError) as exc:
            self._json({"error": str(exc)}, 400)
        except Exception as exc:  # noqa: BLE001
            write_app_log(app.work_dir, traceback.format_exc())
            self._json({"error": f"Nieoczekiwany błąd: {exc}"}, 500)

    def _route(self, method: str, path: str, query: dict[str, list[str]]) -> None:
        app = self.server.app
        if method == "GET" and path == "/api/state":
            self._json(app.state())
        elif method == "GET" and path == "/api/products":
            self._json(app.products())
        elif method == "GET" and path == "/card":
            index = int((query.get("i") or ["-1"])[0])
            body = app.card(index, (query.get("side") or ["after"])[0])
            self._send(200, "text/html; charset=utf-8", body.encode("utf-8"))
        elif method == "POST" and path == "/api/preview":
            data = self._body()
            app.start_preview(
                unquote(self.headers.get("X-Profile", "")),
                unquote(self.headers.get("X-File-Name", "")),
                data,
                all_rows=self.headers.get("X-All-Rows") == "1",
            )
            self._json({"ok": True})
        elif method == "POST" and path == "/api/recalc":
            payload = json.loads(self._body() or b"{}")
            app.start_recalc([str(column) for column in payload.get("columns", [])])
            self._json({"ok": True})
        elif method == "POST" and path == "/api/apply":
            app.start_apply()
            self._json({"ok": True})
        elif method == "POST" and path == "/api/revert":
            app.start_revert()
            self._json({"ok": True})
        elif method == "POST" and path == "/api/reset":
            app.reset()
            self._json({"ok": True})
        elif method == "POST" and path == "/api/ping":
            self._json({"ok": True})
        elif method == "POST" and path == "/api/shutdown":
            if app.job.state == "running":
                raise BusyError("Trwa zadanie - zamknij aplikację, gdy się skończy.")
            self._json({"ok": True})
            threading.Thread(target=self.server.shutdown, daemon=True).start()
        else:
            self._json({"error": "Nieznany adres"}, 404)


class PreviewServer(ThreadingHTTPServer):
    daemon_threads = True

    def __init__(self, app: App, port: int = 0) -> None:
        super().__init__(("127.0.0.1", port), Handler)
        self.app = app
        app.port = self.server_address[1]

    @property
    def url(self) -> str:
        return f"http://127.0.0.1:{self.app.port}/?k={self.app.key}"


def _watchdog(server: PreviewServer) -> None:
    while True:
        time.sleep(30)
        if server.app.idle():
            write_app_log(server.app.work_dir, "koniec - brak aktywnosci strony")
            server.shutdown()
            return


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Podgląd i zapis zmian z pliku XLSX do sklepu (w przeglądarce).")
    parser.add_argument("--port", type=int, default=0, help="Port (domyślnie losowy wolny).")
    parser.add_argument("--theme-dir", default=os.environ.get("ALKAN_THEME_DIR", str(DEFAULT_THEME_DIR)), help="Katalog motywu sklepu (style kart).")
    parser.add_argument("--no-browser", action="store_true", help="Nie otwieraj przeglądarki.")
    args = parser.parse_args(argv)

    app = App(theme_dir=Path(args.theme_dir))
    server = PreviewServer(app, args.port)
    write_app_log(app.work_dir, f"start na porcie {app.port}, motyw: {app.theme_dir or 'brak (style zastepcze)'}")
    threading.Thread(target=_watchdog, args=(server,), daemon=True).start()
    if not args.no_browser:
        webbrowser.open(server.url)
    print(f"Aplikacja dziala: {server.url}")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()
    return 0
