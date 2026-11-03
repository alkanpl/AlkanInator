"""Karty produktu przed/po zmianie w ukladzie motywu sklepu (alkan) i opis zmian dla podgladu.

Stan "przed" nie jest pobierany ze sklepu: to wiersz pliku z polami zmienionymi przez dry-run
cofnietymi do starej wartosci ("było"). Sklep dostaje wiec tylko zapytania samego dry-runu.
"""

from __future__ import annotations

import difflib
import html
import re
import unicodedata
from collections import defaultdict
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from woo_api.workbook import ROW_NUMBER_KEY


ATTRIBUTE_PREFIX = "Atrybut Produktu: "
LIST_SEPARATOR = "|"
# Kolumny z lista wartosci po "|" - zmiane pokazujemy jako dodane/usuniete pozycje.
LIST_COLUMNS = {
    "Kategorie produktów",
    "Product Tags",
    "Up-Sells",
    "Cross-Sells",
    "_alkan_compatible_product_skus",
    "_alkan_compatible_product_ids",
    "Image URL",
}
HTML_COLUMNS = {"Content", "Short Description"}
# Od tych kolumn sklep liczy metody dostawy (Flexible Shipping po klasie wysylkowej i wadze).
SHIPPING_COLUMNS = ("Shipping Class", "Weight", "Length", "Width", "Height")
AVAILABILITY_COLUMNS = {"Backorders", "Wysyłka w ciągu"}
COMPATIBLE_COLUMNS = ("_alkan_compatible_product_skus", "_alkan_compatible_product_ids")
# Kolejnosc arkuszy stylow jak w motywie (functions.php + head.php).
THEME_STYLESHEETS = (
    "fontawesome-free-5.8.2-web/css/fa.min.css",
    "css/woocommerce-layout.css",
    "css/woocommerce.css",
    "style.css",
)
FONTS_URL = "https://fonts.googleapis.com/css?family=Open+Sans:400,600,700|Raleway:400,600,800&subset=latin-ext"
BLOCK_TAG = re.compile(r"^\s*<(p|h[1-6]|ul|ol|li|div|table|blockquote|hr|figure|section|script|style|pre)\b", re.I)

# Poprawki samego podgladu: zakladki motywu przelacza JS, ktory w podgladzie nie dziala.
# Karta ma <meta name="referrer" content="no-referrer">: alkan.pl blokuje zdjecia (403) przy obcym Referer.
CARD_CSS = """
body { margin: 0; padding: 18px 18px 40px; background: #fff; }
.single-product .woocommerce-tabs .product-tab-content { display: block !important; }
.product-tabs-navigation { display: none; }
/* Galerie --with-images odslania dopiero JS suwaka zdjec WooCommerce. */
.woocommerce-product-gallery { opacity: 1 !important; }
.woocommerce-product-gallery img { max-width: 100%; height: auto; }
#thumbnails-list { display: flex; gap: 8px; flex-wrap: wrap; margin-top: 8px; }
#thumbnails-list img { width: 72px; height: 72px; object-fit: contain; border: 1px solid #dce5f3; }
.ai-changed { outline: 3px solid #f59e0b; outline-offset: 4px; border-radius: 2px; }
tr.ai-changed { outline-offset: -2px; background: #fff7e6; }
.ai-crumbs { font: 13px 'Open Sans', Arial, sans-serif; color: #6f7c91; margin: 0 0 10px; }
.ai-empty { color: #6f7c91; font-style: italic; }
.ai-cleared { color: #c62b11; font-style: italic; }
.ai-note { font: 12px/1.5 'Open Sans', Arial, sans-serif; color: #44516a; background: #eef3fb; border-left: 3px solid #133d96; padding: 6px 10px; margin: 8px 0; border-radius: 3px; }
.ai-note.warn { background: #fff7e6; border-color: #f59e0b; }
.ai-note.bad { background: #fdf0ee; border-color: #c62b11; color: #8a1d0a; }
.ai-section { margin-top: 28px; }
.ai-section .product-suggestions__grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(150px, 1fr)); gap: 12px; }
.ai-section .suggestion-card { border: 1px solid #dce5f3; border-radius: 8px; padding: 8px; background: #fff; }
.ai-section .suggestion-card img { width: 100%; height: 130px; object-fit: contain; }
.ai-section .suggestion-card .product-name { font: 600 13px/1.35 'Open Sans', Arial, sans-serif; margin: 6px 0 0; }
.ai-section .suggestion-card .ai-sku { font: 11px 'Open Sans', Arial, sans-serif; color: #6f7c91; margin: 4px 0 0; }
.ai-section .suggestion-card.ai-missing { border-style: dashed; color: #6f7c91; }
.ai-section .suggestion-card.ai-added { border: 2px solid #54af00; }
.ai-section .suggestion-card.ai-removed { border: 2px dashed #c62b11; opacity: .6; }
.wpclv-attributes { margin: 0 0 12px; }
.wpclv-attribute { margin-bottom: 10px; }
.wpclv-attribute-label { font: 600 13px 'Open Sans', Arial, sans-serif; margin-bottom: 6px; }
.wpclv-terms { display: flex; flex-wrap: wrap; gap: 8px; }
.wpclv-term { border: 1px solid #d8e1ef; border-radius: 6px; background: #fff; font: 13px 'Open Sans', Arial, sans-serif; }
.wpclv-term.active { border: 2px solid #133d96; }
.wpclv-term > span, .wpclv-term > a { display: block; padding: 4px 10px; color: #000d24; text-decoration: none; }
.wpclv-term-image > span, .wpclv-term-image > a { padding: 4px; text-align: center; }
.wpclv-term-image-frame img { width: 56px; height: 56px; object-fit: contain; display: block; margin: 0 auto; }
.wpclv-term-image-label { display: block; font-size: 11px; margin-top: 2px; }
.wpclv-term.ai-dup { border-color: #c62b11; }
.ai-availability { font: 600 14px 'Open Sans', Arial, sans-serif; margin: 0 0 4px; }
.ai-availability.ok { color: #3d7f00; }
.ai-availability.order { color: #9a6200; }
.ai-availability.none { color: #c62b11; }
.ai-ship { border: 1px solid #d8e1ef; border-radius: 8px; padding: 12px; margin: 0 0 12px; background: #fff; }
.ai-ship .box-title { font: 600 14px 'Open Sans', Arial, sans-serif; margin: 0 0 8px; }
.ai-ship .ai-ship-class { font: 12px 'Open Sans', Arial, sans-serif; color: #44516a; margin: 0 0 8px; }
.ai-ship-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(120px, 1fr)); gap: 8px; }
.ai-ship-method { border: 1px solid var(--ship-border, #d8e1ef); border-radius: 6px; padding: 6px 8px; font: 12px/1.35 'Open Sans', Arial, sans-serif; display: flex; flex-direction: column; gap: 2px; }
.ai-ship-method img { height: 24px; max-width: 80px; object-fit: contain; align-self: flex-start; }
.ai-ship-method strong { font-size: 13px; }
.ai-ship-method.ai-added { border: 2px solid #54af00; }
.ai-ship-method.ai-removed { border: 2px dashed #c62b11; opacity: .6; }
"""

# Gdy nie ma plikow motywu (inny komputer), karta dostaje proste style o tym samym ukladzie.
FALLBACK_CSS = """
body { font: 14px/1.6 'Open Sans', Arial, sans-serif; color: #000d24; }
.page-name.product-title { font: 800 24px/1.3 Raleway, Arial, sans-serif; margin: 0 0 10px; }
.product-purchase-meta { display: flex; gap: 8px; flex-wrap: wrap; margin-bottom: 14px; }
.product-spec-badge { margin: 0; padding: 3px 10px; background: #dce5f3; border-radius: 12px; font-size: 13px; }
.product-main-grid { display: grid; grid-template-columns: minmax(0, 1fr) minmax(0, 1fr); gap: 24px; }
.product-single-box { border: 1px solid #d8e1ef; border-radius: 6px; padding: 14px; margin-bottom: 12px; }
.product-price { font-size: 26px; font-weight: 700; color: #133d96; margin: 0; }
.woocommerce-tabs { margin-top: 28px; }
.product-tab-content h2 { font: 800 20px Raleway, Arial, sans-serif; color: #133d96; }
table.shop_attributes { border-collapse: collapse; width: 100%; }
table.shop_attributes th, table.shop_attributes td { border-bottom: 1px solid #dce5f3; padding: 6px 8px; text-align: left; vertical-align: top; }
table.shop_attributes th { width: 40%; color: #6f7c91; font-weight: 600; }
table.shop_attributes td p { margin: 0; }
@media (max-width: 700px) { .product-main-grid { grid-template-columns: 1fr; } }
"""


@dataclass
class ProductView:
    index: int
    row: int
    product_id: int | None
    sku: str
    title: str
    status: str
    changes: list[dict[str, str]] = field(default_factory=list)
    errors: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    before: dict[str, str] = field(default_factory=dict)
    after: dict[str, str] = field(default_factory=dict)

    @property
    def changed_columns(self) -> set[str]:
        return {change["column"] for change in self.changes}


def _text(value: Any) -> str:
    return "" if value is None else str(value)


def split_list(value: str) -> list[str]:
    return [part.strip() for part in _text(value).split(LIST_SEPARATOR) if part.strip()]


def product_key(product: dict[str, Any]) -> str:
    return f"id:{product['id']}" if product.get("id") else f"sku:{product.get('sku', '')}"


def build_views(rows: list[dict[str, Any]], dry: dict[str, Any]) -> list[ProductView]:
    """Produkty z dry-runu, ktore cos zmieniaja albo maja bledy/ostrzezenia; bledy na poczatku."""
    by_row = {row[ROW_NUMBER_KEY]: row for row in rows}
    views: list[ProductView] = []
    for product in dry.get("products", []):
        if not (product.get("changes") or product.get("errors") or product.get("warnings")):
            continue
        source = by_row.get(product.get("row"), {})
        base = {key: _text(value) for key, value in source.items() if key != ROW_NUMBER_KEY}
        before, after = dict(base), dict(base)
        changes = []
        for change in product.get("changes", []):
            column, old, new = change["column"], _text(change.get("old")), _text(change.get("new"))
            before[column], after[column] = old, new
            changes.append({"column": column, "old": old, "new": new})
        views.append(
            ProductView(
                index=0,
                row=int(product.get("row") or 0),
                product_id=int(product["id"]) if product.get("id") else None,
                sku=_text(product.get("sku") or base.get("SKU")),
                title=after.get("Title") or _text(product.get("title")),
                status=_text(product.get("status")),
                changes=changes,
                errors=list(product.get("errors", [])),
                warnings=list(product.get("warnings", [])),
                before=before,
                after=after,
            )
        )
    views.sort(key=lambda view: (not view.errors, not view.changes, view.row))
    for index, view in enumerate(views):
        view.index = index
    return views


def changes_signature(dry: dict[str, Any]) -> frozenset[tuple[str, str, str, str]]:
    """Co dokladnie zapisze sklep: (produkt, kolumna, stara, nowa wartosc) dla produktow ze zmianami.

    Stara wartosc tez sie liczy: gdy ktos zmienil produkt po podgladzie, zapis nadpisalby
    wartosc, ktorej uzytkownik nie widzial.
    """
    return frozenset(
        (product_key(product), change["column"], _text(change.get("old")), _text(change.get("new")))
        for product in dry.get("products", [])
        if product.get("status") == "changed"
        for change in product.get("changes", [])
    )


def strip_tags(value: str) -> str:
    text = re.sub(r"<script\b.*?</script>", " ", value, flags=re.S | re.I)
    text = re.sub(r"<[^>]+>", " ", text)
    return re.sub(r"\s+", " ", html.unescape(text)).strip()


def word_diff(old: str, new: str) -> str:
    """Roznica slowo po slowie jako HTML z <del>/<ins>."""
    old_words, new_words = re.split(r"(\s+)", old), re.split(r"(\s+)", new)
    parts = []
    for tag, i1, i2, j1, j2 in difflib.SequenceMatcher(None, old_words, new_words, autojunk=False).get_opcodes():
        if tag == "equal":
            parts.append(html.escape("".join(old_words[i1:i2])))
            continue
        if i2 > i1:
            parts.append(f"<del>{html.escape(''.join(old_words[i1:i2]))}</del>")
        if j2 > j1:
            parts.append(f"<ins>{html.escape(''.join(new_words[j1:j2]))}</ins>")
    return "".join(parts)


def change_html(column: str, old: str, new: str) -> str:
    """Krotki, czytelny opis jednej zmiany do tabeli "Zmiany"."""
    if not new.strip():
        return '<span class="cleared">pole zostanie wyczyszczone</span>'
    if column in LIST_COLUMNS or (column.startswith(ATTRIBUTE_PREFIX) and LIST_SEPARATOR in old + new):
        old_items, new_items = split_list(old), split_list(new)
        removed = [item for item in old_items if item not in new_items]
        added = [item for item in new_items if item not in old_items]
        if not removed and not added:
            return "zmieniona kolejność: " + html.escape(", ".join(new_items))
        lines = [f"<ins>+ {html.escape(item)}</ins>" for item in added] + [f"<del>− {html.escape(item)}</del>" for item in removed]
        return "<br>".join(lines)
    if column in HTML_COLUMNS:
        old_text, new_text = strip_tags(old), strip_tags(new)
        if old_text == new_text:
            return "zmiana tylko w kodzie HTML (tekst bez zmian)"
        if len(old_text) + len(new_text) > 3000:
            return f"zmieniony tekst ({len(old_text)} → {len(new_text)} znaków) - zobacz kartę produktu"
        return word_diff(old_text, new_text)
    return word_diff(old, new)


def autop(value: str) -> str:
    """Przyblizenie wpautop z WordPressa: akapity z pustych linii, <br> z pojedynczych."""
    value = value.replace("\r\n", "\n").strip()
    if not value:
        return ""
    blocks = []
    for chunk in re.split(r"\n\s*\n", value):
        chunk = chunk.strip()
        if not chunk:
            continue
        blocks.append(chunk if BLOCK_TAG.match(chunk) else "<p>" + chunk.replace("\n", "<br>\n") + "</p>")
    return "\n".join(blocks)


def _number(value: str) -> str:
    return value.strip().replace(".", ",")


def _mark(columns: set[str], changed: set[str]) -> str:
    return " ai-changed" if columns & changed else ""


def attribute_rows(values: dict[str, str], changed: set[str]) -> str:
    rows = []
    weight = values.get("Weight", "").strip()
    if weight:
        rows.append(f'<tr class="{_mark({"Weight"}, changed).strip()}"><th>Waga</th><td class="product_weight">{html.escape(_number(weight))} kg</td></tr>')
    dims = [values.get(column, "").strip() for column in ("Length", "Width", "Height")]
    if any(dims):
        text = " × ".join(_number(dim) if dim else "–" for dim in dims) + " cm"
        rows.append(f'<tr class="{_mark({"Length", "Width", "Height"}, changed).strip()}"><th>Wymiary</th><td class="product_dimensions">{html.escape(text)}</td></tr>')
    for column, value in values.items():
        if not column.startswith(ATTRIBUTE_PREFIX):
            continue
        label = column[len(ATTRIBUTE_PREFIX):]
        if not value.strip():
            if column in changed:
                rows.append(f'<tr class="ai-changed"><th>{html.escape(label)}</th><td><span class="ai-cleared">(usunięty)</span></td></tr>')
            continue
        shown = html.escape(", ".join(split_list(value)))
        rows.append(f'<tr class="{_mark({column}, changed).strip()}"><th>{html.escape(label)}</th><td><p>{shown}</p></td></tr>')
    return "\n".join(rows)


def gallery_html(urls: list[str], alt: str) -> tuple[str, str]:
    """(zdjecie glowne, miniatury galerii)."""
    if not urls:
        return '<div class="woocommerce-product-gallery__image--placeholder"><p class="ai-empty">Brak zdjęcia</p></div>', ""
    main = (
        f'<div class="woocommerce-product-gallery__image"><a><img width="500" height="500" class="wp-post-image" '
        f'src="{html.escape(urls[0])}" alt="{html.escape(alt)}"></a></div>'
    )
    thumbs = ""
    if len(urls) > 1:
        items = "".join(
            f'<div class="woocommerce-product-gallery__image"><a><img src="{html.escape(url)}" alt=""></a></div>' for url in urls[1:]
        )
        thumbs = f'<div class="thumbnails-container"><div id="thumbnails-list">{items}</div></div>'
    return main, thumbs


# --- powiazania i warianty -------------------------------------------------------------------


@dataclass
class VariantGroup:
    """Grupa WPC Linked Variation ze zrodlem product_tag: produkty z tym samym tagiem przelaczaja sie po osiach."""

    name: str
    slugs: set[str]
    axes: list[str]
    image_axes: set[str]


ALL_TAGS = "*"  # grupa z opcja "wszystkie tagi" (terms_all) obejmuje kazdy tag produktu


def groups_from_api(groups: list[dict[str, Any]]) -> dict[str, list[VariantGroup]]:
    """Slug tagu -> opublikowane grupy z /variant-groups (wtyczka bridge 2.1+)."""
    by_slug: dict[str, list[VariantGroup]] = defaultdict(list)
    for item in groups:
        if item.get("status") != "publish" or item.get("source") not in ("product_tag", "tags"):
            continue
        group = VariantGroup(
            name=str(item.get("name", "")),
            slugs=set(item.get("terms", [])),
            axes=list(item.get("attributes", [])),
            image_axes={axis.casefold() for axis in item.get("images", [])},
        )
        for slug in [ALL_TAGS] if item.get("terms_all") else group.slugs:
            by_slug[slug].append(group)
    return by_slug


def grouped_tags(views: list["ProductView"], groups: dict[str, list[VariantGroup]]) -> set[str]:
    """Tagi zmienionych produktow (przed i po), ktore naleza do jakiejs grupy wariantow."""
    return {
        tag
        for view in views
        for values in (view.before, view.after)
        for tag in split_list(values.get("Product Tags", ""))
        if tag_slug(tag) in groups or ALL_TAGS in groups
    }


def tag_slug(name: str) -> str:
    """Slug tagu jak sanitize_title w WordPressie (polskie litery bez ogonkow)."""
    text = name.translate(str.maketrans("łŁ", "lL"))
    text = unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode("ascii").lower()
    return re.sub(r"-+", "-", re.sub(r"[^a-z0-9]+", "-", text)).strip("-")


def load_variant_groups(path: Path | None) -> dict[str, list[VariantGroup]]:
    """Slug tagu -> opublikowane grupy wariantow ze zrzutu (nazwa, status, zrodlo, slugi, licznik, osie, img:, sw:, dd:)."""
    groups: dict[str, list[VariantGroup]] = defaultdict(list)
    if not path or not path.is_file():
        return groups
    for line in path.read_text(encoding="utf-8-sig").splitlines():
        parts = line.split("\t")
        if len(parts) < 6 or parts[1] != "publish" or parts[2] != "product_tag":
            continue
        images = next((part[4:] for part in parts[6:] if part.startswith("img:")), "")
        group = VariantGroup(
            name=parts[0],
            slugs={slug.strip() for slug in parts[3].split(",") if slug.strip()},
            axes=[axis.strip() for axis in parts[5].split(",") if axis.strip()],
            image_axes={axis.strip().casefold() for axis in images.split(",") if axis.strip()},
        )
        for slug in group.slugs:
            groups[slug].append(group)
    return groups


class Catalog:
    """Produkty, ktore podglad zna: wiersze pliku i powiazane produkty dociagniete ze sklepu."""

    def __init__(self, rows: list[dict[str, str]]) -> None:
        self.by_sku: dict[str, dict[str, str]] = {}
        self.by_id: dict[str, dict[str, str]] = {}
        self.by_tag: dict[str, list[dict[str, str]]] = defaultdict(list)
        for row in rows:
            if row.get("SKU"):
                self.by_sku.setdefault(row["SKU"], row)
            if row.get("id"):
                self.by_id.setdefault(str(row["id"]), row)
            for tag in split_list(row.get("Product Tags", "")):
                self.by_tag[tag_slug(tag)].append(row)

    def find(self, ref: str) -> dict[str, str] | None:
        return self.by_sku.get(ref) or self.by_id.get(ref)


@dataclass
class PreviewContext:
    before: Catalog
    after: Catalog
    groups: dict[str, list[VariantGroup]]
    groups_source: str = ""

    def catalog(self, side: str) -> Catalog:
        return self.after if side == "after" else self.before


def catalog_rows(rows: list[dict[str, Any]]) -> list[dict[str, str]]:
    return [{key: _text(value) for key, value in row.items() if key != ROW_NUMBER_KEY} for row in rows]


def build_context(
    rows: list[dict[str, Any]], views: list[ProductView], extra_rows: list[dict[str, Any]], groups: dict[str, list[VariantGroup]], groups_source: str = ""
) -> PreviewContext:
    """Katalog "po" to plik; "przed" to plik z wierszami zmienionych produktow cofnietymi do stanu sklepu."""
    after = catalog_rows(rows)
    previous = {view.row: view.before for view in views if view.changes}
    before = [previous.get(int(row[ROW_NUMBER_KEY]), clean) for row, clean in zip(rows, after)]
    extras = catalog_rows(extra_rows)
    return PreviewContext(before=Catalog(before + extras), after=Catalog(after + extras), groups=groups, groups_source=groups_source)


def related_refs(views: list[ProductView]) -> set[str]:
    """SKU i ID z Up-Sells / Cross-Sells / kompatybilnych w zmienionych produktach (przed i po)."""
    refs: set[str] = set()
    for view in views:
        for values in (view.before, view.after):
            for column in ("Up-Sells", "Cross-Sells", *COMPATIBLE_COLUMNS):
                refs.update(split_list(values.get(column, "")))
    return refs


def attribute_value(values: dict[str, str], axis: str) -> str:
    """Wartosc atrybutu osi; naglowki WP All Export maja inne wielkosci liter niz nazwy atrybutow."""
    wanted = (ATTRIBUTE_PREFIX + axis).casefold()
    for column, value in values.items():
        if column.casefold() == wanted:
            return ", ".join(split_list(value))
    return ""


def _first_image(values: dict[str, str]) -> str:
    images = split_list(values.get("Image URL", ""))
    return images[0] if images else ""


def variants_html(values: dict[str, str], catalog: Catalog, context: PreviewContext, changed: set[str]) -> str:
    """Przelacznik wariantow jak w motywie + nota: grupa, os, rodzenstwo w pliku i powtorzenia osi."""
    blocks = []
    own_sku = values.get("SKU", "")
    for tag in split_list(values.get("Product Tags", "")):
        slug = tag_slug(tag)
        siblings = [row for row in catalog.by_tag.get(slug, []) if row.get("SKU") != own_sku]
        groups = context.groups.get(slug, []) + context.groups.get(ALL_TAGS, [])
        if not groups:
            if siblings:
                blocks.append(
                    f'<p class="ai-note warn">Tag „{html.escape(tag)}” ma {len(siblings)} innych produktów, '
                    "ale nie należy do żadnej grupy wariantów - przełącznik się nie pojawi.</p>"
                )
            continue
        members = [values] + siblings
        for group in groups:
            attributes = []
            for axis in group.axes:
                current = attribute_value(values, axis)
                options: dict[str, dict[str, str]] = {}
                for member in members:
                    value = attribute_value(member, axis)
                    if value and value not in options:
                        options[value] = member
                terms = []
                for value, member in options.items():
                    active = value == current
                    inner = html.escape(value)
                    kind = "button"
                    if axis.casefold() in group.image_axes and _first_image(member):
                        kind = "image wpclv-term-image--with-label"
                        inner = (
                            f'<span class="wpclv-term-media has-label"><span class="wpclv-term-image-frame">'
                            f'<img src="{html.escape(_first_image(member))}" alt="{html.escape(value)}"></span>'
                            f'<span class="wpclv-term-image-label">{html.escape(value)}</span></span>'
                        )
                    terms.append(f'<div class="wpclv-term-{kind} wpclv-term{" active" if active else ""}"><span>{inner}</span></div>')
                if current and terms:
                    attributes.append(
                        f'<div class="wpclv-attribute"><div class="wpclv-attribute-label">{html.escape(axis)}</div>'
                        f'<div class="wpclv-terms">{"".join(terms)}</div></div>'
                    )
            own_key = tuple(attribute_value(values, axis) for axis in group.axes)
            duplicates = [row.get("SKU", "") for row in siblings if tuple(attribute_value(row, axis) for axis in group.axes) == own_key]
            missing = [axis for axis in group.axes if not attribute_value(values, axis)]
            notes = [
                f'Grupa wariantów <strong>{html.escape(group.name)}</strong>, oś: <strong>{html.escape(", ".join(group.axes))}</strong>, '
                f"tag „{html.escape(tag)}”, innych produktów z tym tagiem: {len(siblings)}."
            ]
            level = ""
            if missing:
                level = "bad"
                notes.append(f"DO_SPRAWDZENIA: produkt nie ma wartości osi {html.escape(', '.join(missing))} - nie pokaże się w przełączniku.")
            if duplicates:
                level = "bad"
                notes.append(f"DO_SPRAWDZENIA: ta sama wartość osi co {html.escape(', '.join(duplicates))} - przełącznik ich nie odróżni.")
            axis_columns = {(ATTRIBUTE_PREFIX + axis).casefold() for axis in group.axes}
            touched = "Product Tags" in changed or any(column.casefold() in axis_columns for column in changed)
            mark = " ai-changed" if touched else ""
            blocks.append(
                f'<div class="wpclv-attributes{mark}">{"".join(attributes)}</div>'
                f'<p class="ai-note {level}">{" ".join(notes)}</p>'
            )
    return "".join(blocks)


def suggestions_html(title: str, column: str, values: dict[str, str], other: dict[str, str], catalog: Catalog, changed: set[str], extra_class: str = "") -> str:
    """Sekcja kafelkow jak alkan_render_product_suggestions_section; dodane/usuniete wzgledem drugiej strony."""
    refs = split_list(values.get(column, ""))
    other_refs = set(split_list(other.get(column, "")))
    removed = [ref for ref in split_list(other.get(column, "")) if ref not in refs] if column in changed else []
    if not refs and not removed:
        return ""
    tiles = []
    for ref, state in [(ref, "added" if column in changed and ref not in other_refs else "") for ref in refs] + [(ref, "removed") for ref in removed]:
        product = catalog.find(ref)
        css = f"suggestion-card{' ai-' + state if state else ''}{'' if product else ' ai-missing'}"
        if product:
            image = _first_image(product)
            media = f'<a class="suggestion-card__media"><img src="{html.escape(image)}" alt=""></a>' if image else ""
            name = html.escape(product.get("Title", ref))
        else:
            media, name = "", "(produktu nie ma w pliku ani w sklepie)"
        tiles.append(
            f'<div class="product-grid-item {css}"><div class="product-item-content">{media}'
            f'<p class="product-name"><a>{name}</a></p><p class="ai-sku">{html.escape(ref)}</p></div></div>'
        )
    count = len(refs)
    label = f"{count} {'produkt' if count == 1 else 'produkty' if count % 10 in (2, 3, 4) and count % 100 not in (12, 13, 14) else 'produktów'}"
    return (
        f'<section class="ai-section up-sells upsells products product-suggestions {extra_class}{" ai-changed" if column in changed else ""}">'
        f'<div class="product-suggestions__head"><p class="product-suggestions__title">{html.escape(title)}</p>'
        f'<span class="product-suggestions__count">{label}</span></div>'
        f'<div class="products columns-1 product-suggestions__grid">{"".join(tiles)}</div></section>'
    )


def related_html(values: dict[str, str], other: dict[str, str], catalog: Catalog, changed: set[str]) -> str:
    compatible_column = next((column for column in COMPATIBLE_COLUMNS if column in values or column in other), COMPATIBLE_COLUMNS[0])
    sections = [
        suggestions_html("Produkty kompatybilne", compatible_column, values, other, catalog, changed, "product-suggestions--compatible"),
        suggestions_html("Polecamy również", "Up-Sells", values, other, catalog, changed, "product-suggestions--upsells"),
    ]
    cross = suggestions_html("W koszyku: polecane do tego produktu (Cross-Sells)", "Cross-Sells", values, other, catalog, changed)
    if cross:
        sections.append('<p class="ai-note">Cross-Sells klient widzi dopiero w koszyku, nie na karcie produktu.</p>' + cross)
    return "".join(sections)


def _money(value: float | None) -> str:
    if value is None:
        return "—"
    return f"{value:,.2f}".replace(",", " ").replace(".", ",") + " zł"


def availability_html(values: dict[str, str], changed: set[str]) -> str:
    """Linia dostepnosci jak w motywie (bez magazynu zewnetrznego): stan, zamowienia bez stanu, "Wysylka w ciagu"."""
    try:
        stock = float((values.get("Stock") or "0").replace(",", "."))
    except ValueError:
        stock = 0.0
    backorders = values.get("Backorders", "").strip().lower()
    within = values.get("Wysyłka w ciągu", "").strip()
    if stock > 0:
        text, css = "W magazynie - wysyłka w 24 godziny", "ok"
    elif within.casefold() == "niedostępne" or backorders not in ("notify", "yes"):
        text, css = "Niedostępne", "none"
    else:
        text, css = "Produkt na zamówienie", "order"
    detail = f"Wysyłka w ciągu: {html.escape(within)}" if within else ""
    mark = " ai-changed" if AVAILABILITY_COLUMNS & changed else ""
    return (
        f'<div class="ai-ship{mark}"><p class="ai-availability {css}">{html.escape(text)}</p>'
        f'<p class="ai-ship-class">{detail}{" · " if detail else ""}stan w sklepie: {html.escape(values.get("Stock") or "0")}, '
        f'zamówienia bez stanu: {html.escape(backorders or "no")}</p></div>'
    )


def shipping_html(values: dict[str, str], side: str, shipping: dict[str, Any] | None, changed: set[str], note: str = "") -> str:
    """Metody dostawy policzone przez sklep (/shipping-preview); znikajace i nowe metody zaznaczone."""
    mark = " ai-changed" if set(SHIPPING_COLUMNS) & changed else ""
    shipping_class = values.get("Shipping Class", "")
    head = f'<p class="ai-ship-class">Klasa wysyłkowa: <strong>{html.escape(shipping_class or "(brak)")}</strong>'
    if values.get("Weight"):
        head += f" · waga {html.escape(_number(values['Weight']))} kg"
    head += "</p>"
    if not shipping:
        return f'<div class="ai-ship{mark}"><p class="box-title">Metody dostawy</p>{head}<p class="ai-note warn">{html.escape(note or "Sklep nie policzył metod dostawy.")}</p></div>'
    current = shipping.get("before") or {}
    target = (shipping.get("after") or current) if side == "after" else current
    other = current if side == "after" else (shipping.get("after") or current)
    titles = {method["title"] for method in target.get("methods", [])}
    other_titles = {method["title"] for method in other.get("methods", [])}
    tiles = []
    shown = [(method, "added" if side == "after" and method["title"] not in other_titles else "") for method in target.get("methods", [])]
    if side == "after":
        shown += [(method, "removed") for method in other.get("methods", []) if method["title"] not in titles]
    for method, state in shown:
        image = f'<img src="{html.escape(method["image_url"])}" alt="">' if method.get("image_url") else ""
        price = _money(method.get("price"))
        if method.get("cod_price") is not None:
            price += f" / pobranie {_money(method['cod_price'])}"
        tiles.append(
            f'<div class="ai-ship-method{" ai-" + state if state else ""}" style="--ship-border: {html.escape(method.get("border_col") or "#d8e1ef")}">'
            f'{image}<strong>{html.escape(method["title"])}</strong><span>{"usunięta - " if state == "removed" else ""}{price}</span></div>'
        )
    priced = [method["price"] for method in target.get("methods", []) if method.get("price") and not method.get("is_local_pickup")]
    title = f"Dostawa od {_money(min(priced))}" if priced else "Dostępne metody dostawy"
    grid = "".join(tiles) or '<p class="ai-note bad">Brak dostępnych metod dostawy - produktu nie da się wysłać.</p>'
    notes = [f'<p class="ai-note warn">{html.escape(warning)}</p>' for warning in shipping.get("warnings", [])]
    if side == "after" and shipping.get("after"):
        gone = sorted(other_titles - titles)
        new = sorted(titles - other_titles)
        if gone or new:
            parts = ([f"znikają: {', '.join(gone)}"] if gone else []) + ([f"dochodzą: {', '.join(new)}"] if new else [])
            notes.append(f'<p class="ai-note {"bad" if gone else ""}">Po zmianie {html.escape("; ".join(parts))}.</p>')
    return (
        f'<div class="ai-ship{mark}"><p class="box-title">{title} <span class="ai-sku">(1 szt., cena orientacyjna)</span></p>'
        f'{head}<div class="ai-ship-grid">{grid}</div>{"".join(notes)}</div>'
    )


def render_card(
    view: ProductView,
    side: str,
    *,
    theme_base: str | None,
    context: PreviewContext | None = None,
    shipping: dict[str, Any] | None = None,
    shipping_note: str = "",
) -> str:
    """Pelny dokument HTML karty produktu (do ramki iframe bez skryptow)."""
    values = view.after if side == "after" else view.before
    other = view.before if side == "after" else view.after
    changed = view.changed_columns
    catalog = context.catalog(side) if context else Catalog([])
    variants = variants_html(values, catalog, context, changed) if context else ""
    related = related_html(values, other, catalog, changed)
    availability = availability_html(values, changed)
    delivery = shipping_html(values, side, shipping, changed, shipping_note)
    title = values.get("Title", "")
    ean = values.get("hwp_product_gtin", "").strip()
    manufacturer = ", ".join(split_list(values.get(ATTRIBUTE_PREFIX + "Producent", "")))
    badges = ""
    if ean or manufacturer:
        items = []
        if ean:
            items.append(f'<p class="product-spec-badge product-spec-badge--ean"><span class="lbl">EAN:</span> <span class="val">{html.escape(ean)}</span></p>')
        if manufacturer:
            mark = _mark({ATTRIBUTE_PREFIX + "Producent"}, changed)
            items.append(f'<p class="product-spec-badge product-spec-badge--manufacturer{mark}"><span class="lbl">Producent:</span> <span class="val">{html.escape(manufacturer)}</span></p>')
        badges = f'<div class="product-purchase-meta">{"".join(items)}</div>'

    categories = split_list(values.get("Kategorie produktów", ""))
    crumbs = " / ".join(html.escape(part.strip()) for part in categories[0].split(">")) if categories else ""
    crumbs_html = f'<p class="ai-crumbs{_mark({"Kategorie produktów"}, changed)}">Sklep / {crumbs}</p>' if crumbs else ""

    main_image, thumbs = gallery_html(split_list(values.get("Image URL", "")), title)

    short = values.get("Short Description", "")
    if short.strip():
        short_html = f'<div class="product-single-box product-single-box-description{_mark({"Short Description"}, changed)}">{autop(short)}</div>'
    elif "Short Description" in changed:
        short_html = '<div class="product-single-box product-single-box-description ai-changed"><p class="ai-cleared">(krótki opis usunięty)</p></div>'
    else:
        short_html = ""

    content = values.get("Content", "")
    content_html = autop(content) if content.strip() else '<p class="ai-empty">(brak opisu)</p>'
    attributes = attribute_rows(values, changed)
    attribute_columns = {column for column in changed if column.startswith(ATTRIBUTE_PREFIX)} | {"Weight", "Length", "Width", "Height"}

    if theme_base is not None:
        links = "\n".join(f'<link rel="stylesheet" href="{theme_base}/{path}">' for path in THEME_STYLESHEETS)
    else:
        links = f"<style>{FALLBACK_CSS}</style>"

    return f"""<!doctype html>
<html lang="pl"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="referrer" content="no-referrer">
<link rel="stylesheet" href="{FONTS_URL}">
{links}
<style>{CARD_CSS}</style>
</head>
<body class="single-product woocommerce woocommerce-page">
<div id="product-{view.product_id or ''}" class="product type-product{' product-with-thumbnails' if thumbs else ''}">
<div class="product-page-shell">
  <div class="product-main-hero">
    {crumbs_html}
    <div class="product-main-header">
      <h1 class="page-name product-title{_mark({"Title"}, changed)}">{html.escape(title)}</h1>
      {badges}
    </div>
    <div class="product-main-grid">
      <div class="product-media-column">
        <div class="woocommerce-product-gallery woocommerce-product-gallery--with-images{' woocommerce-product-gallery--with-thumbnails' if thumbs else ''} images{_mark({"Image URL"}, changed)}">
          <figure class="woocommerce-product-gallery__wrapper"><div id="big-pic">{main_image}</div>{thumbs}</figure>
        </div>
      </div>
      <div class="summary entry-summary"><div class="product-summary-stack">
        {availability}
        {variants}
        {delivery}
        <div class="product-detail-cards">{short_html}</div>
      </div></div>
    </div>
  </div>
  <div class="woocommerce-tabs wc-tabs-wrapper">
    <div class="product-tab-content{_mark({"Content"}, changed)}" id="tab-description"><h2>Opis</h2>
{content_html}
    </div>
    <div class="product-tab-content{_mark(attribute_columns, changed)}" id="tab-additional_information"><h2>Informacje dodatkowe</h2>
      <table class="shop_attributes">{attributes}</table>
    </div>
  </div>
  {related}
</div>
</div>
</body></html>"""
