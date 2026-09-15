from __future__ import annotations

import sys
import tempfile
import unittest
from pathlib import Path

import fitz

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from build_hager_berker_catalog_cross_upsell import (  # noqa: E402
    CatalogEvidence,
    build_catalog_page_index,
    candidate_catalog_codes,
    catalog_filter_candidate,
    configuration_confirmed,
    explicit_catalog_references,
    explicit_recommendations,
    line_contains_code,
    safe_relaxed_up_sell_candidate,
)
from build_hager_berker_cross_upsell import CatalogProduct  # noqa: E402
from build_hager_berker_variants import Record  # noqa: E402


def record(
    sku: str,
    title: str,
    product_type: str,
    series: str,
    color: str = "Biały",
    multiplicity: str = "",
    alias_code: str = "",
) -> Record:
    return Record(
        excel_row=2,
        product_id=sku,
        title=title,
        sku=sku,
        code=sku.split("/", 1)[0],
        product_type=product_type,
        series=series,
        primary_series=series.split("|", 1)[0],
        color=color,
        manufacturer_color_before=color,
        manufacturer_color=color,
        material="",
        multiplicity=multiplicity,
        tags_before="",
        core=title,
        family="inne",
        frame=product_type == "Ramki",
        color_detail=color,
        material_detail="",
        finish_detail="",
        orientation="",
        network_category="",
        screened="",
        grounding="",
        terminal_type="",
        voltage="",
        nominal_current="",
        pole_count="",
        shutters="",
        detection_range="",
        output_count="",
        alias_code=alias_code,
    )


class CatalogIndexTests(unittest.TestCase):
    def test_spaced_catalog_code_is_detected(self) -> None:
        self.assertTrue(line_contains_code("biały połysk 53 1623 89 92 10", "5316238992"))

    def test_index_maps_code_to_printed_page(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "catalog.pdf"
            document = fitz.open()
            page = document.new_page()
            page.insert_text((60, 60), "Spis indeksow katalogowych")
            page.insert_text((120, 120), "5316238992")
            page.insert_text((190, 120), "36")
            document.save(path)
            document.close()

            reopened = fitz.open(path)
            try:
                index = build_catalog_page_index(reopened, {"5316238992"})
            finally:
                reopened.close()

        self.assertEqual(index["5316238992"], 36)

    def test_hager_sku_uses_berker_alias_first(self) -> None:
        product = CatalogProduct(record(
            "WDEM3035/HAG",
            "Łącznik świecznikowy Hager WDEM3035 (Berker 533035)",
            "Mechanizmy",
            "one.platform",
            color="",
            alias_code="533035",
        ))

        self.assertEqual(candidate_catalog_codes(product)[0], "533035")


class CatalogCompatibilityTests(unittest.TestCase):
    def test_relaxed_up_sell_allows_same_standard_socket_with_shutters(self) -> None:
        source = CatalogProduct(record(
            "BASE/HAG", "Gniazdo SCHUKO białe", "Gniazda", "Lumina", color="Biały",
        ))
        target = CatalogProduct(record(
            "PLUS/HAG", "Gniazdo SCHUKO z przesłonami białe", "Gniazda", "Lumina", color="Biały",
        ))

        recommendation = safe_relaxed_up_sell_candidate(source, target)

        self.assertIsNotNone(recommendation)
        self.assertIn("przesłony", recommendation.reason)

    def test_relaxed_up_sell_rejects_vga_to_schuko(self) -> None:
        source = CatalogProduct(record(
            "VGA/HAG", "Gniazdo VGA białe", "Gniazda", "Q.1", color="Biały",
        ))
        target = CatalogProduct(record(
            "POWER/HAG", "Gniazdo SCHUKO z przesłonami białe", "Gniazda", "Q.1", color="Biały",
        ))

        self.assertIsNone(safe_relaxed_up_sell_candidate(source, target))

    def test_relaxed_up_sell_allows_rj45_category_upgrade(self) -> None:
        source = CatalogProduct(record(
            "CAT5/HAG", "Gniazdo komputerowe RJ45 kat.5 UTP białe", "Gniazda", "Lumina", color="Biały",
        ))
        target = CatalogProduct(record(
            "CAT6/HAG", "Gniazdo komputerowe RJ45 kat.6 UTP białe", "Gniazda", "Lumina", color="Biały",
        ))

        recommendation = safe_relaxed_up_sell_candidate(source, target)

        self.assertIsNotNone(recommendation)
        self.assertIn("kategoria", recommendation.reason)

    def test_relaxed_up_sell_preserves_socket_cover(self) -> None:
        source = CatalogProduct(record(
            "COVERED/HAG", "Gniazdo SCHUKO z pokrywą i polem opisowym białe", "Gniazda", "Lumina", color="Biały",
        ))
        target = CatalogProduct(record(
            "OPEN/HAG", "Gniazdo SCHUKO z przesłonami białe", "Gniazda", "Lumina", color="Biały",
        ))

        self.assertIsNone(safe_relaxed_up_sell_candidate(source, target))

    def test_rotary_cover_does_not_match_rocker_mechanism(self) -> None:
        cover = CatalogProduct(record(
            "COVER/HAG", "Płytka z pokrętłem do łącznika żaluzjowego obrotowego", "Płytki", "K.1",
        ))
        mechanism = CatalogProduct(record(
            "MECH/HAG", "Łącznik żaluzjowy 2-klawiszowy, mechanizm", "Mechanizmy", "one.platform",
        ))

        self.assertFalse(configuration_confirmed(cover, mechanism))

    def test_explicit_catalog_reference_can_link_frame_and_seal(self) -> None:
        frame = CatalogProduct(record(
            "10116019/HAG", "Ramka 1-krotna z polem opisowym", "Ramki", "Q.1", multiplicity="1x",
        ))
        seal = CatalogProduct(record(
            "10107200/HAG", "Zestaw uszczelek IP44", "Akcesoria", "Q.1",
        ))
        products = {frame.sku: frame, seal.sku: seal}
        evidence = {
            frame.sku: CatalogEvidence("10116019", 202, 203, "", "ramka"),
            seal.sku: CatalogEvidence("10107200", 204, 205, "", "zestaw uszczelniajacy"),
        }

        recommendations = explicit_recommendations(
            [(frame.sku, seal.sku, 202, "10116019", "10107200")], products, evidence,
        )

        self.assertEqual(len(recommendations), 2)

    def test_explicit_frame_reference_still_requires_matching_series_and_color(self) -> None:
        frame = CatalogProduct(record(
            "FRAME/HAG", "Ramka 1-krotna biała", "Ramki", "Serie 1930", color="Biały", multiplicity="1x",
        ))
        sensor = CatalogProduct(record(
            "SENSOR/HAG", "Czujnik ruchu antracyt", "Czujniki", "S.1|B.3", color="Antracyt",
        ))
        products = {frame.sku: frame, sensor.sku: sensor}
        evidence = {
            frame.sku: CatalogEvidence("138209", 357, 358, "", "ramka"),
            sensor.sku: CatalogEvidence("85345185", 511, 512, "", "czujnik"),
        }

        recommendations = explicit_recommendations(
            [(sensor.sku, frame.sku, 511, "85345185", "138209")], products, evidence,
        )

        self.assertEqual(recommendations, [])

    def test_double_key_and_double_mechanism_require_catalog_confirmation(self) -> None:
        key = CatalogProduct(record(
            "5316238992/HAG",
            "Klawisze; kremowy; B.Kwadrat 5316238992",
            "Klawisze",
            "B. Kwadrat",
            color="Kremowy",
            multiplicity="2x",
        ))
        mechanism = CatalogProduct(record(
            "WDEM3035/HAG",
            "Łącznik świecznikowy 2-klawiszowy mechanizm WDEM3035 (Berker 533035)",
            "Mechanizmy",
            "one.platform",
            color="",
            multiplicity="2x",
            alias_code="533035",
        ))
        key_evidence = CatalogEvidence(
            "5316238992", 36, 37, "Klawisze podwójne", norm_text("B. Kwadrat klawisze podwójne do łącznika 2-klawiszowego"),
        )
        mechanism_evidence = CatalogEvidence(
            "533035", 35, 36, "Łącznik 2-klawiszowy", norm_text("łącznik 2-klawiszowy seryjny świecznikowy"),
        )

        recommendation = catalog_filter_candidate(key, mechanism, key_evidence, mechanism_evidence)

        self.assertIsNotNone(recommendation)
        self.assertIn("KATALOGU", recommendation.method)
        self.assertIn("s. 36", recommendation.catalog_evidence)

    def test_explicit_optional_reference_is_assigned_to_following_product_block(self) -> None:
        source = CatalogProduct(record(
            "WDEM3035/HAG", "Łącznik 2-klawiszowy mechanizm", "Mechanizmy", "one.platform", color="", alias_code="533035",
        ))
        target = CatalogProduct(record(
            "WDE1675W/HAG", "Moduł LED z zaciskiem N", "Mechanizmy", "one.platform", color="", alias_code="1675",
        ))
        evidence = {
            source.sku: CatalogEvidence("533035", 1, 2, "", "one platform"),
            target.sku: CatalogEvidence("1675", 10, 11, "", "one platform"),
        }
        products = {source.sku: source, target.sku: target}

        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "catalog.pdf"
            document = fitz.open()
            document.new_page()
            page = document.new_page()
            page.insert_text((420, 70), "Wspolpracuje z opcjonalnie")
            page.insert_text((470, 105), "1675 ..")
            page.insert_text((540, 105), "10")
            page.insert_text((345, 170), "53 3035")
            document.save(path)
            document.close()

            reopened = fitz.open(path)
            try:
                references = explicit_catalog_references(reopened, evidence, products)
            finally:
                reopened.close()

        self.assertIn((source.sku, target.sku, 1, "533035", "1675"), references)

    def test_target_page_disambiguates_short_catalog_prefix(self) -> None:
        source = CatalogProduct(record(
            "FRAME/HAG", "Ramka 1-krotna", "Ramki", "Q.1", multiplicity="1x", alias_code="10116019",
        ))
        q_seal = CatalogProduct(record(
            "Q-SEAL/HAG", "Zestaw uszczelek Q.1", "Akcesoria", "Q.1", alias_code="10107200",
        ))
        r_seal = CatalogProduct(record(
            "R-SEAL/HAG", "Zestaw uszczelek R.x", "Akcesoria", "R.x", alias_code="10107600",
        ))
        evidence = {
            source.sku: CatalogEvidence("10116019", 1, 2, "", "ramka"),
            q_seal.sku: CatalogEvidence("10107200", 251, 252, "", "uszczelki q"),
            r_seal.sku: CatalogEvidence("10107600", 349, 350, "", "uszczelki r"),
        }
        products = {item.sku: item for item in (source, q_seal, r_seal)}

        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "catalog.pdf"
            document = fitz.open()
            document.new_page()
            page = document.new_page()
            page.insert_text((420, 70), "Wspolpracuje z opcjonalnie")
            page.insert_text((470, 105), "10107 ..")
            page.insert_text((540, 105), "251")
            page.insert_text((345, 170), "10 1160 19")
            for _ in range(349):
                document.new_page()
            document.save(path)
            document.close()

            reopened = fitz.open(path)
            try:
                references = explicit_catalog_references(reopened, evidence, products)
            finally:
                reopened.close()

        self.assertIn((source.sku, q_seal.sku, 1, "10116019", "10107200"), references)
        self.assertNotIn((source.sku, r_seal.sku, 1, "10116019", "10107600"), references)


def norm_text(value: str) -> str:
    from build_hager_berker_variants import norm

    return norm(value)


if __name__ == "__main__":
    unittest.main()
