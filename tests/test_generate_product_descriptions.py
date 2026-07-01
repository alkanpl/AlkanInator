from __future__ import annotations

import sys
import tempfile
import os
import unittest
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from generate_product_descriptions import (  # noqa: E402
    CODEX_BRIEF_SHEETS,
    H2_QUESTION_BUILDERS,
    algorithmic_style_issues,
    build_codex_brief_for_sku,
    build_description_html,
    classify_family,
    prose_material,
    render_benefit_item,
    validate_codex_description,
    validate_codex_review_file,
    validate_generated_description,
    write_codex_review_from_description,
    write_codex_brief_report,
)
from codex_description_agent import discover_latest_brief_batch  # noqa: E402


class AlgorithmicDescriptionTest(unittest.TestCase):
    def test_garden_lamp_family(self) -> None:
        self.assertEqual(classify_family("Lampa ogrodowa kula STONO 200mm IP65 biała", {}), "ogrodowa")
        self.assertEqual(classify_family("Oprawa elewacyjna LED REKA 7W z czujnikiem", {}), "ogrodowa")
        # "z czujnikiem ruchu" w nazwie oprawy NIE czyni jej akcesorium
        self.assertNotEqual(classify_family("Plafon LED z czujnikiem ruchu BENO", {}), "akcesorium")

    def test_h2_questions_keep_keyword_in_nominative(self) -> None:
        # Zadna forma H2 nie wymusza odmiany frazy (brak "wybrać <fraza>").
        for builder in H2_QUESTION_BUILDERS:
            self.assertNotIn("wybrać", builder("lampa ogrodowa STONO").lower())

    def test_benefit_item_bolds_feature_before_dash(self) -> None:
        self.assertEqual(
            render_benefit_item("Trzonek E27 - daje swobodę doboru żarówki."),
            "  <li>- <strong>Trzonek E27</strong> - daje swobodę doboru żarówki.</li>",
        )

    def test_prose_material_replaces_pipe(self) -> None:
        self.assertEqual(prose_material({"Materiał": "ABS|PE"}), "ABS i PE")

    def test_full_description_passes_own_validation(self) -> None:
        row = pd.Series({
            "name": "Lampa ogrodowa kula STONO 200mm IP65 biała Kanlux 45930",
            "sku": "45930/KAN",
            "features": (
                '{"Typ produktu": "Lampa ogrodowa", "Seria": "STONO", "Trzonek": "E27",'
                ' "Stopień ochrony [IP]": "IP 65", "Kolor": "Biały", "Materiał": "ABS|PE",'
                ' "Kształt": "Okrągły", "Źródło światła": "Niezintegrowane"}'
            ),
        })
        html = build_description_html(row)
        from generate_product_descriptions import DEFAULT_MIN_CHARS_NO_SPACES, build_seo_keyword, collect_attributes
        keyword = build_seo_keyword(collect_attributes(row), "Lampa ogrodowa kula STONO")
        self.assertEqual(validate_generated_description(html, keyword, DEFAULT_MIN_CHARS_NO_SPACES), [])
        self.assertNotIn("|", html.split("Specyfikacja")[0])  # brak surowego separatora w prozie

    def test_algorithmic_style_validator_detects_parameter_list(self) -> None:
        html = """
        <p>Plafon LED TEST to produkt do korytarza.</p>
        <h3>Najważniejsze cechy</h3>
        <ul>
          <li>- <strong>Moc 12 W</strong> - pozwala dobrać oprawę do pomieszczenia.</li>
          <li>- <strong>Trzonek E27</strong> - ułatwia dobór żarówki.</li>
          <li>- <strong>Stopień ochrony IP54</strong> - oznacza ochronę przed pyłem.</li>
        </ul>
        <h3>Specyfikacja techniczna</h3>
        """
        self.assertIn("lista cech zaczyna sie od parametrow zamiast decyzji kupujacego", algorithmic_style_issues(html))

    def test_planned_plafon_description_uses_buyer_decision_highlights(self) -> None:
        row = pd.Series({
            "name": "Plafon LED BENO 18W 4000K IP54 biały Kanlux",
            "sku": "TEST/KAN",
            "features": (
                '{"Typ produktu": "Plafon LED", "Seria": "BENO", "Moc [W]": "18",'
                ' "Temperatura barwowa [K]": "4000", "Czujnik ruchu": "Tak",'
                ' "Stopień ochrony [IP]": "IP 54", "Kolor": "Biały"}'
            ),
        })
        html = build_description_html(row)
        before_spec = html.split("Specyfikacja techniczna")[0]
        self.assertIn("Automatyczne światło", before_spec)
        self.assertNotIn("będzie dobrym wyborem wtedy, gdy", before_spec)
        self.assertNotRegex(before_spec, r"<strong>(Moc|Trzonek|Stopień ochrony|Kolor)")


class CodexDescriptionBriefTest(unittest.TestCase):
    def test_brief_discovery_prefers_complete_newest_batch_over_old_root_file(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            reports = Path(tmp) / "reports" / "descriptions"
            batch = reports / "usage_test_10" / "briefs"
            batch.mkdir(parents=True)
            old_brief = reports / "description_codex_brief_33482.xlsx"
            old_brief.write_bytes(b"old")
            os.utime(old_brief, (1000, 1000))
            expected = []
            for index, sku in enumerate(["22606", "26980", "27163"]):
                path = batch / f"description_codex_brief_{sku}.xlsx"
                path.write_bytes(b"new")
                os.utime(path, (2000 + index, 2000 + index))
                expected.append(path)

            discovered = discover_latest_brief_batch(reports)

        self.assertEqual(discovered, expected)

    def test_brief_discovery_with_sku_uses_newest_matching_file(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            reports = Path(tmp)
            old_dir = reports / "old"
            new_dir = reports / "new"
            old_dir.mkdir()
            new_dir.mkdir()
            old_path = old_dir / "description_codex_brief_33482.xlsx"
            new_path = new_dir / "description_codex_brief_33482.xlsx"
            old_path.write_bytes(b"old")
            new_path.write_bytes(b"new")
            os.utime(old_path, (1000, 1000))
            os.utime(new_path, (2000, 2000))

            discovered = discover_latest_brief_batch(reports, "33482/KAN")

        self.assertEqual(discovered, [new_path])

    def test_codex_brief_uses_filtered_features_for_accessory(self) -> None:
        df = pd.DataFrame(
            [
                {
                    "Kod": "33482",
                    "EAN": "5905339334824",
                    "new_title": "Siatka ochronna do FL AGOR HI GRID 240W Kanlux 33482",
                    "attr_typ": "Siatka ochronna",
                    "attr_seria": "FL AGOR",
                    "attr_moc": "240W",
                    "attr_material": "metal",
                    "attr_pasuje_do": "do FL AGOR HI GRID",
                    "attr_producent": "Kanlux",
                }
            ]
        )

        catalog_knowledge = {
            "top_attribute_values_by_category": [
                {"attribute": "Materiał", "value": "Metal", "count": 10},
            ],
        }

        brief = build_codex_brief_for_sku(df, "33482/KAN", "new_title", catalog_knowledge)
        product_facts = {(row["fact"], row["value"]) for row in brief["product_facts"]}
        rejected_facts = {(row["fact"], row["value"]) for row in brief["rejected_facts"]}
        compatibility_facts = {(row["fact"], row["value"]) for row in brief["compatibility_facts"]}

        self.assertIn(("Typ produktu", "Siatka ochronna"), product_facts)
        self.assertIn(("Materiał", "Metal"), product_facts)
        self.assertNotIn(("Moc [W]", "240"), product_facts)
        self.assertIn(("Moc [W]", "240"), rejected_facts)
        self.assertIn(("Pasuje do", "do FL AGOR HI GRID"), compatibility_facts)

    def test_validator_detects_rejected_power_claim_for_accessory(self) -> None:
        rejected_facts = [
            {
                "fact": "Moc [W]",
                "value": "240",
                "source": "raw_feature_rejected_by_filter",
                "reason": "accessory_compatible_fixture_power_not_product_power",
            }
        ]

        checks = validate_codex_description(
            "<p>Ta siatka ochronna ma moc 240W i sprawdza się w instalacji.</p>",
            [],
            rejected_facts,
        )

        self.assertTrue(any(check["status"] == "ERROR" for check in checks))

    def test_validator_detects_encoding_damage_short_text_and_bad_list_format(self) -> None:
        checks = validate_codex_description(
            """
            <p>Siatka ochronna pomaga zabezpieczy? element o?wietleniowy.</p>
            <h2>Dlaczego warto wybra? produkt?</h2>
            <ul><li>- metalowa konstrukcja</li></ul>
            <h3>Specyfikacja techniczna</h3>
            <ul><li>Materiał: Metal</li></ul>
            """,
            [],
            [],
        )

        checks_by_name = {check["check"]: check["status"] for check in checks}
        self.assertEqual(checks_by_name["encoding_integrity"], "ERROR")
        self.assertEqual(checks_by_name["description_length"], "WARNING")
        self.assertEqual(checks_by_name["html_specification_labels"], "WARNING")

    def test_validator_detects_meta_description_language(self) -> None:
        repeated_context = (
            "Siatka ochronna Kanlux 33482 jest metalowym akcesorium z serii AGOR, przeznaczonym "
            "do zastosowania z kompatybilną oprawą FL AGOR HI GRID. "
        )
        checks = validate_codex_description(
            f"""
            <p>{repeated_context * 4}</p>
            <p>Taki sposób opisu od razu porządkuje wybór i pokazuje, że mówimy o elemencie powiązanym z konkretną oprawą.</p>
            <p>Dla klienta oznacza to prostszy wybór i mniejsze ryzyko pomyłki przy kompletowaniu kompatybilnych elementów.</p>
            <p>Siatka ochronna Kanlux 33482 nie zmienia parametrów oprawy i nie udaje samodzielnego źródła światła.</p>
            <h2>Siatka ochronna do opraw FL AGOR HI GRID</h2>
            <ul><li>Metalowa konstrukcja.</li><li>Kompatybilność z FL AGOR HI GRID.</li></ul>
            <h3>Specyfikacja techniczna</h3>
            <ul><li>Materiał: Metal</li></ul>
            """,
            [],
            [],
        )

        self.assertTrue(
            any(check["status"] == "ERROR" and check["check"] == "meta_description_language" for check in checks)
        )

    def test_validator_detects_abstract_filler_language(self) -> None:
        repeated_context = (
            "Siatka ochronna Kanlux 33482 jest metalowym akcesorium z serii AGOR, przeznaczonym "
            "do zastosowania z kompatybilną oprawą FL AGOR HI GRID. "
        )
        checks = validate_codex_description(
            f"""
            <p>{repeated_context * 4}</p>
            <p>Sam typ produktu jasno wskazuje, że jest to siatka ochronna, czyli komponent wspierający funkcję zabezpieczającą całego zestawu.</p>
            <p>W połączeniu z marką Kanlux oraz serią AGOR otrzymujesz akcesorium przygotowane do pracy w obrębie spójnej rodziny produktowej.</p>
            <h2>Siatka ochronna do opraw FL AGOR HI GRID</h2>
            <ul><li>Metalowa konstrukcja.</li><li>Kompatybilność z FL AGOR HI GRID.</li></ul>
            <h3>Specyfikacja techniczna</h3>
            <ul><li>Materiał: Metal</li></ul>
            """,
            [],
            [],
        )

        self.assertTrue(
            any(check["status"] == "ERROR" and check["check"] == "meta_description_language" for check in checks)
        )

    def test_validator_detects_batch_boilerplate_and_missing_diacritics(self) -> None:
        checks = validate_codex_description(
            """
            <p>Oprawa hermetyczna LED MAH HI 60cm 19W 2400lm marki Kanlux z serii MAH to oprawa
            hermetyczna opisany przez kluczowe parametry techniczne. Taki zestaw danych pozwala
            szybko ocenic model pod katem doboru do projektu i porownac go z innymi pozycjami
            w tej samej grupie produktowej.</p>
            <p>Przy wyborze warto patrzec na zestaw cech, ktore najlatwiej rozrozniaja modele.
            Dzieki temu opis pozostaje techniczny, a sklep chce podac informacje bez nadmiaru ogolnikow.</p>
            <p>W praktyce pomaga to wtedy, gdy klient szuka rozwiazania z okreslonym zakresem parametrow.</p>
            <h2>Zastosowanie i dopasowanie</h2>
            <ul><li>Parametry techniczne.</li><li>Dane z briefu.</li></ul>
            <h3>Specyfikacja techniczna</h3>
            <ul><li>Producent: Kanlux</li></ul>
            """,
            [],
            [],
        )
        checks_by_name = {check["check"]: check["status"] for check in checks}

        self.assertEqual(checks_by_name["missing_polish_diacritics"], "ERROR")
        self.assertEqual(checks_by_name["meta_description_language"], "ERROR")

    def test_validator_allows_concrete_product_type_context(self) -> None:
        type_context = (
            "Siatka ochronna jest akcesorium stosowanym przy oprawach, gdy potrzebna jest dodatkowa "
            "fizyczna osłona przed przypadkowym kontaktem lub uderzeniem. Przy doborze warto sprawdzić "
            "zgodność z konkretną serią i modelem oprawy, ponieważ takie elementy są dobierane do "
            "określonych konstrukcji. "
        )
        product_context = (
            "Model Kanlux 33482 należy do serii AGOR i jest wykonany z metalu. Pasuje do opraw "
            "FL AGOR HI GRID, dlatego opis kompatybilności odnosi się do oprawy, a nie do parametrów "
            "samej siatki ochronnej. "
        )
        checks = validate_codex_description(
            f"""
            <p>{product_context * 2}</p>
            <p>{type_context * 2}</p>
            <p>{product_context}{type_context}</p>
            <h2>Siatka ochronna Kanlux do FL AGOR HI GRID</h2>
            <ul><li>Metalowe wykonanie.</li><li>Dopasowanie do FL AGOR HI GRID.</li></ul>
            <h3>Specyfikacja techniczna</h3>
            <ul><li>Producent: Kanlux</li><li>Typ produktu: Siatka ochronna</li><li>Seria: AGOR</li><li>Materiał: Metal</li></ul>
            """,
            [],
            [],
        )

        self.assertFalse(any(check["status"] == "ERROR" for check in checks))

    def test_validator_requires_seo_opening_and_benefits_section(self) -> None:
        checks = validate_codex_description(
            """
            <p>Oprawa hermetyczna do pomieszczeń technicznych.</p>
            <p>Stopień ochrony IP65 wspiera zastosowanie w wymagających warunkach.</p>
            <p>Neutralna barwa światła sprawdza się podczas codziennej pracy.</p>
            <h2>Oprawa hermetyczna LED do warsztatu</h2>
            <h3>Specyfikacja techniczna</h3>
            <ul><li><strong>Stopień ochrony:</strong> IP65</li></ul>
            """,
            [{"fact": "Typ produktu", "value": "Oprawa hermetyczna", "source": "test"}],
            [],
            product_name="Oprawa hermetyczna LED Kanlux 12345",
        )
        checks_by_name = {check["check"]: check["status"] for check in checks}

        self.assertEqual(checks_by_name["seo_opening"], "WARNING")
        self.assertEqual(checks_by_name["html_benefits_h3"], "WARNING")

    def test_validator_enforces_supplied_keyword_placement_and_limit(self) -> None:
        checks = validate_codex_description(
            """
            <p><strong>Kanlux BLINGO 40W</strong> zapewnia neutralne światło.</p>
            <h2>Panel LED do sufitu kasetonowego</h2>
            <p>Ten model ma moc 40 W.</p>
            <h3>Panel LED - najważniejsze zalety</h3>
            <p>Panel LED sprawdza się we wnętrzach. Panel LED ma stalową obudowę.</p>
            <h3>Specyfikacja techniczna</h3>
            <ul><li><strong>Moc:</strong> 40 W</li></ul>
            """,
            [{"fact": "Typ produktu", "value": "Panel LED", "source": "test"}],
            [],
            product_name="Kanlux BLINGO 40W",
            seo_keyword="panel led",
        )
        checks_by_name = {check["check"]: check["status"] for check in checks}

        self.assertEqual(checks_by_name["seo_keyword_opening"], "WARNING")
        self.assertEqual(checks_by_name["seo_keyword_after_h2"], "WARNING")
        self.assertEqual(checks_by_name["seo_keyword_h3"], "WARNING")
        self.assertEqual(checks_by_name["seo_keyword_occurrences"], "WARNING")

    def test_validator_requires_closing_summary_after_specification(self) -> None:
        checks = validate_codex_description(
            """
            <p><strong>Panel LED Kanlux BLINGO</strong> daje neutralne światło do pracy.</p>
            <h2>Panel LED do sufitu kasetonowego</h2>
            <p><strong>Panel LED</strong> ma moc 40 W i strumień 3800 lm.</p>
            <h3>Najważniejsze cechy</h3>
            <ul>
              <li>Neutralna barwa światła.</li>
              <li>Podtynkowy sposób montażu.</li>
              <li>Wysoki strumień świetlny.</li>
            </ul>
            <h3>Specyfikacja techniczna</h3>
            <ul><li><strong>Moc:</strong> 40 W</li></ul>
            """,
            [{"fact": "Typ produktu", "value": "Panel LED", "source": "test"}],
            [],
            product_name="Panel LED Kanlux BLINGO",
        )

        checks_by_name = {check["check"]: check["status"] for check in checks}
        self.assertEqual(checks_by_name["html_closing_summary"], "WARNING")

    def test_validator_detects_template_language_from_bulk_generator(self) -> None:
        repeated = (
            "Stopień ochrony IP65 porządkuje dobór do przestrzeni technicznych. "
            "Kod producenta pomaga zachować porządek przy zamówieniu, a parametry tworzą "
            "techniczny profil wariantu. "
        )
        checks = validate_codex_description(
            f"""
            <p><strong>Oprawa hermetyczna LED Kanlux 12345</strong> {repeated * 3}</p>
            <p>{repeated * 3}</p>
            <p>{repeated * 3}</p>
            <h2>Oprawa hermetyczna LED IP65</h2>
            <h3>Najważniejsze cechy</h3>
            <ul><li>Stopień ochrony IP65.</li></ul>
            <h3>Specyfikacja techniczna</h3>
            <ul><li><strong>Stopień ochrony:</strong> IP65</li></ul>
            """,
            [{"fact": "Typ produktu", "value": "Oprawa hermetyczna", "source": "test"}],
            [],
            product_name="Oprawa hermetyczna LED Kanlux 12345",
        )

        self.assertTrue(
            any(check["status"] == "ERROR" and check["check"] == "meta_description_language" for check in checks)
        )

    def test_validator_detects_robotic_database_language(self) -> None:
        checks = validate_codex_description(
            """
            <p><strong>Plafon LED Kanlux 12345</strong> został przygotowany do wnętrz,
            w których liczy się wygodne oświetlenie. Klasa IP54 wspiera zastosowanie
            w miejscach o podwyższonej wilgotności.</p>
            <p>Seria produktu pomaga utrzymać spójność wizualną, a napięcie ułatwia
            dopasowanie modelu do instalacji.</p>
            <p>Gwarancja uzupełnia zestaw informacji ważnych przy wyborze.</p>
            <h2>Plafon LED do wnętrz</h2>
            <h3>Najważniejsze cechy</h3>
            <ul><li>Klasa IP54.</li></ul>
            <h3>Specyfikacja techniczna</h3>
            <ul><li><strong>Stopień ochrony:</strong> IP54</li></ul>
            """,
            [{"fact": "Typ produktu", "value": "Plafon", "source": "test"}],
            [],
            product_name="Plafon LED Kanlux 12345",
        )

        self.assertTrue(
            any(check["status"] == "WARNING" and check["check"] == "robotic_language" for check in checks)
        )

    def test_validator_allows_direct_human_language(self) -> None:
        paragraph = (
            "Plafon ma klasę IP54, dlatego jego obudowa ogranicza wnikanie pyłu i wody. "
            "Neutralna barwa 4000 K daje naturalne światło odpowiednie do korytarza, "
            "klatki schodowej lub pomieszczenia pomocniczego. "
        )
        checks = validate_codex_description(
            f"""
            <p><strong>Plafon LED Kanlux 12345</strong> {paragraph * 2}</p>
            <p>{paragraph * 2}</p>
            <p>{paragraph * 2}</p>
            <h2>Plafon LED IP54 z neutralnym światłem</h2>
            <h3>Najważniejsze cechy</h3>
            <ul><li>Obudowa IP54 ogranicza wnikanie pyłu i wody.</li></ul>
            <h3>Specyfikacja techniczna</h3>
            <ul><li><strong>Stopień ochrony:</strong> IP54</li></ul>
            """,
            [{"fact": "Typ produktu", "value": "Plafon", "source": "test"}],
            [],
            product_name="Plafon LED Kanlux 12345",
        )

        self.assertFalse(any(check["check"] == "robotic_language" for check in checks))

    def test_codex_brief_report_contains_required_sheets_and_does_not_modify_input(self) -> None:
        df = pd.DataFrame(
            [
                {
                    "Kod": "33482",
                    "new_title": "Siatka ochronna do FL AGOR HI GRID 240W Kanlux 33482",
                    "attr_typ": "Siatka ochronna",
                    "attr_moc": "240W",
                    "attr_pasuje_do": "do FL AGOR HI GRID",
                    "primary_keyword": "siatka ochronna",
                }
            ]
        )
        original = df.copy(deep=True)
        brief = build_codex_brief_for_sku(df, "33482/KAN", "new_title", {})

        with tempfile.TemporaryDirectory() as tmp:
            output = Path(tmp) / "brief.xlsx"
            write_codex_brief_report(brief, output)
            with pd.ExcelFile(output) as xls:
                sheets = xls.sheet_names
                prompt_header = pd.read_excel(xls, sheet_name="Prompt dla Codex", nrows=0).columns[0]

        self.assertEqual(sheets, CODEX_BRIEF_SHEETS)
        self.assertEqual(prompt_header, "Prompt dla agenta Codex")
        self.assertIn("Nie twórz skryptu", brief["prompt"])
        self.assertIn('"<brief.xlsx>"', brief["prompt"])
        self.assertIn("good_description_arot.html", brief["prompt"])
        self.assertIn("SEO_KEYWORD: siatka ochronna", brief["prompt"])
        self.assertEqual(brief["product"]["seo_keyword"], "siatka ochronna")
        pd.testing.assert_frame_equal(df, original)

    def test_validate_codex_review_file_detects_bad_generated_description(self) -> None:
        brief = {
            "product": {"sku": "33482", "requested_sku": "33482/KAN"},
            "product_facts": [{"fact": "Materiał", "value": "Metal", "source": "filtered_baselinker_features"}],
            "compatibility_facts": [{"fact": "Pasuje do", "value": "do FL AGOR HI GRID", "source": "compatibility_source"}],
            "rejected_facts": [],
            "prompt": "test",
            "generated_description": "<p>Opis zabezpiecza? produkt.</p><ul><li>- zly punkt</li></ul>",
            "validation": [],
        }

        with tempfile.TemporaryDirectory() as tmp:
            output = Path(tmp) / "review.xlsx"
            write_codex_brief_report(brief, output)
            checks = validate_codex_review_file(output)

        checks_by_name = {check["check"]: check["status"] for check in checks}
        self.assertEqual(checks_by_name["encoding_integrity"], "ERROR")
        self.assertEqual(checks_by_name["html_specification_h3"], "WARNING")

    def test_write_codex_review_from_description_writes_valid_review(self) -> None:
        paragraph = (
            "Siatka ochronna Kanlux 33482 jest metalowym akcesorium przeznaczonym do osłony "
            "kompatybilnej oprawy z serii AGOR. Fizyczna bariera ogranicza ryzyko przypadkowego "
            "kontaktu z oprawą, a dokładne oznaczenie modelu wspiera dopasowanie części. "
        )
        description = "\n".join(
            [
                f"<p><strong>Siatka ochronna do FL AGOR HI GRID Kanlux 33482</strong> {paragraph * 2}</p>",
                f"<p>{paragraph * 2}</p>",
                f"<p>{paragraph * 2}</p>",
                "<h2>Siatka ochronna do opraw FL AGOR HI GRID</h2>",
                "<h3>Najważniejsze cechy</h3>",
                "<ul>",
                "<li>Metalowe wykonanie tworzy fizyczną osłonę oprawy.</li>",
                "<li>Kompatybilność z FL AGOR HI GRID pomaga dobrać właściwy element do oprawy.</li>",
                "<li>Seria AGOR pozwala zachować spójność między pasującymi komponentami.</li>",
                "</ul>",
                "<h3>Specyfikacja techniczna</h3>",
                "<ul>",
                "<li><strong>Kod producenta:</strong> 33482</li>",
                "<li><strong>Producent:</strong> Kanlux</li>",
                "<li><strong>Typ produktu:</strong> Siatka ochronna</li>",
                "<li><strong>Seria:</strong> AGOR</li>",
                "<li><strong>Materiał:</strong> Metal</li>",
                "</ul>",
                (
                    "<p>Metalowa siatka Kanlux 33482 jest właściwym wyborem, gdy potrzebna jest "
                    "potwierdzona kompatybilność z oprawą FL AGOR HI GRID i dodatkowa fizyczna "
                    "osłona. Wybierz ten model do wskazanej serii opraw.</p>"
                ),
            ]
        )
        brief = {
            "product": {
                "sku": "33482",
                "requested_sku": "33482/KAN",
                "name": "Siatka ochronna do FL AGOR HI GRID Kanlux 33482",
            },
            "product_facts": [
                {"fact": "Kod producenta", "value": "33482", "source": "product_identity"},
                {"fact": "Producent", "value": "Kanlux", "source": "filtered_baselinker_features"},
                {"fact": "Typ produktu", "value": "Siatka ochronna", "source": "filtered_baselinker_features"},
                {"fact": "Seria", "value": "AGOR", "source": "filtered_baselinker_features"},
                {"fact": "Materiał", "value": "Metal", "source": "filtered_baselinker_features"},
            ],
            "compatibility_facts": [{"fact": "Pasuje do", "value": "do FL AGOR HI GRID", "source": "compatibility_source"}],
            "rejected_facts": [],
            "prompt": "test",
            "generated_description": "",
            "validation": [],
        }

        with tempfile.TemporaryDirectory() as tmp:
            brief_path = Path(tmp) / "brief.xlsx"
            review_path = Path(tmp) / "review.xlsx"
            write_codex_brief_report(brief, brief_path)
            checks = write_codex_review_from_description(brief_path, review_path, description)
            review_checks = validate_codex_review_file(review_path)
            self.assertTrue(review_path.exists())

        self.assertFalse(any(check["status"] in {"ERROR", "WARNING"} for check in checks))
        self.assertFalse(any(check["status"] in {"ERROR", "WARNING"} for check in review_checks))


if __name__ == "__main__":
    unittest.main()
