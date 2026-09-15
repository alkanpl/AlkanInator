from __future__ import annotations

import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from build_hager_berker_cross_upsell import (  # noqa: E402
    CatalogProduct,
    classify_function,
    cross_candidate,
    select_recommendations,
    up_sell_candidate,
    validate_recommendations,
)
from build_hager_berker_variants import Record  # noqa: E402


def product_record(
    sku: str,
    title: str,
    product_type: str,
    series: str = "B. Kwadrat",
    color: str = "Biały",
    core: str = "",
    multiplicity: str = "",
    tags: str = "",
) -> Record:
    return Record(
        excel_row=2,
        product_id=sku.split("/", 1)[0],
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
        tags_before=tags,
        core=core or title,
        family="ramki" if product_type == "Ramki" else "inne",
        frame=product_type == "Ramki",
        color_detail=color,
        material_detail="",
        finish_detail="",
        orientation="Bez orientacji" if product_type == "Ramki" else "",
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
        alias_code="",
    )


class CrossSellTests(unittest.TestCase):
    def test_frame_and_key_with_same_series_and_color_are_cross_sell(self) -> None:
        frame = CatalogProduct(product_record(
            "5310118989/HAG",
            "Ramka 1-krotna biała 5310118989 Berker B.Kwadrat",
            "Ramki",
            multiplicity="1x",
        ))
        key = CatalogProduct(product_record(
            "5316208999/HAG",
            "Klawisz pojedynczy biały 5316208999 Berker B.Kwadrat",
            "Klawisze",
            core="klawisz pojedynczy",
        ))

        recommendation = cross_candidate(key, frame)

        self.assertIsNotNone(recommendation)
        self.assertEqual(recommendation.kind, "Cross-sell")
        self.assertIn("zgodna seria", recommendation.reason)

    def test_different_colors_are_not_linked(self) -> None:
        frame = CatalogProduct(product_record(
            "FRAME/HAG", "Ramka 1-krotna biała", "Ramki", multiplicity="1x",
        ))
        key = CatalogProduct(product_record(
            "KEY/HAG", "Klawisz pojedynczy antracyt", "Klawisze", color="Antracyt",
        ))

        self.assertIsNone(cross_candidate(key, frame))

    def test_products_from_same_variant_group_are_not_recommended(self) -> None:
        tag = "Wariant - Klawisz pojedynczy"
        frame = CatalogProduct(product_record(
            "FRAME/HAG", "Ramka 1-krotna biała", "Ramki", multiplicity="1x", tags=tag,
        ))
        key = CatalogProduct(product_record(
            "KEY/HAG", "Klawisz pojedynczy biały", "Klawisze", tags=tag,
        ))

        self.assertIsNone(cross_candidate(key, frame))

    def test_double_key_is_not_linked_to_timer_mechanism(self) -> None:
        double_key_record = product_record(
            "DOUBLE/HAG", "Klawisze kremowe", "Klawisze", multiplicity="2x",
        )
        timer_record = product_record(
            "TIMER/HAG", "Mechaniczny łącznik czasowy 0-120 min, mechanizm", "Mechanizmy",
            series="one.platform", color="",
        )
        double_key = CatalogProduct(double_key_record)
        timer = CatalogProduct(timer_record)

        self.assertEqual(classify_function(double_key_record), "double_switch")
        self.assertEqual(classify_function(timer_record), "timer_switch")
        self.assertIsNone(cross_candidate(double_key, timer))


class UpSellTests(unittest.TestCase):
    def test_illuminated_key_is_directional_up_sell(self) -> None:
        standard = CatalogProduct(product_record(
            "5316208999/HAG",
            "Klawisz pojedynczy biały 5316208999 Berker B.Kwadrat",
            "Klawisze",
            core="klawisz pojedynczy",
        ))
        illuminated = CatalogProduct(product_record(
            "5316218999/HAG",
            "Klawisz pojedynczy podświetlany biały 5316218999 Berker B.Kwadrat",
            "Klawisze",
            core="klawisz pojedynczy podświetlany",
        ))

        upgrade = up_sell_candidate(standard, illuminated)
        downgrade = up_sell_candidate(illuminated, standard)

        self.assertIsNotNone(upgrade)
        self.assertIn("podświetlenie", upgrade.reason)
        self.assertIsNone(downgrade)

    def test_different_key_multiplicity_is_not_an_up_sell(self) -> None:
        triple = CatalogProduct(product_record(
            "TRIPLE/HAG",
            "Klawisz potrójny aluminium",
            "Klawisze",
            core="klawisz",
            multiplicity="3x",
        ))
        illuminated_single = CatalogProduct(product_record(
            "SINGLE/HAG",
            "Klawisz pojedynczy podświetlany aluminium",
            "Klawisze",
            core="klawisz podświetlany",
            multiplicity="1x",
        ))

        self.assertIsNone(up_sell_candidate(triple, illuminated_single))

    def test_attribute_multiplicity_with_x_is_respected(self) -> None:
        double = CatalogProduct(product_record(
            "DOUBLE/HAG",
            "Klawisze kremowe",
            "Klawisze",
            core="klawisz",
            multiplicity="2x",
        ))
        illuminated_single = CatalogProduct(product_record(
            "SINGLE/HAG",
            "Klawisz pojedynczy podświetlany kremowy",
            "Klawisze",
            core="klawisz podświetlany",
            multiplicity="1x",
        ))

        self.assertEqual(double.configuration_count, 2)
        self.assertIsNone(up_sell_candidate(double, illuminated_single))

    def test_selected_cross_and_up_sell_are_disjoint_and_valid(self) -> None:
        products = [
            CatalogProduct(product_record(
                "FRAME/HAG", "Ramka 1-krotna biała", "Ramki", multiplicity="1x",
            )),
            CatalogProduct(product_record(
                "KEY/HAG", "Klawisz pojedynczy biały", "Klawisze", core="klawisz pojedynczy",
            )),
            CatalogProduct(product_record(
                "KEY-LIGHT/HAG", "Klawisz pojedynczy podświetlany biały", "Klawisze",
                core="klawisz pojedynczy podświetlany",
            )),
        ]

        cross, up = select_recommendations(products)
        validation = validate_recommendations(products, cross, up, 4, 3)

        self.assertTrue(all(validation.values()))
        self.assertIn("KEY/HAG", cross)
        self.assertIn("KEY/HAG", up)
        self.assertTrue(
            {item.target_sku for item in cross["KEY/HAG"]}.isdisjoint(
                item.target_sku for item in up["KEY/HAG"]
            )
        )


if __name__ == "__main__":
    unittest.main()
