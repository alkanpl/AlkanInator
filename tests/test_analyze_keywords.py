from __future__ import annotations

import json
import tempfile
import unittest
from pathlib import Path

import pandas as pd

import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from analyze_keywords import (  # noqa: E402
    KeywordMetric,
    OfflineKeywordProvider,
    analyze_keywords_for_dataframe,
    customer_id_from_google_ads_config,
    keyword_allowed,
    normalize_product_type,
    rank_keyword_metrics,
    select_candidate_keywords,
)


class MockKeywordProvider:
    source_name = "mock"

    def __init__(self) -> None:
        self.ideas = {
            "Panel LED": [
                "panel led",
                "panel led sufitowy",
                "kanlux panel led",
                "panel 5901234567890",
                "oprawa panelowa led",
                "panel led 60x60",
            ]
        }
        self.metrics = {
            "panel led": KeywordMetric("panel led", 100),
            "panel led sufitowy": KeywordMetric("panel led sufitowy", 300),
            "oprawa panelowa led": KeywordMetric("oprawa panelowa led", 300),
            "panel led 60x60": KeywordMetric("panel led 60x60", 200),
        }

    def generate_ideas(self, product_type: str) -> list[str]:
        return self.ideas.get(product_type, [])

    def historical_metrics(self, keywords: list[str]) -> dict[str, KeywordMetric]:
        return {key: value for key, value in self.metrics.items() if value.keyword in keywords}


class AnalyzeKeywordsTest(unittest.TestCase):
    def test_normalize_product_type(self) -> None:
        self.assertEqual(normalize_product_type("  Panel LED  "), "Panel LED")
        self.assertEqual(normalize_product_type(""), "")
        self.assertEqual(normalize_product_type(None), "")

    def test_keyword_filter_blocks_brands_codes_and_ean(self) -> None:
        self.assertFalse(keyword_allowed("kanlux panel led", {"Kanlux"}, set()))
        self.assertFalse(keyword_allowed("panel 27163", set(), {"27163"}))
        self.assertFalse(keyword_allowed("panel 5901234567890", set(), set()))
        self.assertTrue(keyword_allowed("panel led sufitowy", {"Kanlux"}, {"27163"}))

    def test_select_candidate_keywords_keeps_five_relevant(self) -> None:
        ideas = [
            "panel led",
            "Panel LED",
            "kanlux panel led",
            "panel led sufitowy",
            "oprawa panelowa led",
            "panel led 60x60",
            "panel do sufitu",
            "panel przemyslowy",
        ]
        selected = select_candidate_keywords("Panel LED", ideas, {"Kanlux"}, set(), 5)
        self.assertEqual(len(selected), 5)
        self.assertEqual(selected[0], "panel led")
        self.assertNotIn("kanlux panel led", selected)

    def test_rank_keyword_metrics_uses_searches_then_relevance(self) -> None:
        candidates = ["oprawa panelowa led", "panel led sufitowy", "panel led"]
        metrics = {
            "oprawa panelowa led": KeywordMetric("oprawa panelowa led", 300),
            "panel led sufitowy": KeywordMetric("panel led sufitowy", 300),
            "panel led": KeywordMetric("panel led", 100),
        }
        ranked = rank_keyword_metrics("Panel LED", candidates, metrics)
        self.assertEqual([item.keyword for item in ranked], ["panel led sufitowy", "oprawa panelowa led", "panel led"])

    def test_merge_results_to_products_by_type(self) -> None:
        df = pd.DataFrame(
            [
                {"SKU": "A1", "attr_typ": "Panel LED", "Producent": "Kanlux"},
                {"SKU": "A2", "attr_typ": "Panel LED", "Producent": "Kanlux"},
                {"SKU": "A3", "attr_typ": "", "Producent": "Kanlux"},
            ]
        )
        result, summary, missing, errors = analyze_keywords_for_dataframe(df, MockKeywordProvider())
        self.assertEqual(result.loc[0, "primary_keyword"], "panel led sufitowy")
        self.assertEqual(result.loc[1, "secondary_keyword"], "oprawa panelowa led")
        self.assertEqual(result.loc[2, "keyword_status"], "NO_TYPE")
        self.assertEqual(len(summary), 1)
        self.assertEqual(len(missing), 1)
        self.assertEqual(len(errors), 0)
        self.assertEqual(len(json.loads(result.loc[0, "keyword_candidates_json"])), 4)

    def test_offline_provider_reads_csv_metrics(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "keywords.csv"
            path.write_text(
                "product_type,keyword,avg_monthly_searches\n"
                "Panel LED,panel led,100\n"
                "Panel LED,panel led sufitowy,300\n",
                encoding="utf-8",
            )
            provider = OfflineKeywordProvider(path)
            ideas = provider.generate_ideas("Panel LED")
            metrics = provider.historical_metrics(ideas)
            self.assertEqual(ideas, ["panel led", "panel led sufitowy"])
            self.assertEqual(metrics["panel led sufitowy"].avg_monthly_searches, 300)

    def test_customer_id_can_be_read_from_google_ads_yaml(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "google-ads.yaml"
            path.write_text("customer_id: 123-456-7890\nlogin_customer_id: 9999999999\n", encoding="utf-8")
            self.assertEqual(customer_id_from_google_ads_config(str(path)), "123-456-7890")

    def test_login_customer_id_is_yaml_fallback(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "google-ads.yaml"
            path.write_text("login_customer_id: 999-888-7777\n", encoding="utf-8")
            self.assertEqual(customer_id_from_google_ads_config(str(path)), "999-888-7777")


if __name__ == "__main__":
    unittest.main()
