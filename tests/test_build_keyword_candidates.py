import sys
import tempfile
import unittest
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from build_keyword_candidates import (  # noqa: E402
    assign_products,
    assign_type,
    attach_planner_stats,
    candidate_rows,
    load_planner_stats,
    load_product_types,
    type_summary_rows,
    write_reports,
)
from utils import load_yaml  # noqa: E402


CONFIG = {
    "product_types": {
        "fundament": {
            "name": "Fundament do obudowy",
            "match": {"starts_with": ["Fundament"]},
            "keywords": {"katalogowa": ["fundament do obudowy"], "marka_kod": ["incobex"]},
        },
        "fundament_termo": {
            "name": "Fundament termoutwardzalny",
            "match": {"starts_with": ["Fundament termoutwardzalny"]},
            "keywords": {"katalogowa": ["Fundament  termoutwardzalny"], "synonimy": ["cokół do obudowy"]},
        },
        "plyta": {
            "name": "Płyta izolacyjna",
            "match": {"starts_with": ["Płyta izolacyjna"]},
            "keywords": {
                "katalogowa": ["płyta izolacyjna", "płyta montażowa izolacyjna"],
                "synonimy": ["płyta montażowa pcv"],
                "marka_kod": ["Incobex"],
            },
            "ambiguous": {"Płyta izolacyjna": "izolacja budynków"},
        },
        "szyna": {
            "name": "Szyna PEN",
            "match": {"starts_with": ["Szyna PEN"]},
            "keywords": {"katalogowa": ["szyna pen"]},
        },
    }
}


class BuildKeywordCandidatesTest(unittest.TestCase):
    def test_longest_prefix_wins_and_accents_are_ignored(self):
        types = load_product_types(CONFIG)
        self.assertEqual(assign_type("Fundament termoutwardzalny FTN 26/32", types).key, "fundament_termo")
        self.assertEqual(assign_type("Fundament FTN-26 do obudowy STN-26", types).key, "fundament")
        self.assertEqual(assign_type("Plyta izolacyjna gr. 4mm PMN 23x36", types).key, "plyta")

    def test_prefix_must_end_on_word_boundary(self):
        types = load_product_types(CONFIG)
        self.assertIsNone(assign_type("Fundamentowa kotwa", types))

    def test_unassigned_products_are_reported_not_forced(self):
        types = load_product_types(CONFIG)
        df = pd.DataFrame(
            {
                "SKU": ["A/INC", "B/INC", "C/INC"],
                "Title": ["Płyta izolacyjna PMN 23x36", "Wkładka do zamka WRS-K", "Fundament FTN-40"],
            }
        )
        unassigned = assign_products(df, types)
        self.assertEqual([p["SKU"] for p in unassigned], ["B/INC"])
        self.assertEqual(len(types[0].products), 1)

    def test_candidates_skip_types_without_products_and_deduplicate(self):
        types = load_product_types(CONFIG)
        assign_products(pd.DataFrame({"Title": ["Fundament FTN-40", "Fundament termoutwardzalny FTN 26", "Płyta izolacyjna PMN"]}), types)
        rows = candidate_rows(types)
        keywords = [row["keyword"] for row in rows]
        self.assertIn("fundament termoutwardzalny", keywords)  # spacje sciagniete, male litery
        self.assertEqual(keywords.count("incobex"), 1)  # duplikat z typu "plyta" pominiety
        self.assertNotIn("szyna pen", keywords)  # typ bez produktow
        self.assertEqual(next(r for r in rows if r["keyword"] == "incobex")["product_type"], "Fundament do obudowy")

    def test_reports_are_written(self):
        types = load_product_types(CONFIG)
        unassigned = assign_products(pd.DataFrame({"Title": ["Szyna PEN N-80", "Coś innego"]}), types)
        with tempfile.TemporaryDirectory() as tmp:
            paths = write_reports(Path(tmp), types, candidate_rows(types), unassigned)
            self.assertEqual(paths["seed"].read_bytes(), b"Keyword\nszyna pen\n")  # szablon Plannera, bez BOM
            sheets = pd.read_excel(paths["xlsx"], sheet_name=None)
            self.assertEqual(len(sheets["Nieprzypisane"]), 1)
            self.assertEqual(paths["csv"].read_bytes()[:3], b"\xef\xbb\xbf")

    def test_incobex_dictionary_loads(self):
        types = load_product_types(load_yaml("dictionaries/incobex_keyword_candidates.yaml"))
        self.assertGreaterEqual(len(types), 8)
        for product_type in types:
            self.assertTrue(product_type.prefixes, product_type.key)
            self.assertTrue(product_type.keywords["katalogowa"], product_type.key)


    def test_planner_export_utf16_tab_is_parsed(self):
        lines = [
            "Keyword Stats 2026-09-30 at 10_18_53",
            '"September 1, 2025 - August 31, 2026"',
            "Keyword\tCurrency\tSegmentation\tAvg. monthly searches\tThree month change\tYoY change"
            "\tCompetition\tCompetition (indexed value)\tTop of page bid (low range)\tTop of page bid (high range)",
            "\t\tAll\t43700.0\t\t\t\t\t\t",
            "płyta izolacyjna\tPLN\t\t500.0\t0%\t0%\tHigh\t100\t0.49\t2.50",
            "płyta montażowa pcv\tPLN\t\t\t\t\t\t\t\t",
            "płyta montażowa izolacyjna\tPLN\t\t50.0\t0%\t0%\tHigh\t100\t\t",
            "incobex\tPLN\t\t5000.0\t0%\t0%\tHigh\t68\t0.76\t3.11",
            "coś spoza listy\tPLN\t\t50.0\t\t\t\t\t\t",
        ]
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "stats.csv"
            path.write_bytes("\n".join(lines).encode("utf-16"))
            stats = load_planner_stats(path)
        self.assertEqual(len(stats), 5)  # wiersze podsumowania bez frazy pominiete
        self.assertEqual(stats["plyta izolacyjna"]["avg_monthly_searches"], 500)
        self.assertEqual(stats["plyta izolacyjna"]["bid_high_pln"], 2.5)
        self.assertIsNone(stats["plyta montazowa pcv"]["avg_monthly_searches"])

        types = load_product_types(CONFIG)
        assign_products(pd.DataFrame({"Title": ["Płyta izolacyjna PMN 23x36"]}), types)
        candidates = candidate_rows(types)
        extra = attach_planner_stats(candidates, types, stats)
        by_keyword = {row["keyword"]: row for row in candidates}
        self.assertEqual(by_keyword["płyta izolacyjna"]["status"], "NIEJEDNOZNACZNA")
        self.assertEqual(by_keyword["płyta montażowa izolacyjna"]["searches_range"], "10-100")
        self.assertEqual(by_keyword["płyta montażowa pcv"]["status"], "BRAK_WOLUMENU")
        self.assertEqual(by_keyword["incobex"]["status"], "OK")
        self.assertEqual([row["keyword"] for row in extra], ["coś spoza listy"])

        summary = next(row for row in type_summary_rows(types, candidates) if row["products"])
        # fraza niejednoznaczna i fraza marki nie wygrywaja mimo wiekszego wolumenu
        self.assertEqual(summary["primary_keyword"], "płyta montażowa izolacyjna")
        self.assertEqual(summary["ambiguous"], "płyta izolacyjna (100-1 tys.)")

    def test_missing_keyword_header_is_rejected(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "forecast.csv"
            path.write_bytes("Date Range\tCampaign\tKeyword\n".encode("utf-16"))
            with self.assertRaises(SystemExit):
                load_planner_stats(path)


if __name__ == "__main__":
    unittest.main()
