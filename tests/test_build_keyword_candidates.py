import sys
import tempfile
import unittest
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from build_keyword_candidates import (  # noqa: E402
    assign_products,
    assign_type,
    candidate_rows,
    load_product_types,
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
            "keywords": {"katalogowa": ["płyta izolacyjna"], "marka_kod": ["Incobex"]},
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


if __name__ == "__main__":
    unittest.main()
