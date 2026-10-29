from __future__ import annotations

import re
from pathlib import Path
from unittest import TestCase

import yaml


PATTERNS_PATH = Path(__file__).resolve().parents[1] / "configs" / "title_patterns.yaml"
FIELD = re.compile(r"{([^{}]+)}")
REQUIRED_KEYS = {"id", "producer", "category", "examples", "source"}


def render(template: str, values: dict[str, object]) -> str:
    return FIELD.sub(lambda match: str(values[match.group(1)]), template)


class TitlePatternsTests(TestCase):
    """Wspolne wzory tytulow (configs/title_patterns.yaml): przyklady musza wynikac ze wzoru."""

    @classmethod
    def setUpClass(cls) -> None:
        cls.config = yaml.safe_load(PATTERNS_PATH.read_text(encoding="utf-8"))
        cls.families = cls.config["families"]

    def test_families_have_required_keys_and_unique_ids(self) -> None:
        ids = [family.get("id") for family in self.families]
        self.assertEqual(len(ids), len(set(ids)), "powtorzone id rodziny")
        for family in self.families:
            with self.subTest(family=family.get("id")):
                self.assertFalse(REQUIRED_KEYS - set(family), f"brak kluczy: {REQUIRED_KEYS - set(family)}")
                self.assertRegex(family["id"], r"^[a-z0-9_]+$")
                self.assertTrue(family["examples"], "rodzina bez przykladow")

    def test_examples_follow_template(self) -> None:
        for family in self.families:
            template = family.get("template")
            for example in family["examples"]:
                with self.subTest(family=family["id"], title=example["title"]):
                    if not template:
                        self.assertNotIn("values", example, "values bez template")
                        continue
                    values = example.get("values") or {}
                    fields = set(FIELD.findall(template))
                    self.assertEqual(fields, set(values), "pola wzoru i values sie roznia")
                    self.assertEqual(render(template, values), example["title"])

    def test_titles_are_unique_and_clean(self) -> None:
        titles = [example["title"] for family in self.families for example in family["examples"]]
        self.assertEqual(len(titles), len(set(titles)), "powtorzony tytul w przykladach")
        for title in titles:
            with self.subTest(title=title):
                self.assertEqual(title, " ".join(title.split()), "podwojne albo skrajne spacje")
                self.assertNotIn("mm²", title, "w tytule mm2 zamiast mm²")
