"""Testy charakteryzujace dla ekstrakcji atrybutow z tytulow.

Czesc 1 (golden-file): porownuje wynik ``extract_attributes_from_text`` z
zamrozonym snapshotem na 40 realnych tytulach. Kazda niezamierzona zmiana
ktoregokolwiek regexa od razu wywala test. Po SWIADOMEJ zmianie regul
zregeneruj snapshot:  py tests/generate_snapshots.py

Czesc 2 (jawne przypadki): czytelne asercje dokumentujace intencje dla
najtrudniejszych normalizacji (zakresy mocy, CCT, strumienia, gwinty, czujnik).
"""

from __future__ import annotations

import io
import json
import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from extract_attributes import extract_attributes_from_text  # noqa: E402

FIXTURE = Path(__file__).resolve().parent / "fixtures" / "extraction_snapshot.json"


def attrs(title: str) -> dict[str, str]:
    return extract_attributes_from_text(title)["attributes"]


class TestExtractionSnapshot(unittest.TestCase):
    """Golden-file: aktualne zachowanie na realnych tytulach Kanlux."""

    @classmethod
    def setUpClass(cls) -> None:
        with io.open(FIXTURE, encoding="utf-8") as handle:
            cls.cases = json.load(handle)

    def test_snapshot_not_empty(self) -> None:
        self.assertGreaterEqual(len(self.cases), 30)

    def test_extraction_matches_snapshot(self) -> None:
        mismatches = []
        for case in self.cases:
            got = attrs(case["title"])
            if got != case["attributes"]:
                mismatches.append((case["title"], case["attributes"], got))
        if mismatches:
            lines = [
                f"\n  TITLE: {title}\n  EXPECTED: {exp}\n  GOT:      {got}"
                for title, exp, got in mismatches
            ]
            self.fail(
                f"{len(mismatches)} tytul(ow) zmienilo ekstrakcje. Jesli to "
                f"zamierzone, uruchom: py tests/generate_snapshots.py"
                + "".join(lines)
            )


class TestExtractionExplicit(unittest.TestCase):
    """Jawne przypadki dokumentujace intencje normalizacji."""

    def test_basic_panel(self) -> None:
        result = attrs("Panel LED 40W 4000K 4000lm IP20")
        self.assertEqual(result["moc"], "40W")
        self.assertEqual(result["barwa"], "4000K")
        self.assertEqual(result["strumien"], "4000lm")
        self.assertEqual(result["ip"], "IP20")

    def test_power_range_normalized_to_min_max(self) -> None:
        # "12-18W" ma zostac zachowane jako zakres
        self.assertEqual(attrs("Plafoniera BENO 12-18W CCT")["moc"], "12-18W")

    def test_cct_range_collapses_to_min_max(self) -> None:
        result = attrs("BENO barwa 3000/3500/4000K")
        self.assertEqual(result["barwa_zakres"], "3000-4000K")

    def test_flux_range_collapses_to_min_max(self) -> None:
        # "17000 / 12750 / 8500Lm" -> najmniejszy-najwiekszy
        result = attrs("HIGH BAY 17000 / 12750 / 8500Lm")
        self.assertEqual(result["strumien"], "8500-17000lm")

    def test_socket_gu10_with_count(self) -> None:
        self.assertEqual(attrs("Oprawa punktowa BLURRO 2XGU10 CO-B")["gwint"], "2xGU10")

    def test_socket_e27(self) -> None:
        self.assertEqual(attrs("Plafoniera JASMIN 270-B 1XE27")["gwint"], "1xE27")

    def test_motion_sensor_detected(self) -> None:
        self.assertEqual(
            attrs("VARSO LED 18W-NW-O-SE")["czujnik"], "z czujnikiem ruchu"
        )

    def test_lm_per_w_efficiency(self) -> None:
        self.assertEqual(attrs("STIVI 24W 3120lm 130lm/W")["lm_w"], "130lm/W")

    def test_ip_with_space_normalized(self) -> None:
        # "IP 65" (ze spacja) -> "IP65"
        self.assertEqual(attrs("Oprawa HB PRO STRONG IP 65")["ip"], "IP65")

    def test_unusual_power_led_dash(self) -> None:
        # nietypowy zapis mocy "LED-10-B" -> 10W
        self.assertEqual(attrs("Naswietlacz GRUN NV LED-10-B")["moc"], "10W")

    def test_no_false_power_from_plain_text(self) -> None:
        result = attrs("Pilot do oprawy HB PRO STRONG REMOTE")
        self.assertNotIn("moc", result)


if __name__ == "__main__":
    unittest.main()