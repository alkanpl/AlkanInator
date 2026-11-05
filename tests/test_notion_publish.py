from __future__ import annotations

import importlib.util
from pathlib import Path
from unittest import TestCase, mock

ROOT = Path(__file__).resolve().parents[1]
SPEC = importlib.util.spec_from_file_location("publikuj_w_notion", ROOT / "src" / "notion_api" / "publikuj-w-notion.py")
publikuj = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(publikuj)

DOKUMENTY = {"docs/changelog.md": "a" * 32, "docs/instrukcja.md": ""}


class PrzygotujTests(TestCase):
    def setUp(self):
        patcher = mock.patch.dict(publikuj.DOKUMENTY, DOKUMENTY, clear=True)
        patcher.start()
        self.addCleanup(patcher.stop)

    def test_tytul_z_h1_bez_naglowka_w_tresci(self):
        tytul, tresc = publikuj.przygotuj("docs/instrukcja.md", "# AlkanInator – instrukcja\n\nTekst.\n", "abcdef1234", "2026-11-04")
        self.assertEqual(tytul, "AlkanInator – instrukcja")
        self.assertNotIn("# AlkanInator", tresc)
        self.assertIn("repozytorium AlkanInator", tresc)
        self.assertIn("commit `abcdef1`, 2026-11-04", tresc)

    def test_link_do_opublikowanego_dokumentu_prowadzi_do_notion(self):
        _, tresc = publikuj.przygotuj("docs/instrukcja.md", "# T\n\n[Zmiany](changelog.md)\n", "abcdef1", "2026-11-04")
        self.assertIn("[Zmiany](https://app.notion.com/p/%s)" % ("a" * 32), tresc)

    def test_link_do_dokumentu_bez_strony_prowadzi_na_github(self):
        _, tresc = publikuj.przygotuj("docs/changelog.md", "# T\n\n[instrukcja](instrukcja.md)\n", "abcdef1", "2026-11-04")
        self.assertIn("[instrukcja](https://github.com/alkanpl/AlkanInator/blob/main/docs/instrukcja.md)", tresc)

    def test_tabela_i_znaki_specjalne(self):
        md = "# T\n\n| Status | Co znaczy |\n|---|---|\n| OK | `a|b` i 1000×1000 > 0 |\n"
        _, tresc = publikuj.przygotuj("docs/instrukcja.md", md, "abcdef1", "2026-11-04")
        self.assertIn('<table fit-page-width="true" header-row="true">', tresc)
        self.assertIn("<td>`a|b` i 1000×1000 \\> 0</td>", tresc)

    def test_cytat_zostaje_cytatem(self):
        _, tresc = publikuj.przygotuj("docs/instrukcja.md", "# T\n\n> Zlecenie: 4000K > 3000K\n", "abcdef1", "2026-11-04")
        self.assertIn("> Zlecenie: 4000K \\> 3000K", tresc)

    def test_skrot_m_in_nie_staje_sie_linkiem(self):
        _, tresc = publikuj.przygotuj("docs/changelog.md", "# T\n\ndoszli m.in. Legrand, `m.in.` w kodzie\n", "abcdef1", "2026-11-04")
        self.assertIn("doszli m.⁠in.⁠ Legrand", tresc)
        self.assertIn("`m.in.` w kodzie", tresc)

    def test_dokumenty_projektu_maja_naglowek_h1(self):
        for sciezka in publikuj.DOKUMENTY:
            tekst = (ROOT / sciezka).read_text(encoding="utf-8")
            tytul, _ = publikuj.przygotuj(sciezka, tekst, "abcdef1", "2026-11-04")
            self.assertTrue(tytul)
