import os
import re
import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from build_ekonom_names import (  # noqa: E402
    add_card_cri,
    drop_lone_onoff,
    gtin_other_product,
    parse_dim,
    parse_glass,
    undo_excel_date,
    parse_deco_size,
    card_family_lines,
    card_family_name,
    card_families,
    family_from_name,
    family_phrases,
    DEFAULT_RULES,
    add_all_distinguishers,
    add_version_suffixes,
    apply_web_card,
    apply_pack_size,
    apply_package_brand,
    build_name,
    check_catalog_cct,
    full_name,
    load_web_params,
    package_from_pdfs,
    pdf_texts,
    lumen_from_window,
    lumens_from_pdf_codes,
    lumens_from_pdfs,
    move_brand_tail,
    diameter_in_mm,
    parse_ballast,
    parse_breaker,
    kx_lumen,
    load_params_web,
    parse_cct,
    parse_driver,
    parse_optics,
    parse_power,
    parse_shape,
    parse_socket,
    parse_view,
    propagate_family_brand,
    read_csv,
    split_for_ekonom,
    split_head_tail,
    strip_brand_from_names,
    write_ekonom_csv,
    EKONOM_FIELDS,
    load_pawbol_names,
    pawbol_code_key,
    load_onninen_cards,
)
from utils import load_yaml  # noqa: E402

RULES = load_yaml(DEFAULT_RULES)["suppliers"]["ledvance"]


def body(name):
    """Nazwa bez ogona "marka + linia" (Ledvance od 2026-10-19 ma marke na koncu) - tu siedza wyrozniki."""
    words = name.split()
    i = next((k for k, w in enumerate(words) if w in ("Ledvance", "Osram", "Bellalux")), len(words))
    return " ".join(words[:i])


def row(name, ean="4099854000000"):
    return {"Nazwa towaru": name[:40], "Nazwa dodatkowa": name[40:], "Kod producenta": ean}


class ParseTest(unittest.TestCase):
    def test_cct_codes(self):
        self.assertEqual(parse_cct(parse_view("Lampa LED CLASSIC A 7.5W 840 Clear E27")), "4000K")
        self.assertEqual(parse_cct(parse_view("Oprawa PANEL COMPACT 625 MULTI SELECT 30W ML 83040")), "3000/4000K")
        self.assertEqual(parse_cct(parse_view("Lampa LED Vintage 1906 GOLD 95 dim 4,8W/822 szkło E27")), "2200K")
        self.assertEqual(parse_cct(parse_view("Oprawa OFFICE LINE DESK RETRO 9W830 D BK LEDV")), "3000K")
        self.assertEqual(parse_cct(parse_view("Lampa LED SUPERSTAR CCT SELECT 4.2W 827/840/865 GU10")), "2700/4000/6500K")

    def test_lumen_in_name_is_not_cct(self):
        # "880LM" nie jest kodem barwy 8000K
        self.assertEqual(parse_cct(parse_view("Oprawa DOWNLIGHT IP44 GEN2 DN 115 8W 830 880LM BK")), "3000K")

    def test_glued_osram_codes(self):
        u = parse_view("Lampa LED SMARTWIFI B40 4,9W/827230VDIMFRE144X3LEDV")
        self.assertEqual(parse_socket(u), "E14")
        self.assertEqual(parse_cct(u), "2700K")
        self.assertEqual(parse_socket(parse_view("Lampa DULUX LED T/E26 HF V 10W 830 GX24Q.-3LEDV")), "GX24q-3")

    def test_shape(self):
        self.assertEqual(parse_shape(parse_view("Lampa LED CLASSIC B V 5.9W 830 Frosted E14")), "świeczka")
        self.assertEqual(parse_shape(parse_view("Lampa LED VALUE Classic P40 non-dim 4,9W 865 E14")), "kulka")
        self.assertEqual(parse_shape(parse_view("Lampa LED Retrofit CLASSIC GLOBE125 11W 827 Frosted E27")), "globe G125")
        self.assertEqual(parse_shape(parse_view("Lampa LED STAR CL STICK FR 60 non-dim 8W 827 E27")), "sztyft")
        self.assertEqual(parse_shape(parse_view("Lampa LED Retrofit CLASSIC A DIM 7.5W 827 Clear E27")), "A60")

    def test_bulb_line_glass_dim(self):
        res = build_name(row("Lampa LED Superior Class CLASSIC A Filament szkło przezroczyste 100 DIM 8.2W/827 E27 LEDVANCE"), RULES, {})
        self.assertEqual(res["nazwa"], "Żarówka LED E27 8,2W 2700K A60 przezroczysta ściemnialna Ledvance Superior")
        res = build_name(row("Lampa LED PERFORMANCE CLASS CLASSIC B 40 FR non-dim 4.9W/827 E14 LEDVANCE"), RULES, {})
        self.assertEqual(res["nazwa"], "Żarówka LED E14 4,9W 2700K świeczka matowa Ledvance Performance")
        u = parse_view("Lampa LED SMART WF A60 TW 6W/827 230VGLFR E274X1LEDV")
        self.assertEqual(parse_socket(u), "E27")

    def test_smart_protocol_and_sensor_bulbs(self):
        res = build_name(row("Lampa LED SMART+ WiFi CLASSIC A 60 FR 9W/RGBW E27"), RULES, {})
        self.assertEqual(res["nazwa"], "Żarówka LED E27 9W RGBW A60 matowa Ledvance Smart+ WiFi")
        res = build_name(row("Lampa LED SMART+ ZB CL A Frosted RGBW 60 yes 9W/ E27"), RULES, {})
        self.assertEqual(res["nazwa"], "Żarówka LED E27 9W RGBW A60 matowa Ledvance Smart+ Zigbee")
        res = build_name(row("Lampa LED DAYLIGHT SENSOR CLASSIC A 10W 840 Frosted E27"), RULES, {})
        # rodzaj czujnika z nazwy (przeglad 2026-10-21): DAYLIGHT = zmierzchu, MOTION = ruchu
        self.assertEqual(res["nazwa"], "Żarówka LED E27 10W 4000K A60 matowa z czujnikiem zmierzchu Ledvance")

    def test_decor_luminaires(self):
        # seria dekoru za marka, na koncu (2026-10-20: "Lampa stolowa E27 brazowa Ledvance Decor Cardboard Square")
        cases = {
            "Oprawa Decor Cardboard Cylinder Pendant E27 Brown": "Lampa wisząca E27 brązowa Ledvance Decor Cardboard Cylinder",
            "Oprawa ORBIS PARIS 400 mm 2XE27 WT": "Plafon fi400 2xE27 biały Ledvance Orbis Paris",
            "Oprawa Endura Classic IP65 E27 Orbick Wall Dark Gray": "Kinkiet zewnętrzny E27 IP65 grafitowy Ledvance Endura Classic Orbick",
        }
        for old, new in cases.items():
            self.assertEqual(build_name(row(old), RULES, {})["nazwa"], new)

    def test_decor_remote_and_reflector_accessory(self):
        # 2026-10-20: "WITH REMOTE CONTROL" -> "z pilotem" (bylo "Cob Strips With"); odblysnik to akcesorium, nie
        # "Reflektor szynowy ... Reflektor Tracklight Spot Reflector D95"
        res = build_name(row("Oprawa COB STRIPS CCT WITH REMOTE CONTROL 5M TW + RC"), RULES, {})
        self.assertEqual(res["nazwa"], "Taśma LED 5m TW z pilotem Ledvance Cob Strips")
        res = build_name(row("Reflektor do szynowych opraw TRACKLIGHT SPOT REFLECTOR D95 SP 4X1 LEDV"), RULES, {})
        self.assertEqual(res["nazwa"], "Odbłyśnik do Tracklight Spot fi95 Ledvance")  # "95 mm" + SP/FL w calym przebiegu
        res = build_name(row("Oprawa FLEX BENDABLE RGB + W 3M RGB 830 Remote Control"), RULES, {})
        self.assertEqual(res["nazwa"], "Taśma LED 3m 3000K RGBW z pilotem Ledvance Flex Bendable")

    def test_smart_luminaire_and_device(self):
        res = build_name(row("Oprawa SMART+ WIFI ORBIS WAVE 200 IP44 TW CR"), RULES, {})
        self.assertEqual(res["nazwa"], "Plafon Smart+ WiFi fi200 TW IP44 chrom Ledvance Orbis Wave")
        res = build_name(row("Czujnik SMART+ wykrywający dym z technologią WiFi"), RULES, {})
        self.assertEqual(res["nazwa"], "Czujnik dymu Smart+ WiFi Ledvance")
        res = build_name(row("Gniazdko SMART+ WIFI COMPACT OUTDOOR PLUG EU"), RULES, {})
        self.assertEqual(res["nazwa"], "Gniazdko Smart+ WiFi zewnętrzne Ledvance")

    def test_accessories_with_system(self):
        res = build_name(row("Oprawa TRUSYS FLEX P RAIL 4500 5P WT"), RULES, {})
        self.assertEqual(res["nazwa"], "Szyna do TruSys Flex 450cm 5P biała Ledvance")
        res = build_name(row("Ramka do oprawy DOWNLIGHT COMFORT FRAME D150 GY"), RULES, {})
        self.assertEqual(res["nazwa"], "Ramka do Downlight Comfort fi150 szara Ledvance")
        res = build_name(row("Oprawa DAMP PROOF HOUSING 1500 P 2XLAMP IP65 LEDV"), RULES, {})
        self.assertEqual(res["nazwa"], "Oprawa hermetyczna pod tuby LED 2x150cm IP65 Ledvance")

    def test_led_strip_accessories(self):
        res = build_name(row("Profil LS AY-PF01/UW/22x6/10/2"), RULES, {})
        self.assertEqual(res["nazwa"], "Profil do taśm LED 22x6mm 2m PF01/UW Ledvance")
        res = build_name(row("Przesłona LS AY-PC/W01/D/1"), RULES, {})
        self.assertEqual(res["nazwa"], "Przesłona do profilu LED mleczna 1m PC/W01 Ledvance")
        res = build_name(row("Złącza LS AY SUP-CSW/P2/50"), RULES, {})
        self.assertEqual(res["nazwa"], "Złączka do taśm LED SUP-CSW/P2/50 Ledvance")

    def test_discharge_lamps(self):
        res = build_name(row("Lampa wyładowcza HQI TS 2000W NDL/S K12S"), RULES, {})
        self.assertEqual(res["nazwa"], "Lampa metalohalogenkowa HQI-TS 2000W K12s 4000K Ledvance")
        res = build_name(row("Wysokoprężna lampa sodowa NAV-E 70/I"), RULES, {})
        self.assertEqual(res["nazwa"], "Lampa sodowa NAV-E 70W z zapłonnikiem Ledvance")  # "/I" = zaplonnik wewnetrzny

    def test_energy_class_line(self):
        res = build_name(row("Lampa LED CLASSIC B ENERGY EFFICIENCY C DIM 2.9W 827 Clear E14"), RULES, {})
        self.assertTrue(res["nazwa"].endswith("świeczka przezroczysta ściemnialna Ledvance klasa C"))

    def test_reflector_bulb_beam(self):
        res = build_name(row("Lampa LED VALUE PAR16 80 36 6,9W/840 230V GU10 BK"), RULES, {})
        self.assertEqual(res["nazwa"], "Żarówka LED GU10 6,9W 4000K PAR16 36° czarna Ledvance Value")  # BK = czarna
        res = build_name(row("Lampa LED PARATHOM non-dim R80 100 36° 9,1W 827 E27"), RULES, {})
        self.assertEqual(res["nazwa"], "Żarówka LED E27 9,1W 2700K R80 36° Ledvance Parathom")
        # zwykla zarowka nie dostaje kata, nawet gdy katalog podaje 300°
        res = build_name(row("Lampa LED PIN 28 320° P 2.2W 827 Clear G4"), RULES, {})
        self.assertNotIn("°", res["nazwa"])

    def test_optics(self):
        self.assertEqual(parse_optics(parse_view("Oprawa FLOODLIGHT H M LUM V 300W 730 SYM R30 LEDV")), "30°")
        self.assertEqual(parse_optics(parse_view("Oprawa FLOODLIGHT FLEX ASYMMETRIC 45 X 140 80W ML 840 A45X140 WAL")), "asym. 45x140°")
        self.assertEqual(parse_optics(parse_view("Oprawa HIGH BAY FLEX DALI 240W 110 840 WT")), "110°")
        self.assertEqual(parse_optics(parse_view("Oprawa TRUSYS FLEX P 50W 840 VW CL WT DALI")), "opt. VW")
        self.assertEqual(parse_optics(parse_view("Oprawa PANEL PERFORMANCE 600 UGR < 19 TPA 36 W 4000 K WT")), "")

    def test_driver(self):
        self.assertEqual(parse_driver("ZASILACZ OT 35/220-240/700"), "35W 700mA")
        self.assertEqual(parse_driver("ZASILACZ OTI DALI 160/220-240/24 2CHDT6/8"), "160W 24V")
        self.assertEqual(parse_driver("ZASILACZ OTI DALI 35/220-240/1A0 LT2"), "35W 1000mA")
        self.assertEqual(parse_driver("CONSTANT CURRENT LINEAR - DIMMABLE 90/400-750/220"), "90W 400-750mA")

    def test_split_for_ekonom_on_space(self):
        first, rest = split_for_ekonom("Oprawa hermetyczna LED Ledvance 150cm 42W 4000K IP69")
        self.assertLessEqual(len(first), 40)
        self.assertTrue(rest.startswith(" "))
        self.assertEqual(first + rest, "Oprawa hermetyczna LED Ledvance 150cm 42W 4000K IP69")


class PdfLumenTest(unittest.TestCase):
    def test_lumen_next_to_power_wins_over_equivalent_watt(self):
        # wiersz katalogu: zamiennik 40 W, moc 2.2, 470 lm, ... wymiar 85 (85/2.2 ~ 40 = "zamiennik")
        window = ["2", "40", "2.2", "470", "A", "Warm White", "50 000", "85", "45", "1", "6"]
        self.assertEqual(lumen_from_window(window, 2.2, {2700}), 470)

    def test_row_without_power_is_ignored(self):
        self.assertIsNone(lumen_from_window(["40", "470", "2700"], 5.9, {2700}))

    def test_thousands_with_space(self):
        self.assertEqual(lumen_from_window(["60", "9.5", "1 055", "E"], 9.5, {2700}), 1055)

    def test_texts_of_deleted_pdfs_come_from_cache(self):
        # katalogi PDF (setki MB) moga zniknac z input/ - lumeny dalej licza sie z tekstu w cache
        with tempfile.TemporaryDirectory() as tmp:
            pdf_dir, cache_dir = Path(tmp, "pdf"), Path(tmp, "cache")
            pdf_dir.mkdir()
            cache_dir.mkdir()
            (pdf_dir / "b_katalog.pdf").write_bytes(b"x" * 7)
            (cache_dir / "b_katalog.7.txt").write_text("B 9.5 1 055", encoding="utf-8")
            (cache_dir / "a_usuniety.2024.1.100.txt").write_text("stara wersja", encoding="utf-8")
            newer = cache_dir / "a_usuniety.2024.1.200.txt"
            newer.write_text(" 2.2 \n470", encoding="utf-8")
            os.utime(newer, (2_000_000_000, 2_000_000_000))

            texts = pdf_texts(pdf_dir, cache_dir)

        self.assertEqual(list(texts), ["a_usuniety.2024.1.pdf", "b_katalog.pdf"])
        self.assertEqual(texts["a_usuniety.2024.1.pdf"], ["2.2", "470"])
        self.assertEqual(texts["b_katalog.pdf"], ["B 9.5 1 055"])


COLUMNS = RULES["package_brand"]["catalog_columns"]
# polski katalog: jeden wiersz, EAN-y marek OSRAM | LEDVANCE | BELLALUX
PL_CATALOG = ["Nazwa produktu", "OSRAM ", "EAN10", "LEDVANCE ", "EAN10", "BELLALUX ", "EAN10", "Gwint",
              "VALUE CLA404,9W/827230VFRE2710X1EUEOSRAM", "4052899326927", "4099854073687", "4099854540653",
              "E27", "40", "470", "4,9", "2700"]
# katalog retail: ... trwalosc | dlugosc | srednica | sztuk w opak. | karton, potem strona i przypisy
RETAIL = ["LED STAR Spot PAR16 GL 50 827", "4099854459771", "2", "50", "3.7", "350", "E", "Warm White",
          "36 °", "15 000", "54", "50", "3", "6", "", "108", "1", "2", "3", "4",
          "LED STAR Spot PAR16 GL 50 840", "4099854459597", "2", "50", "3.7", "350", "E", "Cool White",
          "36 °", "15 000", "54", "50", "2", "6", "LED STAR Spot PAR16 GL 35 827"]


def built(name, ean="4099854000000", pkg=None):
    r = row(name, ean)
    res = build_name(r, RULES, {})
    res["row"], res["stara"] = r, full_name(r)
    return res


class PackageTest(unittest.TestCase):
    def setUp(self):
        self.pkg = package_from_pdfs({"zrodla_swiatla_led.pdf": PL_CATALOG, "retail_lamps_program.pdf": RETAIL},
                                     COLUMNS, RULES["package_catalogs"])

    def test_brand_columns_of_polish_catalog(self):
        self.assertEqual(self.pkg["4052899326927"]["brand"], "Osram")
        self.assertEqual(self.pkg["4099854073687"]["brand"], "Ledvance")
        self.assertEqual(self.pkg["4099854540653"]["brand"], "Bellalux")

    def test_pack_after_lifetime_ignores_footnotes(self):
        self.assertEqual(self.pkg["4099854459771"]["pack"], 3)
        self.assertEqual(self.pkg["4099854459597"]["pack"], 2)

    def test_pack_only_from_trusted_catalog(self):
        pkg = package_from_pdfs({"inny_katalog.pdf": RETAIL}, COLUMNS, RULES["package_catalogs"])
        self.assertIsNone(pkg["4099854459771"]["pack"])

    def test_brand_from_catalog_and_from_code(self):
        osram = built("Lampa LED VALUE CLASSIC A 40 4.9W/2700K E27", "4052899326927")
        coded = built("Lampa LED BASE Classic A60 8,5W/827 230V plastik E27 FS3 OSRAM")
        bella = built("Lampa LED BELLALUX CLA 40 4.9W/2700K E27", "4099854540653")
        apply_package_brand([osram, coded, bella], RULES, self.pkg)
        self.assertEqual(osram["nazwa"], "Żarówka LED E27 4,9W 2700K A60 Osram Value")
        self.assertEqual(coded["nazwa"], "Żarówka LED E27 8,5W 2700K A60 Osram Base")
        self.assertEqual(bella["nazwa"], "Żarówka LED E27 4,9W 2700K A60 Bellalux")

    def test_brand_removed_unless_only_difference(self):
        # 2026-10-19: nazwa bez producenta; marka zostaje tylko u blizniakow roznych marek
        osram = built("Lampa LED VALUE CLASSIC A 40 4.9W/2700K E27", "4052899326927")
        ledv = built("Lampa LED VALUE CLASSIC A 40 4.9W/2700K E27", "4099854999999")
        alone = built("Lampa LED CLASSIC B V 5.9W 830 Frosted E14", "4099854999998")
        bella = built("Lampa LED BELLALUX CLA 40 4.9W/2700K E27", "4099854540653")
        results = [osram, ledv, alone, bella]
        no_brand = {**RULES, "brand_in_name": False}  # tryb wylaczony z powrotem tego samego dnia
        apply_package_brand(results, no_brand, self.pkg)
        strip_brand_from_names(results, no_brand)
        self.assertEqual(osram["nazwa"], "Żarówka LED E27 4,9W 2700K A60 Osram Value")
        self.assertEqual(ledv["nazwa"], "Żarówka LED E27 4,9W 2700K A60 Ledvance Value")
        self.assertEqual(alone["nazwa"], "Żarówka LED E14 5,9W 3000K świeczka matowa")
        self.assertNotRegex(bella["nazwa"], r"(?i)ledvance|bellalux")

    def test_card_cri_separates_letter_twins(self):
        # 2026-10-20: "MR16 DIM S" / "DIM P" bez linii z litery maja te sama nazwe - rozni je CRI z karty GTIN
        s = built("Lampa LED MR16 DIM S 2.2W 940 GU5.3", "4099854455995")
        p = built("Lampa LED MR16 DIM P 2.2W 940 GU5.3", "4099854455996")
        other = built("Lampa LED MR16 DIM P 2.2W 930 GU5.3", "4099854455997")
        web = {"4099854455995": {"exact": True, "cri": 97, "title": "LED MR16 20 36° DIM S 2.2W 940 GU5.3"},
               "4099854455996": {"exact": True, "cri": 90, "title": "LED MR16 20 36° DIM P 2.2W 940 GU5.3"},
               "4099854455997": {"exact": True, "cri": 90, "title": "LED MR16 20 36° DIM P 2.2W 930 GU5.3"}}
        self.assertEqual(s["nazwa"], p["nazwa"])
        add_card_cri([s, p, other], RULES, web)
        move_brand_tail([s, p, other], RULES)
        self.assertEqual(s["nazwa"], "Żarówka LED GU5.3 12V 2,2W 4000K MR16 ściemnialna CRI 97 Ledvance")
        self.assertEqual(p["nazwa"], "Żarówka LED GU5.3 12V 2,2W 4000K MR16 ściemnialna CRI 90 Ledvance")
        self.assertNotIn("CRI", other["nazwa"])  # bez powtorzenia nic nie dopisujemy

    def test_lumen_from_gtin_card(self):
        # 2026-10-20: "brakuje lumenow a na stronie ledvance jest" - lm z karty GTIN, gdy moc i barwa sie zgadzaja
        card = {"exact": True, "lumen": 250.0, "power": 1.8, "cct": 2700, "url": "https://www.ledvance.com/en/gtin/4099854467622"}
        rules = {**RULES, "_gtin_cards": {"4099854467622": card}}
        res = build_name(row("Lampa LED CLASSIC GLOBE P 1.8W 827 Clear E27", "4099854467622"), rules, {})
        self.assertEqual(res["nazwa"], "Żarówka LED E27 1,8W 250lm 2700K globe przezroczysta Ledvance")
        self.assertEqual(res["status"], "OK")
        # kartoteka z inna moca / barwa niz karta - lm nie przenosimy
        res = build_name(row("Lampa LED CLASSIC GLOBE P 3.4W 827 Clear E27", "4099854467622"), rules, {})
        self.assertNotIn("lm", res["nazwa"])
        res = build_name(row("Lampa LED CLASSIC GLOBE P 1.8W 840 Clear E27", "4099854467622"), rules, {})
        self.assertNotIn("lm", res["nazwa"])
        # Power Select / Multi Lumen: lm stopnia o mocy z Ekonoma, "40W" = najwyzszy stopien 39,3 W; komplet SET6 - lm na sztuke
        cards = {"4099854307898": {"exact": True, "lumen": 6650.0, "power": 55.0, "cct": 3000,
                                   "powers": [55.0, 49.0, 43.0, 35.0], "lumens": [6650.0, 5900.0, 5250.0, 4350.0]},
                 "4099854000101": {"exact": True, "lumen": 6400.0, "power": 39.3, "cct": 4000,
                                   "powers": [39.3, 35.0, 31.1], "lumens": [6400.0, 5700.0, 5070.0]},
                 "4099854000102": {"exact": True, "lumen": 21870.0, "power": 198.0, "cct": None,
                                   "powers": [198.0], "lumens": [21870.0]}}
        rules = {**RULES, "_gtin_cards": cards}
        self.assertIn("55W 6650lm", build_name(row("Oprawa LINEAR INDIVILED DI 1200 P 55W 930 EM BK LEDV", "4099854307898"), rules, {})["nazwa"])
        self.assertIn("40W 6400lm", build_name(row("Oprawa DAMP PROOF HIGH EFFICIENCY 1500 P 40W ML 840 IP65", "4099854000101"), rules, {})["nazwa"])
        self.assertIn("33W 3645lm", build_name(row("Oprawa PANEL PRINT SKY 600 V 33W SET6", "4099854000102"), rules, {})["nazwa"])
        # kilka mocy w nazwie z Ekonoma - lm kazdej mocy w tej samej kolejnosci
        tube = {"exact": True, "lumen": 2300.0, "power": 14.0, "cct": 3000, "powers": [14.0, 11.3], "lumens": [2300.0, 1900.0]}
        rules = {**RULES, "_gtin_cards": {"4099854432743": tube}}
        self.assertIn("11,3/14W 1900/2300lm",
                      build_name(row("Tuba LED T8 EM S 1200 11.3W/14W 830 LEDV", "4099854432743"), rules, {})["nazwa"])

    def test_line_from_card_family(self):
        # 2026-10-20: "LED CLASSIC B V" = Value, bo karta tej rodziny pisze "VALUE CLAS B 60"; litera w Ekonomie nie decyduje
        web = {"4099854074158": {"exact": True, "family": "LED CLASSIC B V", "name": "VALUE CLAS B 60 6.8 W/2700 K E14"},
               "4099854484216": {"exact": True, "family": "LED CLASSIC B V", "name": "LED Classic B 60 Filament V 5.9W 830 Clear E14"},
               "4099854000201": {"exact": True, "family": "LED CLASSIC P V", "name": "LED Classic P 40 V 4.9W 827"},
               "4099854000202": {"exact": True, "family": "MIESZANA", "name": "VALUE X"},
               "4099854000203": {"exact": True, "family": "MIESZANA", "name": "PERFORMANCE Y"}}
        lines = card_family_lines(web)
        self.assertEqual(lines["4099854484216"], "Value")
        self.assertNotIn("4099854000201", lines)  # zadna karta rodziny nie pisze linii slowem
        self.assertNotIn("4099854000202", lines)  # karty rodziny sobie przecza
        # "Phase out" to status, nie rodzina - nie przenosi linii miedzy niepowiazanymi produktami
        web["4099854000204"] = {"exact": True, "family": "Phase out", "name": "LED STRIP VALUE-300"}
        web["4099854000205"] = {"exact": True, "family": "Phase out", "name": "LUMILUX T8"}
        self.assertNotIn("4099854000205", card_family_lines(web, RULES["family_from_card"]["placeholders"]))
        rules = {**RULES, "_family_lines": lines}
        res = build_name(row("Lampa LED CLASSIC B V 5.9W 830 Clear E14", "4099854484216"), rules, {})
        self.assertTrue(res["nazwa"].endswith("Ledvance Value"), res["nazwa"])

    def test_family_from_card(self):
        # 2026-10-20 "brakuje nazw wlasnych": nazwa wlasna oprawy z rodziny na karcie GTIN, bez slow technicznych
        cfg = RULES["family_from_card"]
        self.assertEqual(card_family_name("DAMP PROOF MAX EMERGENCY SENSOR", cfg), "Damp Proof Max")
        self.assertEqual(card_family_name("LINEAR INDIVILED DIRECT 1200 DALI SENSOR", cfg), "Linear IndiviLED Direct")
        self.assertEqual(card_family_name("TruSys PERFORMANCE EMERGENCY VERY WIDE", cfg), "TruSys Very Wide")
        self.assertEqual(card_family_name("FLOODLIGHT FLEX ASYMMETRIC 45 X 140", cfg), "Floodlight Flex")
        self.assertEqual(card_family_name("Orbis Livia", cfg), "Orbis Livia")
        self.assertEqual(card_family_name("HIGH BAY V DALI", cfg), "High Bay")  # V = litera klasy; "High Bay" wypada
        # dopiero jako powtorzenie rzeczownika "Oprawa High Bay LED"
        self.assertEqual(card_family_name("DAMP PROOF EMERGENCY GEN 2", cfg), "Damp Proof")
        # pojedyncze slowo tez jest nazwa wlasna ("LED BULKHEAD" -> "Bulkhead", 2026-10-20)
        self.assertEqual(card_family_name("SMART+ MATTER Floodlights", cfg), "Floodlights")
        web = {"4058075000004": {"exact": True, "family": "LED BULKHEAD", "name": "LED Bulkhead 11W 4000K Black"}}
        rules = {**RULES, "_card_families": card_families(web, cfg)}
        res = build_name(row("Oprawa LED BULKHEAD 11W 4000K Black", "4058075000004"), rules, {})
        self.assertEqual(res["nazwa"], "Plafon LED 11W 4000K czarny Ledvance Bulkhead")
        self.assertEqual(card_family_name("PANEL PERFORMANCE 600 UGR<19", cfg), "Panel")  # wypada jako powtorzenie "Panel LED"
        self.assertEqual(card_family_name("Phase out", cfg), "")
        web = {"4099854000301": {"exact": True, "family": "DAMP PROOF MAX DALI", "name": "DP MAX DA 1200 S 43W 840 VW IP69"}}
        rules = {**RULES, "_card_families": card_families(web, cfg)}
        res = build_name(row("Oprawa DAMP PROOF MAX DA 1200 S 43W 840 VW IP69 LEDV", "4099854000301"), rules, {})
        self.assertTrue(res["nazwa"].endswith("DALI Ledvance Damp Proof Max"), res["nazwa"])
        # bez karty: nazwa wlasna innych kart, gdy stoi doslownie w nazwie z Ekonoma
        web = {"4099854000303": {"exact": True, "family": "DAMP PROOF HE DALI", "name": "DP HE DA 1500 P 49W 840"},
               "4099854000305": {"exact": False, "family": "PANEL PRINT SKY 625 DALI", "title": "PL PRINT SKY DA 625"}}
        self.assertEqual(family_from_name("OPRAWA PANEL PRINT SKY DA 625 V 33W SET9", family_phrases(web, cfg)),
                         "Panel Print Sky")
        rules = {**RULES, "_card_families": card_families(web, cfg), "_family_phrases": family_phrases(web, cfg)}
        res = build_name(row("Oprawa DAMP PROOF HE DA 1500 P 49W 865 IP65 LEDV", "4099854000304"), rules, {})
        self.assertTrue(res["nazwa"].endswith("Ledvance Damp Proof High Efficiency"), res["nazwa"])
        # "PLANON™" w Ekonomie, karta pod tym EAN zwraca blad - nazwa z innych kart; ROUND -> "okragly"
        web = {"4058075470736": {"exact": True, "family": "PLANON™ Frameless", "name": "Planon Frameless Round 450mm"}}
        rules = {**RULES, "_card_families": card_families(web, cfg), "_family_phrases": family_phrases(web, cfg)}
        res = build_name(row("Oprawa PLANON™ Frameless Round 300mm 19W 3000K", "4058075470712"), rules, {})
        self.assertEqual(res["nazwa"], "Panel LED okrągły fi30cm 19W 3000K Ledvance Planon Frameless")
        # czujnik z rodziny karty, gdy Ekonom pisze tylko litere "S"; oprawa pod tuby: nazwa wlasna za "pod tube LED"
        web = {"4099854102516": {"exact": True, "family": "SURFACE SLIM IK10 RD 200 SENSOR", "name": "SF SLIM RD S 200"},
               "4099854000306": {"exact": True, "family": "DAMP PROOF HOUSING GEN 3", "name": "DP HOUSING 1200"}}
        rules = {**RULES, "_card_families": card_families(web, cfg), "_family_phrases": family_phrases(web, cfg),
                 "_card_family_raw": {e: c["family"] for e, c in web.items()}}
        res = build_name(row("Oprawa ścienna SURFACE SLIM RD S 200 P 10W 840 IP65 EUELEDV", "4099854102516"), rules, {})
        self.assertIn("z czujnikiem", res["nazwa"])
        res = build_name(row("Oprawa DAMP PROOF HOUSING 1200 P 1XLAMP IP65", "4099854000306"), rules, {})
        self.assertEqual(res["nazwa"], "Oprawa hermetyczna pod tubę LED 1x120cm IP65 Ledvance Damp Proof Housing")
        # Orbis: Click-DIM / Click-CCT z nazwy z Ekonoma ("Orbis CCT-DIM" na karcie gubil Click), 2026-10-20
        web = {"4058075651753": {"exact": True, "family": "Orbis CCT-DIM", "name": "Orbis Click-CCT 410mm 24W"}}
        rules = {**RULES, "_card_families": card_families(web, cfg), "_family_phrases": family_phrases(web, cfg)}
        res = build_name(row("Oprawa ORBIS CCT / DIM Click-CCT 410mm 24W", "4058075651753"), rules, {})
        self.assertIn("24W Click-CCT", res["nazwa"])
        self.assertNotIn("CCT Click-CCT", res["nazwa"])
        res = build_name(row("Oprawa ORBIS CCT / DIM Remote-CCT 310mm 14W"), RULES, {})
        self.assertIn("14W Remote-CCT", res["nazwa"])
        res = build_name(row("Oprawa DAMP PROOF HE 5X THROUGHWIRING 1500 P 40W ML 840 IP65"), RULES, {})
        self.assertTrue(res["nazwa"].endswith("Ledvance High Efficiency 5x Throughwiring"), res["nazwa"])
        # hermetyki 2026-10-20: Eco Class z karty, Throughwiring, bez "Narrow" przy kacie, "Phase out" -> nazwa karty
        web = {"4058075000001": {"exact": True, "family": "ECO CLASS DAMP PROOF HLO GEN 2", "name": "ECO CLASS DAMP PROOF HLO 1500 66 W 4000 K IP65 TH"},
               "4058075000002": {"exact": True, "family": "Phase out", "name": "DAMP PROOF ECO VERY WIDE"}}
        rules = {**RULES, "_card_families": card_families(web, cfg), "_gtin_cards": web,
                 "_card_family_raw": {e: c["family"] for e, c in web.items()}}
        res = build_name(row("Oprawa DAMP PROOF LED ECO HIGH LUMEN OUTPUT 66W 840 1500 IP65 GY", "4058075000001"), rules, {})
        self.assertTrue(res["nazwa"].endswith("Ledvance Eco Class Damp Proof High Lumen Output Throughwiring"), res["nazwa"])
        res = build_name(row("Oprawa DAMP PROOF 1500 E 66W 840 VW IP65 GY EUE", "4058075000002"), rules, {})
        self.assertTrue(res["nazwa"].endswith("Ledvance Damp Proof Eco"), res["nazwa"])
        # plafony 2026-10-20: srednica z samej liczby, kwadrat "Square 200mm", kolory dwutonowe, "Color + White"
        res = build_name(row("Oprawa ścienna SURFACE BULKHEAD 250 10W 3000K WT IP65 EMERGENCY"), RULES, {})
        self.assertIn("fi250 10W", res["nazwa"])
        res = build_name(row("Oprawa LED CLICK White Square 200mm 16W"), RULES, {})
        self.assertIn("200x200 mm 16W", res["nazwa"])
        res = build_name(row("Oprawa Orbis Plate 310MM 19W 830 Click DIM Black and Blue"), RULES, {})
        self.assertIn("Click-DIM czarno-niebieski", res["nazwa"])
        res = build_name(row("Oprawa Orbis Livia 300MM 16W 830 Click DIM White and Wood Decor"), RULES, {})
        self.assertIn("biały/drewno", res["nazwa"])
        res = build_name(row("Oprawa LED COLOR + WHITE Round 200mm 18W + RC"), RULES, {})
        self.assertNotIn("biały", res["nazwa"])
        self.assertEqual(card_family_name("LED Color and White", cfg), "Color + White")
        # "Round" / "RD" / "Circular" w nazwie wlasnej - bez osobnego "okragly" (2026-10-20)
        web = {"4058075000003": {"exact": True, "family": "CEILING ROUND IP44", "name": "CEILING ROUND 210 12W 830 IP44"}}
        rules = {**RULES, "_card_families": card_families(web, cfg)}
        res = build_name(row("Oprawa sufitowa CEILING ROUND IP44 210 12W 830 IP44", "4058075000003"), rules, {})
        self.assertEqual(res["nazwa"], "Plafon LED fi210 12W 3000K IP44 Ledvance Ceiling Round")
        res = build_name(row("Oprawa Orbis Yummy 285MM 16W 830 Click DIM Black"), RULES, {})
        self.assertIn("3000K Click-DIM", res["nazwa"])
        # slowo z poczatku nazwy nie wraca w nazwie wlasnej: "Panel LED Flex ...", nie "Panel LED Panel Flex"
        web = {"4099854000302": {"exact": True, "family": "PANEL FLEX 600 UGR<19 PS", "name": "PL FLEX 600 P 30W 840 U19"}}
        rules = {**RULES, "_card_families": card_families(web, cfg)}
        res = build_name(row("Oprawa PL FLEX 600 P 30W 840 U19 PS", "4099854000302"), rules, {})
        self.assertTrue(res["nazwa"].startswith("Panel LED"), res["nazwa"])
        self.assertEqual(res["nazwa"], "Panel LED 60x60 30W 4000K UGR19 Ledvance Flex")
        self.assertNotIn("Panel Flex", res["nazwa"])

    def test_diameter_in_mm(self):
        # 2026-10-19: "zamiast fi najlepiej dac 220 mm"
        self.assertEqual(diameter_in_mm("Downlight LED fi220 30W 3000K biały Ledvance"),
                         "Downlight LED 220 mm 30W 3000K biały Ledvance")
        self.assertEqual(diameter_in_mm("Panel LED fi30cm 19W 3000K Ledvance"), "Panel LED 300 mm 19W 3000K Ledvance")
        self.assertEqual(diameter_in_mm("Złączka fix 3P"), "Złączka fix 3P")

    def test_brand_tail_after_suffixes(self):
        # 2026-10-19: marka + linia na koncu, takze za wersja / wyroznikiem dopisanymi pozniej, przed "op. N szt."
        res = {"nazwa": "Zasilacz LED 150W 24V Ledvance Performance DR-PFM-150 op. 3 szt.", "po_marce": ["Performance"]}
        bulb = {"nazwa": "Żarówka LED E27 9W TW A60 matowa Osram Smart+ Zigbee Globe 95", "po_marce": ["Smart+", "Zigbee"]}
        move_brand_tail([res, bulb], RULES)
        self.assertEqual(res["nazwa"], "Zasilacz LED 150W 24V DR-PFM-150 Ledvance Performance op. 3 szt.")
        self.assertEqual(bulb["nazwa"], "Żarówka LED E27 9W TW A60 matowa Globe 95 Osram Smart+ Zigbee")

    def test_single_brand_in_name(self):
        # 2026-10-19 "jednak chcemy nazwe producenta": marka zostaje, ale jedna ("Ledvance Bellalux" -> "Bellalux")
        alone = built("Lampa LED CLASSIC B V 5.9W 830 Frosted E14", "4099854999998")
        bella = built("Lampa LED BELLALUX CLA 40 4.9W/2700K E27", "4099854999997")
        downlight = built("Oprawa SPOT ADJ P 8W 927 PS DIM IP20 WT", "4099854999996")
        results = [alone, bella, downlight]
        apply_package_brand(results, RULES, self.pkg)
        strip_brand_from_names(results, RULES)
        self.assertEqual(alone["nazwa"], "Żarówka LED E14 5,9W 3000K świeczka matowa Ledvance")
        self.assertEqual(bella["nazwa"], "Żarówka LED E27 4,9W 2700K A60 Bellalux")
        self.assertTrue(downlight["nazwa"].startswith("Downlight LED ") and "Ledvance" in downlight["nazwa"])

    def test_pack_suffix(self):
        fs = built("Lampa LED BASE Classic A60 8,5W/827 230V plastik E27 FS3 OSRAM")
        code = built("Lampa LED SMARTWIFI A60 9W/827 230VDIMFRE274X3LEDV")
        beam = built("Oprawa FLOODLIGHT H M LUM V 300W 740 ASYM 110X30 LEDV")
        cat = built("Lampa LED STAR PAR16 3.7W 827 GU10", "4099854459771")
        apply_pack_size([fs, code, beam, cat], RULES, self.pkg)
        self.assertTrue(fs["nazwa"].endswith("2700K A60 Ledvance Base op. 3 szt."))
        self.assertTrue(code["nazwa"].endswith(" op. 3 szt."))
        self.assertNotIn("szt.", beam["nazwa"])  # 110X30 to kat swiecenia
        self.assertTrue(cat["nazwa"].endswith(" op. 3 szt."))

    def test_pack_word_not_doubled(self):
        res = built("Złączka LF-2TERM 10 pack")
        apply_pack_size([res], RULES, self.pkg)
        self.assertEqual(res["nazwa"].count("10"), 1)
        self.assertTrue(res["nazwa"].endswith("op. 10 szt."))

    def test_swapped_cct_fixed_from_catalog(self):
        res = built("Lampa LED STAR PAR16 3.7W 840 GU10", "4099854459771")
        check_catalog_cct([res], self.pkg)
        self.assertIn("2700K", res["nazwa"])
        self.assertNotIn("4000K", res["nazwa"])

    def test_catalog_row_of_other_product_is_flagged(self):
        res = built("Lampa LED STAR PAR16 5.6W 840 GU10", "4099854459771")
        check_catalog_cct([res], self.pkg)
        self.assertIn("4000K", res["nazwa"])  # nazwa bez zmian
        self.assertEqual(res["status"], "DO_SPRAWDZENIA")


class DistinguishTest(unittest.TestCase):
    @staticmethod
    def group(names, src="lamps_launch_trade_2024.pdf", old="Lampa VINTAGE 1906 LED CLASSIC GLOBE 4W 818 Gold E27"):
        """Produkty o tej samej nazwie w Ekonomie i ich wiersze katalogu."""
        results, pkg = [], {}
        for i, cat in enumerate(names):
            ean = f"40998540000{i:02d}"
            results.append(built(old, ean))
            pkg[ean] = {"name": cat, "src": src, "numbers": [], "brand": "", "pack": None}
        return results, pkg

    @staticmethod
    def tails(results, base):
        return [body(r["nazwa"])[len(body(base)):].strip() for r in results]

    def test_only_as_much_as_needed(self):
        results, pkg = self.group(["Vintage 1906 Globe 80 40 Spiral Filament DIM 4W 818 Gold E271",
                                   "Vintage 1906 Globe 80 40 Heart Filament DIM 4W 818 Gold E271",
                                   "Vintage 1906 Globe 95 40 Heart Filament DIM 4W 818 Gold E271"])
        base = results[0]["nazwa"]
        add_all_distinguishers(results, RULES, pkg, {})
        # Spiral jest jeden; Heart sa dwa - drugi wyroznik to srednica
        self.assertEqual(self.tails(results, base), ["Spiral", "Globe 80 Heart", "Globe 95"])

    def test_minimal_words_and_colors(self):
        results, pkg = self.group(["DECOR LED FLAT ELECTRO GLASS SMALL ROSE 3.5W 827 E27",
                                   "DECOR LED FLAT ELECTRO GLASS LARGE ROSE 3.5W 827 E27",
                                   "DECOR LED FLAT ELECTRO GLASS LARGE GOLD 3.5W 827 E27"],
                                  src="vintage_edition_1906.pdf")
        base = results[0]["nazwa"]
        add_all_distinguishers(results, RULES, pkg, {})
        # "Gold" z nazwy w Ekonomie jest juz w nazwie jako kolor szkla (2026-10-21), wiec trzeci produkt odroznia sie
        # brakiem dopisku - "duza" bez "rozowa"
        self.assertIn("Gold", base)
        self.assertEqual(self.tails(results, base), ["mała", "duża różowa", ""])

    def test_class_letter_not_added_twice(self):
        # 2026-10-19: "CLASSIC B V 60" - litera V = Value; kolejna runda wyroznikow dopisywala "Value Value"
        results = [built("Lampa LED CLASSIC B V 60 7.5 W/2700 K E14", "4099854074158"),
                   built("Lampa LED VALUE CLASSIC B 60 7.5 W/2700 K E14", "4058075479159"),
                   built("Lampa LED CLASSIC B 60 7.5 W/2700 K E14", "4099854074159")]
        add_all_distinguishers(results, RULES, {}, {})
        for res in results:
            self.assertNotRegex(res["nazwa"], r"(\w+) \1")

    def test_beam_angle(self):
        results, pkg = self.group(["LED PAR16 80 36° P 6.9W 827 GU101", "LED PAR16 80 60° P 6.9W 827 GU101"])
        add_all_distinguishers(results, RULES, pkg, {})
        self.assertEqual([body(r["nazwa"])[-3:] for r in results], ["36°", "60°"])

    def test_mixed_catalog_formats_left_alone(self):
        # ten sam produkt opisany stylem dwoch katalogow - to nie jest roznica
        results, pkg = self.group(["LED STAR CLASSIC A FR 40 Recycled Plastic 827",
                                   "LED Classic A 40 Recycled Plastic 4.9W 827 Frosted E271"])
        pkg[results[0]["row"]["Kod producenta"]]["src"] = "retail_lamps_program_2025.pdf"
        before = [r["nazwa"] for r in results]
        add_all_distinguishers(results, RULES, pkg, {})
        self.assertEqual([r["nazwa"] for r in results], before)

    def test_luminaire_catalog_not_used(self):
        results, pkg = self.group(["SF BLKH 300 EM 15 W 4000 K WT IP65", "SF BLKH 300 EM 15 W 3000 K BK IP65"],
                                  src="led_luminaires_assortment_professional.pdf")
        before = [r["nazwa"] for r in results]
        add_all_distinguishers(results, RULES, pkg, {})
        self.assertEqual([r["nazwa"] for r in results], before)

    def test_old_name_color_code(self):
        olds = [f"Constant Current Track - Performance high efficacy 25/450-600/42 {c}" for c in ("BK", "GY", "WT")]
        results = [built(o, f"409985400000{i}") for i, o in enumerate(olds)]
        self.assertEqual(len({r["nazwa"] for r in results}), 1)
        add_all_distinguishers(results, RULES, {}, {})
        self.assertEqual([body(r["nazwa"]).split()[-1] for r in results], ["czarny", "szary", "biały"])

    def test_pack_from_gtin_page(self):
        results = [built("Lampa LED VALUE CLASSIC A 60 8 W/3000 K E27", f"409985400000{i}") for i in range(2)]
        web = {"4099854000000": {"title": "LED VALUE CLAS A 60 8 W/3000 K E27", "exact": True, "pieces": 2},
               "4099854000001": {"title": "LED VALUE CLAS A 60 8 W/3000 K E27", "exact": True, "pieces": 3}}
        apply_pack_size(results, RULES, {}, web)
        self.assertEqual([r["nazwa"][-10:] for r in results], ["op. 2 szt.", "op. 3 szt."])

    def test_length_socket_and_twins(self):
        flex = [built(f"Oprawa FLEX RGB WITH REMOTE CONTROL {n}M RGB + RC", f"409985400000{i}") for i, n in enumerate((3, 5))]
        add_all_distinguishers(flex, RULES, {}, {})
        # taśma z pilotem, nie akcesorium "Pilot RC 3m" (2026-10-20); dlugosc juz w nazwie
        self.assertEqual([r["nazwa"] for r in flex], ["Taśma LED 3m RGB z pilotem Ledvance Flex",
                                                      "Taśma LED 5m RGB z pilotem Ledvance Flex"])
        olds = (["Żarówka halogenowa HALOLINE SUPERSTAR 120 W 230 V R7S"] * 2
                + ["Żarówka halogenowa HALOLINE SUPERSTAR 120 W 240 V R7S"] * 2)
        halo = [built(o, f"409985400001{i}") for i, o in enumerate(olds)]
        add_all_distinguishers(halo, RULES, {}, {})
        self.assertEqual([body(r["nazwa"]).split()[-1] for r in halo], ["230V", "230V", "240V", "240V"])

    def test_web_name_last(self):
        results = [built("Lampa LED VALUE CLASSIC A 60 8 W/3000 K E27", f"409985400000{i}") for i in range(2)]
        web = {"4099854000000": {"title": "LED Value Classic A 60 8W 830 Frosted E27", "exact": True, "pieces": 1},
               "4099854000001": {"title": "LED Value Classic A 60 8W 830 Clear E27", "exact": True, "pieces": 1}}
        add_all_distinguishers(results, RULES, {}, web)
        self.assertEqual([body(r["nazwa"]).split()[-1] for r in results], ["matowa", "przezroczysta"])
        self.assertIn("ledvance.com", results[0]["uwagi"][-1])


class BuildNameTest(unittest.TestCase):
    def test_floodlight(self):
        res = build_name(row("Naświetlacz ECO HP FLOOD 300W 857 W 40500LM black"), RULES, {})
        # "857 W": W to optyka (karta 4058075372498: 90°), nie moc - wczesniej wypadala (2026-10-20)
        self.assertEqual(res["nazwa"], "Naświetlacz LED 300W 40500lm 5700K opt. W czarny Ledvance")
        self.assertEqual(res["status"], "OK")
        rules ={**RULES, "_gtin_cards": {"4058075372498": {"exact": True, "beam": "90°"}}}
        res = build_name(row("Naświetlacz ECO HP FLOOD 300W 857 W 40500LM black", "4058075372498"), rules, {})
        self.assertEqual(res["nazwa"], "Naświetlacz LED 300W 40500lm 5700K 90° czarny Ledvance")
        res = build_name(row("Oprawa LOW BAY FLEX 1500 P 140W 840 OV", "4099854163715"),
                         {**RULES, "_gtin_cards": {"4099854163715": {"exact": True, "beam": "40x90°"}}}, {})
        self.assertIn("4000K 40x90°", res["nazwa"])

    def test_missing_lumen_is_flagged(self):
        res = build_name(row("Lampa LED PAR16 DIM P 8W 830 GU10"), RULES, {})
        self.assertEqual(res["nazwa"], "Żarówka LED GU10 8W 3000K PAR16 ściemnialna Ledvance")
        self.assertEqual(res["status"], "DO_SPRAWDZENIA")
        self.assertIn("lumen", res["braki"])

    def test_pdf_lumen_used(self):
        pdf = {"4099854000001": {"lm": [806.0], "src": ["katalog.pdf"]}}
        res = build_name(row("Lampa LED Retrofit CLASSIC P 5.9W 865 Frosted E14", "4099854000001"), RULES, pdf)
        self.assertEqual(res["nazwa"], "Żarówka LED E14 5,9W 806lm 6500K kulka matowa Ledvance")
        self.assertEqual(res["status"], "OK")

    def test_luminaire_types_from_generic_oprawa(self):
        # 2026-10-19: "Oprawa LED nic nie mowi" - konkretne typy zamiast ogolnej oprawy
        cases = {
            "Oprawa SLCOMPM6V055W 740RV40STNWAC2BTONEUELEDV": "Oprawa uliczna LED",
            "Oprawa SF DISC 300 V 21W ML 83040 LEDV": "Plafon LED",
            "Oprawa FLOOR STANDING COMPACT SN EU V 88W 840 WT": "Lampa stojąca LED",
            "Oprawa CANOPY COMP V 60W 840 IP66 DIM LEDV": "Oprawa do zadaszeń LED",
            "Oprawa WALLPACK COMBO V 24W ML 84065 IP65": "Oprawa ścienna LED",
            "Oprawa TubeKIT™ 600mm 8.9W 3000K": "Oprawa liniowa LED 60cm",
            "Oprawa LED TUBE T8 EM ULTRA OUTPUT 1200 mm 20W 830": "Tuba LED T8 120cm",
            "Oprawa Highbay Ufo 80W 865 IP65 Black": "Oprawa High Bay LED",
            "Oprawa EM EXIT E 0,8W 30M 3H AT FS1 LEDV": "Oprawa awaryjna LED",
            "Oprawa MAGNETIC LED RING 12W ML 83040 AC": "Moduł LED",
            "Oprawa SPOT ADJ P 8W 927 PS DIM IP20 WT": "Downlight LED",
        }
        for name, start in cases.items():
            self.assertTrue(build_name(row(name), RULES, {})["nazwa"].replace(" Ledvance", "").startswith(start), name)
        # akcesoria EM Exit bez mocy zostaja akcesoriami
        self.assertEqual(build_name(row("Oprawa EM EXIT PICTO UP LEDV"), RULES, {})["typ"], "akcesorium")

    def test_line_in_every_type(self):
        # 2026-10-19 "bedziemy chcieli dodawac linie jesli jest" - ale tylko zapisana slowem; sama litera V/P/S
        # wylaczona tego samego dnia ("moga byc bledne": "MR16 DIM S" bez klasy na karcie, "URBAN LITE S" = rozmiar)
        self.assertEqual(build_name(row("Oprawa SF DISC 300 V 21W ML 83040 LEDV"), RULES, {})["nazwa"],
                         "Plafon LED fi300 21W 3000/4000K Ledvance")  # 300 = srednica plafonu (2026-10-20)
        self.assertNotIn("Performance", build_name(row("Oprawa SPOT ADJ P 8W 927 PS DIM IP20 WT"), RULES, {})["nazwa"])
        self.assertNotIn("Superior", build_name(row("Lampa LED PAR16 DIM S 3.7W 927 GU10"), RULES, {})["nazwa"])
        self.assertIn("Ledvance Value", build_name(row("Oprawa PANEL VALUE 600 36W 4000K"), RULES, {})["nazwa"])
        # "CLASSIC P 40" to kulka, nie Performance
        self.assertNotIn("Performance", build_name(row("Lampa LED CLASSIC P 40 4.9W 827 Frosted E14"), RULES, {})["nazwa"])
        # linia z serii / "Smart+" rzeczownika nie jest dublowana
        self.assertEqual(build_name(row("Lampa LED SMART+ WiFi CLASSIC A 60 FR 9W/RGBW E27"), RULES, {})["nazwa"].count("Smart+"), 1)

    def test_deco_size_is_not_cct_code(self):
        res = build_name(row("Oprawa Decor Felt Cylinder Pendant 28W 830 Dark Gray"), RULES, {})
        self.assertNotIn("fi830", res["nazwa"])

    def test_tube_has_length_and_ballast(self):
        pdf = {"4099854000002": {"lm": [1800.0], "src": ["katalog.pdf"]}}
        res = build_name(row("LEDTUBE T8 EM S 1050 11.4W 830", "4099854000002"), RULES, pdf)
        # sama litera klasy ("EM S 1050") nie daje linii (2026-10-19)
        self.assertEqual(res["nazwa"], "Tuba LED T8 105cm 11,4W 1800lm 3000K EM Ledvance")

    def test_tube_and_linear_lines(self):
        # 2026-10-19: "niektore linie dzialaja i sa przypisane, a niektore nie" - Ultra Output (UO) laczy sie z klasa
        # zapisana slowem, jedna pisownia (Entry, nie ENTRY); sama litera V/P/S nie jest linia
        cases = {
            "Tuba LED T8 EM ULTRA OUTPUT S 1500 mm 22.1W 965": "Ultra Output",
            "LEDTUBE T8 EM UO P 1500 23.1W 865": "Ultra Output",
            "Tuba LED T8 EM SUPERIOR 900 mm 9.7W 840": "Superior",
            "Tuba LED T8 ENTRY AC 16 W/3000 K 1200 mm": "Entry",
            "Tuba LED T8 ADV Ultra Output 1200mm 15,6W 830 EM": "Advanced Ultra Output",
            "Tuba LED T8 PRO 1200mm 12,7W 840 EM": "Pro",
            "Oprawa LINEAR ULTRA OUTPUT GEN 2 P 1500 48W CPS WT": "Ultra Output",
            "Oprawa URBAN LITE S 15W 730 1650LM IP65 GRAY": "Urban Lite S",
            "Oprawa URBAN VILLA RETROFIT ASYMMETRIC 40W 722 A60X120": "Urban Villa Retrofit",
            "Oprawa URBAN WALL VARIO SQ P 29W ML 82227 GY": "Urban Wall Vario",
        }
        for old, line in cases.items():
            name = build_name(row(old), RULES, {})["nazwa"]
            # rodzina oprawy (Urban) tez na koncu za marka (2026-10-20: "nazwa wlasna jednak na koncu wszedzie")
            self.assertTrue(name.endswith(f"Ledvance {line}"), f"{old} -> {name}")
            self.assertNotRegex(name, r"\b(UO|ENTRY|ULTRA OUTPUT|ADV)\b")

    def test_more_lines(self):
        cases = {
            "Oprawa LED SPOT RECESS TWIST PRO 70mm 5W 4000K White White": "Downlight LED 70 mm 5W 4000K biały Ledvance Pro",
            "Oprawa DOWNLIGHT PFM DN195 21W 830 WT IP54 UGR19": "Performance",
            "Lampa LED PARATHOM DIM PRO Classic A60 Filament szkło przezroczyste 7,5W/927 E27": "Parathom Pro",
        }
        for old, expected in cases.items():
            name = diameter_in_mm(build_name(row(old), RULES, {})["nazwa"])
            self.assertTrue(name == expected or name.endswith(f"Ledvance {expected}"), f"{old} -> {name}")

    def test_square_size(self):
        # 2026-10-19: liczba za SQUARE / SQ to wymiar ("trzeba dopisac 60x60 mm na przyklad"); bez liczby nic
        cases = {
            "Oprawa SURFACE SQUARE 330 24W 830 IP44": "330x330 mm",
            "SURFACE FLAT LIGHT SOURCES 500 SQ 27W 830 840 SF FLAT": "500x500 mm",
            "SURFACE FLAT LIGHT SOURCES 500 R 27W 830 840 SF FLAT": "500 mm",
            "Oprawa DOWNLIGHT SLIM SQUARE 155 12W 3000K WT IP20": "155x155 mm",
            "Oprawa Orbis Alicia 292X292MM 16W 830 Click DIM Wood Decor": "292x292 mm",
        }
        for old, size in cases.items():
            name = diameter_in_mm(build_name(row(old), RULES, {})["nazwa"])
            self.assertIn(f" {size} ", name, old)
            self.assertNotRegex(name, r"\bSQ\b")
        self.assertNotIn(" mm", build_name(row("Świetlówka pinowa SQUARE 16W 827"), RULES, {})["nazwa"])
        self.assertNotIn(" mm", build_name(row("Oprawa Decor Cardboard Square Ceiling E27 Brown"), RULES, {})["nazwa"])

    def test_accessory_without_power_is_not_a_luminaire(self):
        res = build_name(row("Oprawa HB BALL PROOF COVER 310 LEDV"), RULES, {})
        self.assertEqual(res["typ"], "akcesorium")
        self.assertEqual(res["nazwa"], "Osłona do High Bay fi310 Ledvance")

    def test_luminaire_color_optics_ugr(self):
        res = build_name(row("Oprawa DOWNLIGHT SURFACE DN 160 P 30W 830 36D IP65 BK"), RULES, {})
        self.assertEqual(res["nazwa"], "Downlight LED fi160 30W 3000K 36° IP65 czarny Ledvance")
        res = build_name(row("Oprawa PANEL PERFORMANCE 600 36W 4000K UGR19"), RULES, {})
        self.assertEqual(res["nazwa"], "Panel LED 60x60 36W 4000K UGR19 Ledvance Performance")
        res = build_name(row("Oprawa AR SPOT1 RD V 6W 830 SYM24 BK ONEUELEDV"), RULES, {})
        # 830 to barwa, nie srednica; RD = okragly; AR SPOT = reflektor (noun_overrides, 2026-10-21)
        self.assertEqual(res["nazwa"], "Reflektor LED okrągły 6W 3000K 24° czarny Ledvance")

    def test_version_suffix_when_name_still_repeats(self):
        results = []
        for name in ["Oprawa SP FIX P 8W 927 PS DIM IP44 WT", "Oprawa SP DARKLIGHT FIX P 8W 927 PS DIM IP44 WT"]:
            res = build_name(row(name), RULES, {})
            res["stara"] = name
            results.append(res)
        add_version_suffixes(results, RULES)
        self.assertEqual(results[0]["nazwa"], "Downlight LED 8W 2700K IP44 biały Ledvance")
        self.assertEqual(results[1]["nazwa"], "Downlight LED 8W 2700K IP44 biały Ledvance DARKLIGHT")

    def test_dulux_socket_from_table(self):
        res = build_name(row("Świetlówka pinowa DULUX D 18W 827"), RULES, {})
        self.assertEqual(res["nazwa"], "Świetlówka kompaktowa G24d-2 18W 2700K Ledvance")


KOBI = load_yaml(DEFAULT_RULES)["suppliers"]["kobi"]


def kobi_row(name, code="004666/KOB"):
    return {"Nazwa towaru": name[:40], "Nazwa dodatkowa": name[40:], "Kod towaru": code, "Kod producenta": ""}


class KobiTest(unittest.TestCase):
    def test_bulb_shape_code_and_feed_lumen(self):
        feed = {"4666": {"_name": "Żarówka LED GS 10W E27 4000K Kobi", "Marka": "Kobi",
                         "Strumień świetlny |lm|": "950", "Kolor produktu": "biały"}}
        res = build_name(kobi_row("LED GS 10W E27 4000K"), KOBI, {}, feed)
        self.assertEqual(res["nazwa"], "Żarówka LED E27 10W 950lm 4000K A60 Kobi")
        self.assertEqual(res["status"], "OK")

    def test_brand_line_from_name_without_feed(self):
        res = build_name(kobi_row("LED MB 9W E27 4000K PREMIUM"), KOBI, {}, {})
        self.assertEqual(res["nazwa"], "Żarówka LED E27 9W 4000K kulka Kobi Premium")
        self.assertIn("lumen", res["braki"])

    def test_old_caps_name_uses_feed_name_and_sensor(self):
        feed = {"4666": {"_name": "Plafon z czujnikiem ruchu LED ZOE 24W 4000K IP44 LX czarny Kobi Premium",
                         "Marka": "Kobi Premium", "Strumień świetlny |lm|": "1 540",
                         "Typ czujnika": "ruchu PIR, zmierzchu", "Kolor produktu": "czarny"}}
        res = build_name(kobi_row("LED ZOE LX 24W 4000K CZARNA PREMIUM"), KOBI, {}, feed)
        self.assertEqual(res["nazwa"], "Plafon LED 24W 1540lm 4000K IP44 z czujnikiem ruchu i zmierzchu czarny Zoe Kobi Premium")

    def test_feed_with_other_product_is_ignored(self):
        feed = {"4666": {"_name": "Żarówka LED GS 13W E27 6500K Kobi", "Strumień świetlny |lm|": "1400"}}
        res = build_name(kobi_row("LED GS 13W E27 6000K"), KOBI, {}, feed)
        self.assertNotIn("1400lm", res["nazwa"])
        self.assertTrue(any("inny produkt" in u for u in res["uwagi"]))

    def test_extension_cord_and_psu(self):
        res = build_name(kobi_row("Przedłużacz LINEA 3gn/5m/Zu+W czarny Kobi"), KOBI, {}, {})
        self.assertEqual(res["nazwa"], "Przedłużacz 3 gn. 5m z uziemieniem z wyłącznikiem czarny Linea Kobi")
        res = build_name(kobi_row("ZASILACZ DESKTOP 12V 72W 6,0A"), KOBI, {}, {})
        self.assertEqual(res["nazwa"], "Zasilacz LED desktop 12V 72W Kobi")

    def test_track_spot_degrees_st(self):
        res = build_name(kobi_row("Reflektor szynowy LED NEXTRACK CORE 20W 3CCT 36ST biały Kobi Pro"), KOBI, {}, {})
        self.assertEqual(res["nazwa"], "Reflektor szynowy 20W CCT 36° biały Nextrack Core Kobi Pro")

    def test_caps_noun_does_not_eat_model(self):
        # rzeczownik ze slownika ma inna dlugosc niz tekst - wczesniej "AKUM. 18650" dawalo "Akumulator Kobi 0"
        res = build_name(kobi_row("AKUM. 18650 3,7V 3AH (SOLAR MHC 5W)"), KOBI, {}, {})
        self.assertTrue(res["nazwa"].startswith("Akumulator 18650 3,7V") and res["nazwa"].endswith(" Kobi"))
        res = build_name(kobi_row("SMART SOCKET 16A WIFI"), KOBI, {}, {})
        self.assertEqual(res["nazwa"], "Gniazdko Smart 16A WiFi Kobi")
        res = build_name(kobi_row("KOSTKA HALOGENOWA Z WYSIĘGNIKIEM - K006"), KOBI, {}, {})
        self.assertEqual(res["nazwa"], "Kostka halogenowa z wysięgnikiem K006 Kobi")

    def test_brand_before_preposition(self):
        res = build_name(kobi_row("Rozeta sufitowa do LED KOLINE K2 biała Kobi Pro"), KOBI, {}, {})
        self.assertEqual(res["nazwa"], "Rozeta sufitowa do KOLINE K2 biała Kobi Pro")
        res = build_name(kobi_row("Adapter do szynoprzewodu 3-obwodowy biały Kobi"), KOBI, {}, {})
        self.assertEqual(res["nazwa"], "Adapter do szynoprzewodu 3-obwodowy biały Kobi")
        # przyimek bedacy czescia rodzaju oprawy zostaje w rzeczowniku
        res = build_name(kobi_row("Oprawa do nadbudowania OH12 1xGX53 biała Kobi"), KOBI, {}, {})
        self.assertEqual(res["nazwa"], "Oprawa do nadbudowania 1xGX53 biała OH12 Kobi")

    def test_lumen_not_glued_to_ip(self):
        res = build_name(kobi_row("Kinkiet ogrodowy LED QUERK 2x2W 3000K IP54 550lm czarny LED2B"), KOBI, {}, {})
        self.assertEqual(res["nazwa"], "Kinkiet ogrodowy 2x2W 550lm 3000K IP54 czarny Querk LED2B")

    def test_tube_length_and_socket_fixture(self):
        res = build_name(kobi_row("LED T8 9W 60CM 4000K LED2B"), KOBI, {}, {})
        self.assertEqual(res["nazwa"], "Tuba LED T8 60cm 9W 4000K LED2B")
        # oprawa na wymienne zrodlo nie ma wlasnego strumienia - brak lm to nie blad
        res = build_name(kobi_row("Reflektor szynowy LAITILINE 1xGU10 biały Laitica"), KOBI, {}, {})
        self.assertEqual(res["nazwa"], "Reflektor szynowy 1xGU10 biały Laitiline Laitica")
        self.assertEqual(res["status"], "OK")
        res = build_name(kobi_row("Kinkiet GRANDE 1xR7s Art of Light"), KOBI, {}, {})
        self.assertEqual(res["nazwa"], "Kinkiet 1xR7s Grande Art of Light")

    def test_new_cards_model_first_use_feed_name(self):
        # nowe kartoteki 2026-11-05: nazwa od kodu modelu - typ z nazwy feedu ("Oprawa drogowa", "Oprawa hermetyczna")
        feed = {"4666": {"_name": "Oprawa drogowa LED VESPA PRO 40W 4000K IP66 Kobi Pro", "Marka": "Kobi Professional"}}
        res = build_name(kobi_row("LED VESPA PRO 40W 4000K Lampa uliczna"), KOBI, {}, feed)
        self.assertEqual(res["nazwa"], "Oprawa uliczna LED 40W 4000K IP66 Vespa Kobi Pro")
        feed = {"4666": {"_name": "Oprawa hermetyczna HERMETIC 2x60 IP65 Kobi", "Marka": "Kobi"}}
        res = build_name(kobi_row("HERMETIC 2X60 Oprawa pod tuby led"), KOBI, {}, feed)
        self.assertEqual(res["nazwa"], "Oprawa hermetyczna pod tuby LED 2x60cm IP65 Hermetic Kobi")
        # "Nexus" w polu Marka feedu to Kobi Pro (nazwa i kategoria feedu)
        feed = {"4666": {"_name": "Oprawa liniowa LED NEXLINE1 31W P Kobi Pro", "Marka": "Nexus"}}
        res = build_name(kobi_row("OPRAWA NEXLINE1 31W P"), KOBI, {}, feed)
        self.assertTrue(res["nazwa"].endswith(" NEXLINE1 P Kobi Pro"), res["nazwa"])

    def test_new_cards_old_caps_names(self):
        # plafon na zarowki to nie "Plafon LED"; rodzaj przed modelem; oczka OH-xx z ksztaltem i modelem
        res = build_name(kobi_row("SAMIRA S 2XE27 CHROM"), KOBI, {},
                         {"4666": {"_name": "Plafon SAMIRA S 2xE27 chrom Kobi", "Marka": "Kobi"}})
        self.assertEqual(res["nazwa"], "Plafon 2xE27 chrom Samira S Kobi")
        res = build_name(kobi_row("PURO NERO OPRAWA N/T GIPSOWA 2XGU10 CZARNA"), KOBI, {}, {})
        self.assertEqual(res["nazwa"], "Oprawa natynkowa 2xGU10 gipsowa czarna Puro Nero Kobi")
        res = build_name(kobi_row("OH-20 OCZKO KWADRAT STAŁE CZARNE"), KOBI, {}, {})
        self.assertEqual(res["nazwa"], "Downlight kwadratowy stały czarny OH-20 Kobi")
        res = build_name(kobi_row("SOMA PC 100 SIATKA METAL"), KOBI, {}, {})
        self.assertEqual(res["nazwa"], "Oprawa kanałowa z siatką metalową Soma PC-100 Kobi")
        # POWER BOX to bank energii (decyzja uzytkownika 2026-11-05)
        res = build_name(kobi_row("KOBI POWER BOX 600W"), KOBI, {}, {})
        self.assertEqual((res["nazwa"], res["status"]), ("Bank energii 600W Power Box Kobi", "OK"))
        # "Lampa solarna" nie dostaje drugiego "solarna"
        res = build_name(kobi_row("Lampa solarna LED AURA 3000K LED2B"), KOBI, {}, {})
        self.assertEqual(res["nazwa"].count("solarn"), 1)

    def test_extension_cord_without_earth(self):
        res = build_name(kobi_row("Przedłużacz LINEA G 1gn/20m/Bu/2x1mm pomarańczowy Kobi"), KOBI, {}, {})
        self.assertEqual(res["nazwa"], "Przedłużacz 1 gn. 20m 2x1mm2 bez uziemienia pomarańczowy Linea G Kobi")
        feed = {"4666": {"_name": "Przedłużacz LINEA PRO CUBE 3gn+3USB/1USB-C/Zu+W czarny Kobi", "Marka": "Kobi"}}
        res = build_name(kobi_row("Przedłużacz LINEA PRO CUBE 3gn+3USB/1USB.-C/Zu.+W czarny Kobi"), KOBI, {}, feed)
        self.assertEqual(res["nazwa"], "Przedłużacz 3 gn. + 3USB/1USB-C z uziemieniem z wyłącznikiem czarny Linea Pro Cube Kobi")

    def test_version_number_gets_model_word(self):
        results = []
        for name in ["LED SOLARIS 2 5,5W 4000K", "LED SOLARIS 3 5,5W 4000K",
                     "QUAZAR 11 1XGU10 CZARNA", "QUAZAR 15 1XGU10 CZARNA"]:
            res = build_name(kobi_row(name), KOBI, {}, {})
            res["stara"] = name
            results.append(res)
        add_version_suffixes(results, KOBI)
        # od 2026-11-16 model jest seria za marka (nazwa wlasna), a nie wyroznikiem wersji WIELKIMI
        self.assertEqual(results[0]["nazwa"], "Plafon LED 5,5W 4000K Solaris 2 Kobi")
        self.assertEqual(results[1]["nazwa"], "Plafon LED 5,5W 4000K Solaris 3 Kobi")
        # seria jest juz w nazwie - numer idzie za nia, bez powtorzenia serii
        self.assertEqual(results[2]["nazwa"], "Kinkiet ogrodowy 1xGU10 czarny Quazar 11 Kobi")
        self.assertEqual(results[3]["nazwa"], "Kinkiet ogrodowy 1xGU10 czarny Quazar 15 Kobi")


KANLUX = {**load_yaml(DEFAULT_RULES)["suppliers"]["kanlux"], "params": True}


def kanlux_row(name, code="37246"):
    return {"Nazwa towaru": name[:40], "Nazwa dodatkowa": name[40:], "Kod towaru": f"{code}/KAN",
            "Kod producenta": code}


class KanluxTest(unittest.TestCase):
    def kx(self, name, code, rec, rules=KANLUX):
        return build_name(kanlux_row(name, code), rules, {}, {code: {**rec, "_name": name}} if rec else {})

    def test_replaceable_socket_fixture_from_params(self):
        # NOTUS 418 = 4 swietlowki G13 - moc/lm/K zaleza od wlozonego zrodla, w nazwie trzonek z krotnoscia
        res = self.kx("NOTUS 3 EVG 418 NT Oprawa rastrowa natynkowa", "19731",
                      {"Moc maksymalna [W]": "4 x 18", "Trzonek": "G13", "Kolor": "biały"})
        self.assertEqual(res["nazwa"], "Oprawa rastrowa natynkowa 4xG13 biała Kanlux NOTUS EVG NT")
        self.assertEqual(res["status"], "OK")

    def test_params_lumen_power_and_module(self):
        res = self.kx("SIGO R50 LED E14-WW Źródło światła LED", "22735",
                      {"Moc znamionowa [W]": "6", "Trzonek": "E14", "Temperatura barwowa [K]": "3000",
                       "Użyteczny strumień świetlny źródła światła Φuse [lm]": "490"})
        self.assertEqual(res["nazwa"], "Żarówka LED E14 6W 490lm 3000K R50 Kanlux SIGO")
        # zasilacz: "16-40" to zakres obciazenia, moc z "Moc maksymalna"
        res = self.kx("TRETO IP67 12V 40W Zasilacz CV", "39512",
                      {"Moc maksymalna [W]": "40", "Moc znamionowa [W]": "16-40", "Stopień IP": "67"})
        self.assertEqual(res["nazwa"], "Zasilacz CV 40W 12V IP67 Kanlux TRETO")
        res = self.kx("MODv2 LED 12W-WW Źródło światła LED", "29300",
                      {"Moc znamionowa [W]": "12", "Całkowity znamionowy strumień świetlny [lm]": "1200",
                       "Trzonek": "wolne końce przewodów", "Temperatura barwowa [K]": "3000"})
        self.assertEqual(res["nazwa"], "Moduł LED 12W 1200lm 3000K Kanlux MODv2")
        self.assertEqual(res["status"], "OK")

    def test_withdrawn_and_size_distinguisher(self):
        rules = {**KANLUX, "_withdrawn": {"26835", "26837"}}
        results = []
        for name, code in (("STOBI DLP 350-W Oprawa sufitowa punktowa", "26835"),
                           ("STOBI DLP 450-W Oprawa sufitowa punktowa", "26837")):
            res = build_name(kanlux_row(name, code), rules, {}, {})
            res["stara"] = name
            results.append(res)
        self.assertIn("brak w ofercie kanlux.com (wycofany) - dane tylko z nazwy", results[0]["uwagi"])
        add_version_suffixes(results, rules)
        self.assertEqual(results[0]["nazwa"], "Oprawa sufitowa punktowa STOBI DLP 350 biała Kanlux")
        self.assertEqual(results[1]["nazwa"], "Oprawa sufitowa punktowa STOBI DLP 450 biała Kanlux")

    def test_split_head_tail(self):
        self.assertEqual(split_head_tail("XLED GU10 4,9W-WW Źródło światła LED"),
                         ("XLED GU10 4,9W-WW", "Źródło światła LED"))
        # typ przyklejony do kodu po obcieciu nazwy na 40 znakach
        self.assertEqual(split_head_tail("BILO 6,5W T SMDE14-WWŹródło światła LED"),
                         ("BILO 6,5W T SMDE14-WW", "Źródło światła LED"))
        # nazwa przerobiona na opis - kod i typ z konca
        self.assertEqual(split_head_tail("Rozdzielnica modułowa natynkowa KDB-IP65.-S09T MK2 9P IP65 IK08 "
                                         "odporna na UV kolor szary KDB-IP65-S09T MK2 Rozdzielnica serii"),
                         ("KDB-IP65-S09T MK2", "Rozdzielnica serii"))

    def test_breaker_codes(self):
        self.assertEqual(parse_breaker("KMB10-B16/1"), "B16 1P 10kA")
        self.assertEqual(parse_breaker("KRD6-2/25/100-A"), "2P 25A 100mA typ A 6kA")
        self.assertEqual(parse_breaker("KRD6-4/63/300"), "4P 63A 300mA typ AC 6kA")
        self.assertEqual(parse_breaker("KCP-09-230"), "9A cewka 230V")
        res = build_name(kanlux_row("KMB10-C20/4 Wyłącznik nadmiarowo-prądowy, 4P", "1"), KANLUX, {}, {})
        self.assertEqual(res["nazwa"], "Wyłącznik nadprądowy C20 4P 10kA KMB10-C20/4 Kanlux")

    def test_bulb_params_win_over_name_code(self):
        # WW u zarowek Kanlux bywa 2700K - parametry producenta sa zrodlem prawdy
        params = {"37246": {"_name": "XLEDIMW A60 B22 7W-WW Źródło światła LED", "Moc znamionowa [W]": "7",
                            "Całkowity znamionowy strumień świetlny [lm]": "806", "Temperatura barwowa [K]": "2700",
                            "Trzonek": "B22", "Źródło światła": "A60", "Kolor": "transparentny"}}
        res = build_name(kanlux_row("XLEDIMW A60 B22 7W-WW Źródło światła LED"), KANLUX, {}, params)
        self.assertEqual(res["nazwa"], "Żarówka LED B22 7W 806lm 2700K A60 Kanlux XLEDIMW")
        self.assertEqual(res["status"], "OK")

    def test_bulb_without_params_uses_name_and_needs_lumen(self):
        res = build_name(kanlux_row("IQ-LED A60 9W-CW Źródło światła LED", "1"), KANLUX, {}, {})
        self.assertEqual(res["nazwa"], "Żarówka LED E27 9W 6500K A60 Kanlux IQ-LED")
        self.assertEqual(res["braki"], ["lumen"])
        self.assertEqual(res["status"], "DO_SPRAWDZENIA")

    def test_linear_fixture_from_params(self):
        params = {"33900": {"_name": "AL-SL-NW-MAT-W-NT Oprawa liniowa LED", "Kolor": "biały",
                            "Długość [mm]": "580", "Szerokość [mm]": "60", "Moc maksymalna [W]": "10",
                            "Zintegrowane źródło światła LED": "1", "Strumień świetlny oprawy [lm]": "1300",
                            "Temperatura barwowa [K]": "4000", "Stopień IP": "20"}}
        res = build_name(kanlux_row("AL-SL-NW-MAT-W-NT Oprawa liniowa LED", "33900"), KANLUX, {}, params)
        self.assertEqual(res["nazwa"], "Oprawa liniowa LED 58cm 10W 1300lm 4000K biała Kanlux AL-SL-MAT-NT")

    def test_strip_per_meter(self):
        params = {"46458": {"_name": "LCOBB 10W/M 48IP65-WW Taśma LED", "Długość [mm]": "25000",
                            "Moc znamionowa [W]": "1m = 10", "Strumień świetlny oprawy [lm]": "1m = 860",
                            "Temperatura barwowa [K]": "3000", "Napięcie znamionowe [V]": "48 DC", "Stopień IP": "65"}}
        res = build_name(kanlux_row("LCOBB 10W/M 48IP65-WW Taśma LED", "46458"), KANLUX, {}, params)
        self.assertEqual(res["nazwa"], "Taśma LED 25m 10W/m 860lm/m 3000K 48V IP65 Kanlux LCOBB")

    def test_replaceable_source_has_no_power(self):
        # "Moc maksymalna" przy trzonku GU10 to limit zarowki, nie moc oprawy
        params = {"1": {"_name": "ATL 1F PEND22 GU10 W Projektor na szynoprzewód 1 - fazowy",
                        "Moc maksymalna [W]": "10", "Trzonek": "GU10", "Kolor": "biały"}}
        res = build_name(kanlux_row("ATL 1F PEND22 GU10 W Projektor na szynoprzewód 1 - fazowy", "1"),
                         KANLUX, {}, params)
        self.assertEqual(res["typ"], "oprawa_wymienna")
        self.assertNotIn("10W", res["nazwa"])

    def test_wiring_accessory_color_code(self):
        res = build_name(kanlux_row("DOMO 01-1060-250 sz Łącznik podwójny schodowy", "1"), KANLUX, {}, {})
        self.assertEqual(res["nazwa"], "Łącznik podwójny schodowy DOMO szampański 01-1060-250 Kanlux")
        res = build_name(kanlux_row("LOGI 02-1480-003 kr Ramka potrójna pozioma", "1"), KANLUX, {}, {})
        self.assertEqual(res["nazwa"], "Ramka potrójna pozioma LOGI kremowa 02-1480-003 Kanlux")

    def test_socket_adapters_from_card_params(self):
        # Ekonom podaje tylko grupe "Rozgałęźniki, gniazda, wtyczki" - gniazda i uziemienie z karty kanlux.com
        # (typ E = z uziemieniem, typ F = Schuko), uklad jak rozgalezniki Kobi / Bemko (2026-11-05)
        card = {"Kolor": "biały", "Liczba gniazd": "3", "Rodzaj uziemienia": "typu E", "Stopień IP": "20",
                "Prąd znamionowy [A]": "16"}
        res = self.kx("ZARK KRT/2-1 Rozgałęźniki, gniazda, wtyczki", "2180", card)
        self.assertEqual((res["nazwa"], res["status"]),
                         ("Rozgałęźnik 3 gn. z uziemieniem biały ZARK KRT/2-1 Kanlux", "OK"))
        card = {"Kolor": "czarny", "Rodzaj uziemienia": "typu E", "Stopień IP": "44", "Prąd znamionowy [A]": "16"}
        res = self.kx("GRIP W/US Rozgałęźniki, gniazda, wtyczki", "4863", card)
        self.assertEqual(res["nazwa"], "Wtyczka z uziemieniem 16A IP44 czarna GRIP W/US Kanlux")
        card = {"Kolor": "czarny", "Liczba gniazd": "1", "Rodzaj uziemienia": "typu F", "Stopień IP": "44"}
        res = self.kx("GRIP GP/S Rozgałęźniki, gniazda, wtyczki", "4862", card)
        self.assertEqual(res["nazwa"], "Gniazdo przenośne 1 gn. Schuko IP44 czarne GRIP GP/S Kanlux")

    def test_series_keeps_code_without_color_suffix_in_middle(self):
        # "-S" w "PRO-S" to nie kolor srebrny - kolor jest tylko na koncu kodu
        res = build_name(kanlux_row("PRO-S S 30W NW Oprawa uliczna LED", "1"), KANLUX, {}, {})
        self.assertEqual(res["nazwa"], "Oprawa uliczna LED 30W 4000K Kanlux PRO-S S")


GTV = load_yaml(DEFAULT_RULES)["suppliers"]["gtv"]


def gtv_row(name, code="LD-TEST-10"):
    return {"Nazwa towaru": name[:40], "Nazwa dodatkowa": name[40:], "Kod towaru": f"{code}/GTV",
            "Kod producenta": code}


class GtvTest(unittest.TestCase):
    def test_web_params_fill_gaps_and_withdrawn(self):
        # karta gtv.com.pl uzupelnia brak K; kodu nie ma na stronie = wycofany
        rules = {**GTV, "_web_params": {"LD-BERGA50W-30": {"Temperatura barwowa światła [K]": "4000",
                                                            "Strumień świetlny [lm]": "6500", "_name": "x"}},
                 "_withdrawn": {"AE-CAP60-40"}, "_vocab": frozenset({"PRZEWODEM"}), "_models": frozenset()}
        res = build_name(gtv_row("OPRAWA HERMETYCZNA BERGA LED, 50W, 7000lm, AC220-240V, 50/60 Hz, PF>0,9, "
                                 "RA>80, IP65", "LD-BERGA50W-30"), rules, {})
        # lm z nazwy maja pierwszenstwo przed strona
        self.assertEqual(res["nazwa"], "Oprawa hermetyczna LED 50W 7000lm 4000K IP65 GTV Berga")
        self.assertEqual(res["status"], "OK")
        res = build_name(gtv_row("POKRYWKA I PIERŚCIEŃ DO PRZEDŁUŻACZY ELITE 60mm CHROM", "AE-CAP60-40"), rules, {})
        self.assertIn("brak na gtv.com.pl (wycofany) - dane tylko z nazwy", res["uwagi"])
        # slowo pisane na stronie GTV malymi ("z przewodem") nie zostaje "Przewodem"
        res = build_name(gtv_row("ZŁĄCZKA XC11 DO TAŚM LED 10mm Z PRZEWODEM 2m", "LD-ZTL102M-2N"), rules, {})
        self.assertEqual(res["nazwa"], "Złączka XC11 do taśm LED 10mm z przewodem 2m GTV")

    def test_noun_from_producer_category(self):
        # 2026-10-19 "Oprawa LED nic nie mowi": rodzaj z kategorii gtv.com.pl; bez srednicy z "AC220-240V" / "110°"
        res = build_name(gtv_row("OPRAWA LED ARCCO, 25W, 2250lm, AC220-240V, 50/60 Hz, PF>0,9, RA?80, IP44, IK08, "
                                 "110°, 4000K", "LD-ARC25W-NB"), GTV, {})
        self.assertEqual(res["nazwa"], "Plafon LED 25W 2250lm 4000K IP44 GTV Arcco")
        self.assertEqual(res["typ"], "plafon")
        # ta sama nazwa rodziny w innym typie (reflektor szynowy Sierra) zostaje bez zmian
        res = build_name(gtv_row("OPRAWA LED NA SZYNOPRZEWÓD SIERRA, 8W, 720lm, AC220-240V, 50/60 Hz, PF>0,5, Ra?80, "
                                 "IP20, IK08, 4000K, 48°, BIAŁY", "XLD-SRA08WB-NB"), GTV, {})
        self.assertTrue(res["nazwa"].startswith("Reflektor szynowy LED 8W"))

    def test_brand_and_model_last_description_stays(self):
        # 2026-10-19 (zasady z Ledvance): marka + model serii na koncu, opis z serii ("z akumulatorem") zostaje
        res = build_name(gtv_row("NAŚWIETLACZ LED Z AKUMULATOREM ALLEDO 10W, 800lm, AC220-240V, 50/60 Hz, PF>0,5, "
                                 "RA>80, IP54, 120°, 6400K, ZIELONY", "LD-ALCX10W-64"),
                         {**GTV, "_vocab": frozenset({"AKUMULATOREM"})}, {})  # pisownia ze stron GTV
        self.assertEqual(res["nazwa"], "Naświetlacz LED z akumulatorem 10W 800lm 6400K IP54 zielony GTV Alledo")
        # marka wielowyrazowa wraca na koniec za dopiskami
        res = {"nazwa": "Naświetlacz LED 150W GTV G-TECH CIRCLE", "marka": "GTV G-TECH", "po_marce": []}
        move_brand_tail([res], GTV)
        self.assertEqual(res["nazwa"], "Naświetlacz LED 150W CIRCLE GTV G-TECH")

    def test_gtv_name_fixes(self):
        res = build_name(gtv_row("TAŚMA FLASH RGBW, 300 LED, max. 42W, BEZ ŻELU 10mm, ROLKA 5m, 12V", "LD-RGBW-300-20"), GTV, {})
        self.assertEqual(res["nazwa"], "Taśma LED 5m 300LED 42W RGBW 12V GTV Flash")  # "max." przy tasmie to jej moc
        res = build_name(gtv_row("OPRAWA LED COB JOVITA OKRĄGŁA, 3W, 12V DC, 1X COB, CIEPŁY BIAŁY, 150cm PRZEWÓD "
                                 "Z MINIAMP (2 WKRĘTY)", "LD-OJVCOB-CB"), GTV, {})
        self.assertEqual(res["nazwa"], "Oprawa meblowa LED 150cm 3W ciepła biała okrągła GTV COB Jovita")
        res = build_name(gtv_row("OPRAWA LED POINT REGULOWANA, 12V DC, 1.2W, 1x3030 LED, 3000K, CZARNA, PRZEWÓD 1,5m",
                                 "LD-POWW-20"), GTV, {})
        self.assertNotIn(" 5m", res["nazwa"])  # "1,5m" to nie "5m"
        res = build_name(gtv_row("ZAPINKA METALOWA DO OPRAWY HAGEN, BOCZNA", "LD-HAG236KL-ME"), GTV, {})
        self.assertTrue(res["nazwa"].startswith("Zapinka metalowa") and res["nazwa"].endswith(" GTV"))

    def test_hogert_brand_after_noun_phrase(self):
        self.assertEqual(build_name(gtv_row("Palnik gazowy", "HT2C507"), GTV, {})["nazwa"], "Palnik gazowy Högert")
        self.assertEqual(build_name(gtv_row("Cyfrowy miernik uniwersalny", "HT1E601"), GTV, {})["nazwa"],
                         "Cyfrowy miernik uniwersalny Högert")
        self.assertEqual(build_name(gtv_row("Brzeszczot do wyrzynarki do czystych cięć drewna, 10TPI, 100 mm, 5 szt.",
                                            "HT6D665"), GTV, {})["nazwa"],
                         "Brzeszczot do wyrzynarki do czystych cięć drewna, 10TPI, 100 mm, 5 szt. Högert")

    def test_decor_fixture_type_words_and_sockets(self):
        # "max.40W" przy oprawie na zarowke to limit zrodla, nie moc oprawy
        res = build_name(gtv_row("OPRAWA SUFITOWA MARITE, 7108, AC220-240V, 50/60Hz, 1*E27, max.40W, IP20, "
                                 "ŚREDNICA 26 cm, POJEDYNCZA, BRĄZOWA"), GTV, {})
        self.assertEqual(res["typ"], "oprawa_dekor")
        self.assertEqual(res["nazwa"], "Oprawa sufitowa Marite 26cm 1xE27 IP20 brązowa GTV")

    def test_led_fixture_without_space_after_comma(self):
        res = build_name(gtv_row("OPRAWA AWARYJNA TERNO-1 LED,3W,250lm,AC220-240V,50-60Hz,IP65,1h,6400K,"
                                 "DWUZADANIOWA, PRZYCISK TEST, bez CNBOP"), GTV, {})
        self.assertEqual(res["nazwa"], "Oprawa awaryjna LED 3W 250lm 6400K IP65 1h dwuzadaniowa GTV TERNO-1")

    def test_light_colour_words_are_not_housing_colour(self):
        res = build_name(gtv_row("TAŚMA FLASH 2835, 300 LED ZIMNY BIAŁY, 6W/m WODOODPORNA 8mm, ROLKA 5m, 12V"), GTV, {})
        self.assertEqual(res["nazwa"], "Taśma LED 5m 300LED wodoodporna 6W/m zimna biała 12V GTV Flash 2835")

    def test_pdf_by_code_fills_missing_lumen_per_meter(self):
        name = "TAŚMA LED COB 24V, 480LED/m, 8mm, 10W/m, BEZ ŻELU, ZIMNY BIAŁY, 50m"
        texts = {"kat.pdf": ["LD-COB24V-20-ZBPQN-50", "24 VDC", "10 W/1 m", "950 lm/m", "6500 K",
                             "LD-COB24V-20-CBPQN-50", "950 lm/m"]}
        pdf = lumens_from_pdf_codes(texts, {"LD-COB24V-20-ZBPQN-50"})
        res = build_name(gtv_row(name, "LD-COB24V-20-ZBPQN-50"), GTV, pdf)
        self.assertIn("950lm/m 6500K", res["nazwa"])

    def test_lumen_from_efficacy(self):
        res = build_name(gtv_row("LAMPA ULICZNO-PARKOWA GALI, LED, 50W, 165lm/W, 4000K, AC180-250V, IP65"), GTV, {})
        self.assertEqual(res["nazwa"], "Oprawa uliczna LED 50W 8250lm 4000K IP65 GTV Gali")

    def test_hogert_tools_and_workwear(self):
        res = build_name(gtv_row("Wiertło do drewna 7 x 109 mm", "HT6D527"), GTV, {})
        self.assertEqual(res["nazwa"], "Wiertło do drewna 7 x 109 mm Högert")
        self.assertEqual(res["status"], "OK")
        res = build_name(gtv_row("EDGAR II spodnie ochronne ciemnobeżowe S-48", "HT5K935-S"), GTV, {})
        self.assertEqual(res["nazwa"], "Spodnie ochronne ciemnobeżowe EDGAR II S-48 Högert")


class GtvReview20261105Test(unittest.TestCase):
    """Przeglad WARNING GTV 2026-11-05: opis akcesoriow i typow bez schematu, przedluzacze wg schematu V3."""

    def name(self, text, code="LD-TEST-10"):
        return build_name(gtv_row(text, code), {**GTV, "_vocab": frozenset({"WYSOKI"})}, {})["nazwa"]  # pisownia GTV

    def test_parenthesis_is_not_cut_by_word_limit(self):
        self.assertEqual(self.name("ZESTAW GLAX NAKŁADANY MINI BIAŁY - (PROFIL 2mb + OSŁONKA MLECZNA + 2 ZAŚLEPKI "
                                   "KOŃCOWE + 2 UCHWYTY MONTAŻOWE)", "PA-ZGLAXMNKOMZ-10"),
                         "Zestaw GLAX nakładany MINI biały (profil 2mb + osłonka mleczna + 2 zaślepki końcowe + "
                         "2 uchwyty montażowe) GTV")
        self.assertEqual(self.name("ZAŚLEPKA DO PROFILU LED GLAX MINI Z KOŁNIERZEM WYSOKI BIAŁA (10 szt. W BLISTRZE)",
                                   "PA-ZASGLAXMW-10"),
                         "Zaślepka do profilu LED GLAX MINI z kołnierzem wysoki biała (10 szt.) GTV")

    def test_unit_after_number_stays(self):
        self.assertEqual(self.name("OSŁONKA TRANSPARENTNA DO PROFILU LED GLAX 12 m", "PA-OSTRGLAX12M-00"),
                         "Osłonka transparentna do profilu LED GLAX 12m GTV")

    def test_case_of_phase_usb_and_count(self):
        self.assertEqual(self.name("ZAŚLEPKA KOŃCOWA DO SZYNY 1-FAZOWEJ, 2-LINIOWEJ, BIAŁY", "XLT-1FAC70-00"),
                         "Zaślepka końcowa do szyny 1-fazowej 2-liniowej biały GTV")
        self.assertEqual(self.name("PRZEDŁUŻACZ BIURKOWY WPUSZCZANY PRESTINO, 2xGNIAZDO SCHUKO, USB A+C, RJ45, HDMI, "
                                   "PRZEWÓD 1,5m, BIAŁY", "AE-PB2SPRES-10"),
                         "Przedłużacz biurkowy wpuszczany Prestino 2x gniazdo SCHUKO USB A+C RJ45 HDMI przewód 1,5m "
                         "biały GTV")
        # kody zostaja w calosci: "10xE27", nie "10x e27"
        self.assertIn("10xE27", self.name("GIRLANDA OGRODOWA 10m, AC220-240V, 50-60Hz,10xE27, 2x0,75mm", "AE-GIR10-10"))

    def test_lost_cp1250_signs_and_glued_lumen(self):
        # przeglad 2026-11-05: "A60 820lm" to 820 lm (nie 60820), znaki spoza cp1250 w Ekonomie jako "?"
        self.assertIn(" 820lm ", self.name("Żarówka LED GTV 10W E27 A60 820lm 3000K AC220-240V ściemnialne",
                                            "LD-SPCA60-10W"))
        self.assertEqual(self.name("Smarownica 400 cm?", "HT8G905"), "Smarownica 400 cm3 Högert")
        self.assertEqual(self.name('Przedłużka 1/4", 76.0 mm??, ??CrV', "HT1A702"), 'Przedłużka 1/4", 76.0 mm, CrV Högert')
        self.assertEqual(self.name("Szczypce do zaciskania tulejek cienkościennych 6-16 mm?", "HT1P162"),
                         "Szczypce do zaciskania tulejek cienkościennych 6-16 mm2 Högert")

    def test_extension_cords_follow_schema_v3(self):
        self.assertEqual(self.name("PRZEDŁUŻACZ DOMOWY 3 GNIAZDA, 1,5m 3X1,00mm2", "AE-PF3G1M-00"),
                         "Przedłużacz 3 gn. 1,5m 3x1,0mm2 GTV")
        self.assertEqual(self.name("PRZEDŁUŻACZ SCHUKO 4 GNIAZDA, 3m", "AE-PS4G3M-00"), "Przedłużacz 4 gn. 3m Schuko GTV")
        self.assertEqual(self.name("PRZEDŁUŻACZ BĘBNOWY 3X1,5mm2, 30m, 4X SCHUKO", "AE-SBEBEN-30"),
                         "Przedłużacz bębnowy 4 gn. 30m 3x1,5mm2 Schuko GTV")
        self.assertEqual(self.name("Przedłużacz bębnowy w izolacji gumowej 30 m 3x2.5 mm", "HT1E735"),
                         "Przedłużacz bębnowy 30m 3x2,5mm2 guma Högert")
        res = build_name(gtv_row("PRZEDŁUŻACZ OGRODOWY 2X1,0mm2 UNI, 10m", "AE-POGRODUN-10"), GTV, {})
        self.assertEqual((res["nazwa"], res["status"]), ("Przedłużacz ogrodowy 10m 2x1,0mm2 Uni GTV", "OK"))


BEMKO = load_yaml(DEFAULT_RULES)["suppliers"]["bemko"]


def bemko_row(name, code):
    return {"Nazwa towaru": name[:40], "Nazwa dodatkowa": name[40:], "Kod producenta": code,
            "Kod towaru": f"{code}/BMK"}


class BemkoTest(unittest.TestCase):
    def test_floodlight_series_and_sensor(self):
        res = build_name(bemko_row("NAŚWIETLACZ LED SMD KASTEL2 10W 4000K 800LM IP65 CZARNY CZ. PIR", "C64-S-LFK2-010BL-4K"),
                         BEMKO, {})
        self.assertEqual(res["nazwa"], "Naświetlacz LED 10W 800lm 4000K IP65 z czujnikiem PIR czarny Bemko Kastel 2")
        self.assertEqual(res["status"], "OK")

    def test_brand_from_end_of_name(self):
        res = build_name(bemko_row("PLAFONIERA LED CELIA 10-20W 3000-4000K 2300LM IP54 Luxmena", "C35-PLC-290-200-CCT"),
                         BEMKO, {})
        self.assertEqual(res["nazwa"], "Plafon LED 10-20W 2300lm 3000-4000K IP54 Luxmena Celia")

    def test_sensor_from_code_when_name_lacks_it(self):
        res = build_name(bemko_row("PLAFONIERA TOFIR PHR 2X20W E27 IP44 FI-35 PC+AL", "C30-PHR220-MS"), BEMKO, {})
        self.assertEqual(res["nazwa"], "Plafon Tofir PHR 2xE27 fi35cm z czujnikiem mikrofalowym IP44 Bemko")

    def test_breaker_and_code_check(self):
        res = build_name(bemko_row("WYŁĄCZNIK NADPRĄDOWY 10KA 1P B 06A VDE SCHELINGER Schelinger", "A02-G9-1P-B06"),
                         BEMKO, {})
        self.assertEqual(res["nazwa"], "Wyłącznik nadprądowy B6 1P 10kA Schelinger")
        self.assertEqual(res["status"], "OK")
        # kod mowi 63A 100mA, nazwa w Ekonomie 25A 300mA - nazwy nie poprawiamy, idzie do sprawdzenia
        res = build_name(bemko_row("WYŁĄCZNIK RÓŻNICOWOPRĄDOWY AC 2P 25A 300mA SCHELINGER Schelinger", "A03-V7-2-63-100"),
                         BEMKO, {})
        self.assertEqual(res["status"], "DO_SPRAWDZENIA")
        self.assertTrue(any("63A 100mA" in u for u in res["uwagi"]))

    def test_extension_cord_and_empty_t8_housing(self):
        res = build_name(bemko_row('PRZEDŁUŻACZ NEW 1.0 3 GN. Z "0" Z WYŁ. 1.0MM X 1.5M', "B04-NA3GW1M"), BEMKO, {})
        # schemat nazwy przedluzaczy V3 (2026-11-05): gniazda, dlugosc, przekroj, uziemienie (Schuko zamiast
        # "z uziemieniem"), wylacznik
        self.assertEqual(res["nazwa"], "Przedłużacz 3 gn. 1,5m 1,0mm2 z uziemieniem z wyłącznikiem Bemko")
        res = build_name(bemko_row('PRZEDŁUŻACZ NA SCHUKO 3X 2P+Z+WYŁ. 1.5MM X 1.5M', "B06-NB3GSW1M"), BEMKO, {})
        self.assertEqual(res["nazwa"], "Przedłużacz 3 gn. 1,5m 1,5mm2 Schuko z wyłącznikiem Bemko")
        # bebnowe: gniazda, dlugosc, przekroj, material przewodu, hermetycznosc, uziemienie / Schuko przy koncu
        res = build_name(bemko_row('PRZEDŁUŻACZ BĘBNOWY 4 GN. +"0" 20m 3x1mm2', "B15-DG-4ZRF04-20"), BEMKO, {})
        self.assertEqual(res["nazwa"], "Przedłużacz bębnowy 4 gn. 20m 3x1mm2 z uziemieniem Bemko")
        res = build_name(bemko_row('PRZEDŁUŻACZ BĘBNOWY SCHUKO 4 GN. +"0" 15m 3x1,5mm2', "B16-DG-4ZRS05-15"), BEMKO, {})
        self.assertEqual(res["nazwa"], "Przedłużacz bębnowy 4 gn. 15m 3x1,5mm2 Schuko Bemko")
        res = build_name(bemko_row('PRZEDŁUŻACZ BĘBNOWY 4 GN. +"0" IP44 25m H05RR-F 3x1,5mm2', "B15-X"), BEMKO, {})
        self.assertEqual(res["nazwa"], "Przedłużacz bębnowy 4 gn. 25m 3x1,5mm2 guma IP44 z uziemieniem Bemko")
        res = build_name(bemko_row('OPRAWA HERMETYCZNA 1X18W PC "EMPTY"', "S-C10-E-OHB118PC"), BEMKO, {})
        self.assertEqual(res["nazwa"], "Oprawa hermetyczna 60cm 1xT8 klosz PC Bemko")

    def test_family_brand_propagation(self):
        rows = [bemko_row("NAŚWIETLACZ LED CAMPUS 200W 3000K 30000lm IP66 90X90 SZARY Luxmena", "C66-LFC-200DG-3K-9090"),
                bemko_row("NAŚWIETLACZ LED CAMPUS 60W 3000K 8400lm IP66 90X90 SZARY", "C66-LFC-060DG-3K-9090")]
        results = []
        for r in rows:
            res = build_name(r, BEMKO, {})
            res["row"] = r
            results.append(res)
        propagate_family_brand(results, BEMKO)
        self.assertEqual(results[1]["nazwa"], "Naświetlacz LED 60W 8400lm 3000K IP66 szary Luxmena Campus")

    def test_catalog_card_rows_by_code(self):
        # karta Bemko: kod | EAN | opakowanie | kolor | K | W | lm - 4000 to barwa, nie strumien
        texts = {"karta.pdf": ["C22-MLK-400-4K-BL", "5900280978127", "1/20", "czarny", "4000", "8", "600",
                               "C26-LOR-40BL-4K", "5900280984456", "1/10", "czarny", "4000", "100", "16000"]}
        pdf = lumens_from_pdfs(texts, {"C22-MLK-400-4K-BL": (8.0, {4000}), "C26-LOR-40BL-4K": (100.0, {4000})},
                               by_code=True)
        self.assertEqual(pdf["C22-MLK-400-4K-BL"]["lm"], [600])
        self.assertEqual(pdf["C26-LOR-40BL-4K"]["lm"], [16000])
        res = build_name(bemko_row("KINKIET LED KEIRA 8W 4000K 40CM IP44 CZARNY Bemko", "C22-MLK-400-4K-BL"), BEMKO, pdf)
        self.assertEqual(res["nazwa"], "Kinkiet LED 40cm 8W 600lm 4000K IP44 czarny Bemko Keira")

    def test_accessory_brand_before_preposition_keeps_repeated_words(self):
        res = build_name(bemko_row("UCHWYT REGULOWANY DO OPRAW Z UCHWYTEM FI60mm NA SŁUPACH Z FI60mm", "C82-KUL1-60-UR"),
                         BEMKO, {})
        self.assertEqual(res["nazwa"], "Uchwyt regulowany do opraw z uchwytem fi60mm na słupach z fi60mm Bemko")
        res = build_name(bemko_row("ZESTAW MONTAŻOWY DO PANELI LED - KOMPLET Z LINKĄ 100 cm", "S-C71-PLZ-NTL"), BEMKO, {})
        self.assertEqual(res["nazwa"], "Zestaw montażowy do paneli LED komplet z linką 100cm Bemko")

    def web_rules(self, cards):
        return {**BEMKO, "_web_params": {code: {"_name": name, "Marka": brand, "_url": "https://bemko.pl/x"}
                                         for code, (name, brand) in cards.items()}}

    def build_with_card(self, name, code, cards):
        rules = self.web_rules(cards)
        r = bemko_row(name, code)
        res = build_name(r, rules, {})
        res.update(row=r, stara=full_name(r))
        return apply_web_card(res, r, rules, {}, None)

    def test_web_card_fixes_name_of_neighbour_product(self):
        # Ekonom: kartoteka 450-360 skopiowana z 350-240 ("24W 2160LM"), karta bemko.pl: 36W 3250LM
        res = self.build_with_card("PLAFONIERA LED PLATO 24W 3000K 2160LM IP44 Bemko", "C37-PLA-450-360-3K",
                                   {"C37-PLA-450-360-3K": ("PLAFONIERA LED PLATO 36W 3000K 3250LM IP44", "Bemko")})
        self.assertEqual(res["nazwa"], "Plafon LED 36W 3250lm 3000K IP44 Bemko Plato")
        self.assertEqual(res["status"], "WARNING")
        self.assertTrue(any("moc 24W / 36W" in u for u in res["uwagi"]))
        # kod mowi 63A 100mA, karta tez - nazwa z karty zamiast DO_SPRAWDZENIA
        res = self.build_with_card("WYŁĄCZNIK RÓŻNICOWOPRĄDOWY AC 2P 25A 300mA SCHELINGER Schelinger", "A03-V7-2-63-100",
                                   {"A03-V7-2-63-100": ("WYŁĄCZNIK RÓŻNICOWOPRĄDOWY AC 2P 63A 100mA SCHELINGER",
                                                        "Schelinger")})
        self.assertEqual(res["nazwa"], "Wyłącznik różnicowoprądowy 2P 63A 100mA typ AC Schelinger")
        self.assertEqual(res["status"], "WARNING")

    def test_web_card_brand_for_name_without_brand(self):
        res = self.build_with_card("NAŚWIETLACZ LED CAMPUS 60W 3000K 8400lm IP66 90X90 SZARY", "C66-LFC-060DG-3K-9090",
                                   {"C66-LFC-060DG-3K-9090": ("NAŚWIETLACZ LED CAMPUS 60W 3000K 8400lm IP66 90X90 SZARY",
                                                              "Luxmena")})
        self.assertEqual(res["nazwa"], "Naświetlacz LED 60W 8400lm 3000K IP66 szary Luxmena Campus")
        self.assertTrue(any("z karty producenta" in u for u in res["uwagi"]))


class BemkoZestawienie20261103Test(unittest.TestCase):
    """Zestawienie Bemko "kart" z 2026-11-03 (340 nowych kartotek)."""

    def name(self, text, code):
        return build_name(bemko_row(text, code), BEMKO, {})["nazwa"]

    def test_short_head_of_joined_name_gets_space(self):
        # pierwsza czesc krotsza niz 40 znakow = ciecie na spacji, ktora zestawienie zgubilo
        header = "LP;Kod towaru;Nazwa towaru;J.M.;Kod producenta\r\n"
        with tempfile.TemporaryDirectory() as tmp:
            src = Path(tmp) / "in.csv"
            src.write_bytes((header + '1;"a";"PANEL BACKLIGHT SUPERIOR BLS 34W 4000K | 5000LM IP44";"szt";"a"\r\n'
                             '2;"b";"CZUJNIK MIKROFALOWY 1200W 360* IP20 OKRĄ | GŁY BIAŁY";"szt";"b"\r\n'
                             '3;"c";"Uchwyt do końcówek wkrętakowych 1/4"", 6 | 0 mm";"szt";"c"\r\n')
                            .encode("cp1250"))
            _, rows = read_csv(src)
        self.assertEqual(full_name(rows[0]), "PANEL BACKLIGHT SUPERIOR BLS 34W 4000K 5000LM IP44")
        self.assertEqual(full_name(rows[1]), "CZUJNIK MIKROFALOWY 1200W 360* IP20 OKRĄGŁY BIAŁY")
        # liczba przecieta w srodku zostaje cala (GTV: Ekonom liczy cudzyslow podwojnie)
        self.assertEqual(full_name(rows[2]), 'Uchwyt do końcówek wkrętakowych 1/4", 60 mm')

    def test_typos_in_ekonom_names(self):
        self.assertEqual(self.name("OPRAWA AWARYJNA 03 ALWIR 3 35W 4000K 600 0LM IP66 2H CNBOP",
                                   "C90-AB-HLA3-120-350-4K-AW2H-C"),
                         "Oprawa awaryjna LED 35W 6000lm 4000K IP66 2h CNBOP Bemko 03 Alwir 3")
        self.assertEqual(self.name("HIGHBAY LED VOYAGER 100W 4000K 14000LMCRI80 IP65 120° BEMKO", "C29-HBV-100"),
                         "Oprawa High Bay LED 100W 14000lm 4000K IP65 120° Bemko Voyager")
        self.assertEqual(self.name("PANEL BACKLIGHT BLO NT OKRĄGŁY 34W 4400L M 4000K IP40 FI60 BIAŁY", "C72-BLO-NT"),
                         "Panel LED backlight 34W 4400lm 4000K IP40 okrągły fi60cm biały Bemko Blo NT")

    def test_new_types(self):
        self.assertEqual(self.name("ŻARNIK HALOGENOWY 100W 118MM 0930", "J118 100W"),
                         "Żarnik halogenowy R7s 100W 118mm Bemko")
        self.assertEqual(self.name("MR11 ŻARÓWKA HALOG. G4 10W 12V 30ST.0316", "MR11 G4 10W"),
                         "Żarówka halogenowa MR11 G4 10W 12V 30° Bemko")
        self.assertEqual(self.name("ŚWIETLÓWKA T4 8W 4000K 340 mm", "T4 8W"), "Świetlówka T4 8W 4000K 340mm Bemko")
        # ramki roznia sie wymiarem, ktory `inne` wycinalo
        self.assertEqual(self.name("RAMKA MONTAŻOWA DO PANELU LED 30x120 BIAŁA ECO WYS. 68mm", "C71-RNT123-WH-ECO2"),
                         "Ramka montażowa do panelu LED 30x120 Eco wys. 68mm biała Bemko")
        self.assertEqual(self.name("ZASILACZ DO PANELI LEDOWYCH PLM 40W WH (950 mA)", "E88-LD-PLM-400"),
                         "Zasilacz do paneli LED 40W PLM 950mA Bemko")
        self.assertEqual(self.name("OPRAWA STADIONOWA LED STADIUM 500W 5000K 80000LM IP66 60x60 Luxmena",
                                   "C69-STD2-S-500BL-5K-6060"),
                         "Oprawa stadionowa LED 500W 80000lm 5000K IP66 Luxmena Stadium")
        self.assertEqual(self.name("OPRAWA RAMA LED 40W 4000LM 4000K 60x60 PLAZA BIAŁA", "C73-FLP-066-400-4K-WH"),
                         "Panel LED 60x60 40W 4000lm 4000K biały Bemko Plaza")
        self.assertEqual(self.name("KAFLER 2X36W PAR OPRAWA NASTROPOWA EVG", "K-PAR-236"),
                         "Oprawa rastrowa 2xT8 natynkowa 120cm Kafler")

    def test_emergency_panel_keeps_light_parameters(self):
        # panel z modulem awaryjnym niesie lm, barwe i wymiar; mala oprawa awaryjna - jak dotad
        self.assertEqual(self.name("OPRAWA AWARYJNA 08 MINTAL 40W 4000K 4000LM 120x30 BIAŁA 2H CNBOP",
                                   "C90-AB-BLM-123-400-4K-WH-AW2H-C"),
                         "Oprawa awaryjna LED 120x30 40W 4000lm 4000K 2h CNBOP biała Bemko 08 Mintal")
        self.assertEqual(self.name("OPRAWA AWARYJNA MALER NATYNKOWA 2W IP20 AUTOTEST 3H CNBOP", "C90-MLR"),
                         "Oprawa awaryjna LED 2W natynkowa autotest 3h CNBOP IP20 Bemko Maler")

    def test_plato_without_lumen_and_24v(self):
        res = build_name(bemko_row("PLAFONIERA LED PLATO 12W 4000K IP44 BEMKO", "C37-PLA-260-120-4K"), BEMKO, {})
        self.assertEqual(res["typ"], "plafon_led")
        self.assertIn("lumen", res["braki"])
        self.assertEqual(self.name("PLAFONIERA PLATO 1xE27 IP44 24V DC CZ. MIKR.", "C37-PLA-E27-24V-MS"),
                         "Plafon Plato 1xE27 24V DC z czujnikiem mikrofalowym IP44 Bemko")

    def test_card_brand_does_not_glue_to_short_name(self):
        rules = {**BEMKO, "_web_params": {"B50-SES02BL": {"_name": "CZUJNIK RUCHU 800W 360* CZARNY", "Marka": "Bemko",
                                                          "_url": "https://bemko.pl/x"}}}
        r = {"Nazwa towaru": "CZUJNIK RUCHU 800W 360* CZARNY", "Nazwa dodatkowa": "", "Kod producenta": "B50-SES02BL",
             "Kod towaru": "B50-SES02BL/BMK"}
        res = build_name(r, rules, {})
        res.update(row=r, stara=full_name(r))
        res = apply_web_card(res, r, rules, {}, None)
        self.assertEqual(res["nazwa"], "Czujnik ruchu 800W 360° czarny Bemko")

    def test_card_does_not_replace_cct(self):
        rules = {**BEMKO, "_web_params": {"C72-BLG-044-30B-CCT-9016A": {
            "_name": "PANEL BACKLIGHT GRANDER BLG 26W/30W/34W CCT (2990-4250lm) IP44 40x40 BIAŁY", "Marka": "Luxmena",
            "Barwa światła": "3000K / 4000K / 6500K", "_url": "https://bemko.pl/x"}}}
        res = build_name(bemko_row("PANEL BACKLIGHT GRANDER BLG 26W/30W/34W CCT (2990-4250lm) IP44 40x40 BIAŁY",
                                   "C72-BLG-044-30B-CCT-9016A"), rules, {})
        self.assertIn(" CCT ", res["nazwa"])

    def test_upper_case_card_names_are_not_spelling_source(self):
        # jedno "mini" na koncu nazwy karty robilo z "BIAŁY" nazwe modelu
        with tempfile.TemporaryDirectory() as tmp:
            Path(tmp, "a.json").write_text('{"code": "a", "found": true, "name": '
                                           '"CZUJNIK MIKROFALOWY 1200W IP20 PROSTOKĄTNY BIAŁY mini", "params": {}}',
                                           encoding="utf-8")
            _, _, vocab, models = load_web_params(Path(tmp))
        self.assertEqual((vocab, models), (frozenset(), frozenset()))


class BemkoReview20261104Test(unittest.TestCase):
    """Przeglad nazw Bemko z 2026-11-04: klasa ogranicznika, "+", jednostki i kody modeli."""

    def name(self, text, code):
        return build_name(bemko_row(text, code), BEMKO, {})["nazwa"]

    def test_surge_class_without_plus(self):
        self.assertEqual(self.name("OGRANICZNIK PRZEPIĘĆ 3P 8kA DC 1000V KLASY T1T2 SCHELINGER", "A53-SPD-MOV-DC-3P-T1T2"),
                         "Ogranicznik przepięć T1+T2 3P 8kA 1000V DC Schelinger")

    def test_plus_and_layout_kept_in_accessory(self):
        self.assertEqual(self.name("STYCZNIK MODUŁOWY SC301 4P 40A TYP 40 (4NO + 0NC) SCHELINGER", "A62-SC-301-4P-40-40"),
                         "Stycznik modułowy SC301 4P 40A typ 40 (4NO+0NC) Schelinger")
        self.assertEqual(self.name("BLOK ROZDZIELCZY 30 ZACISKÓW 2x15 (11x5,3mm + 2x7,5mm + 2x9,1mm)", "A18-MDB-215"),
                         "Blok rozdzielczy 30 zacisków 2x15 (11x5,3mm+2x7,5mm+2x9,1mm) Bemko")

    def test_units_lower_case(self):
        self.assertEqual(self.name("ZAŚLEPKA DO SZYN ŁĄCZENIOWYCH 1-FAZOWYCH 10MM2 PACZKA 10SZT. Bemko", "A19-ZSL1F"),
                         "Zaślepka do szyn łączeniowych 1-fazowych 10mm2 paczka 10 szt. Bemko")
        self.assertEqual(self.name("SŁUP OŚWIETLENIOWY SKŁADANY 3M 160X160MM FI60MM CZARNY", "C84-SSO-S3M-160060-BL"),
                         "Słup oświetleniowy składany 3m 160x160mm fi60mm czarny Bemko")
        self.assertEqual(self.name("ROZETA SUFITOWA QUALIS 100x50x50 mm GRAFIT", "C27-QLS-ZMR2-A7016-3W"),
                         "Rozeta sufitowa Qualis 100x50x50mm grafit Bemko")

    def test_model_codes_stay_upper(self):
        self.assertEqual(self.name("PRZEDŁUŻKA 1M DO SKŁADANEGO SŁUPA OŚWIETLENIOWEGO SSO SZARA Bemko", "C84-SSO-A1M-DG"),
                         "Przedłużka 1m do składanego słupa oświetleniowego SSO szara Bemko")
        self.assertEqual(self.name("NAŚWIETLACZ LED LFX 20W 4000K 1800LM IP65 BIAŁY Bemko", "C65-LFX-020WH-4K"),
                         "Naświetlacz LED 20W 1800lm 4000K IP65 biały Bemko LFX")


if __name__ == "__main__":
    unittest.main()


class LedvanceReview20261021Test(unittest.TestCase):
    """Przeglad nazw Ledvance partiami (2026-10-21): dekor / Smart+, akcesoria, elektronika, tuby."""

    def test_excel_dates_in_second_name_part(self):
        # Excel zamienil koncowki kodow tasm na daty: "7/5" -> "05-lip", "27/5" -> "27-maj", "40/5" -> "maj-40"
        self.assertEqual(undo_excel_date("05-lip"), "7/5")
        self.assertEqual(undo_excel_date("27-maj"), "27/5")
        self.assertEqual(undo_excel_date("maj-40"), "40/5")
        self.assertEqual(undo_excel_date("5/5/IP67"), "5/5/IP67")
        res = build_name({"Nazwa towaru": "Pasek LED Performance LS PFM-1000/COB/92", "Nazwa dodatkowa": "05-lip",
                          "Kod producenta": ""}, RULES, {})
        self.assertEqual(res["nazwa"], "Taśma LED 5m 1000lm/m 2700K COB Ledvance Performance")
        res = build_name(row("Pasek LED Superior LS SUP-2000/TW/927-965/5"), RULES, {})
        self.assertEqual(res["nazwa"], "Taśma LED 5m 2000lm/m TW Ledvance Superior")  # bylo "927m"

    def test_smart_devices_and_cameras(self):
        cases = {
            "Termostat grzejnikowy SMART+ z technologią WiFi": "Termostat grzejnikowy Smart+ WiFi Ledvance",
            "Czujnik SMART+ wykrywający wodę z technologią WiFi": "Czujnik zalania Smart+ WiFi Ledvance",
            # kamera z oswietleniem = rodzaj oprawy z kamera (decyzja 2026-10-21); W / lm / K z karty - LedvanceDecisions
            "Oprawa SMART OUTD WIFI ENTRANCENUMBER CAM": "Kinkiet LED z kamerą Smart+ WiFi Ledvance Entrance Number",
            "Oprawa SMART+ WIFI CEILING FAN RD REM TW WT":
                "Wentylator sufitowy Smart+ WiFi TW z pilotem biały Ledvance Fan Round",
            "Oprawa SUNATHOME PANAN DESK DISC CLIP TWBK": "Lampa biurkowa Smart+ TW czarna Ledvance SUN@HOME Panan Disc Clip",
        }
        for old, new in cases.items():
            self.assertEqual(build_name(row(old), RULES, {})["nazwa"], new)

    def test_decor_pendants_click_and_parts(self):
        cases = {
            "Oprawa DECOR TETRAGON 2XPEND 23W830 CLICKDLEDV":
                "Lampa wisząca 2-punktowa 23W 3000K Click-DIM Ledvance Decor Tetragon",
            "Oprawa Obris Tube Frost Glass 40CM Click CCT DIM IP44 Black":
                "Plafon 40cm Click-CCT IP44 czarny Ledvance Orbis Tube Frost Glass",
            "Oprawa ORBIS PARIS 300 mm 1XE27 BW": "Plafon fi300 1xE27 brązowy Ledvance Orbis Paris",
            "Oprawa Endura Hybrid Components Solar Panel Black": "Panel solarny czarny Ledvance Endura Hybrid",
            "Oprawa URBAN BOLLARD CONE CONCRETE ANCHORS LEDV": "Kotwy betonowe do słupka Ledvance Urban Bollard Cone",
            "Oprawa LINEAR EDGE 30CM RGBW DIM RC BK LEDV": "Oprawa 30cm RGBW z pilotem czarna Ledvance Linear Edge",
        }
        for old, new in cases.items():
            self.assertEqual(build_name(row(old), RULES, {})["nazwa"], new)

    def test_accessories_mount_kits(self):
        cases = {
            "Ramka o wysokości 75mm do montażu powierzchniowego PANEL 600 SURFACE MOUNT KIT H75":
                "Ramka natynkowa do paneli 60cm wys. 75 mm Ledvance",
            "Oprawa PANEL PL 600X600 ERECESSEDMOUNTFRAME EUE LEDV": "Ramka do montażu wpuszczanego do paneli 600x600 Ledvance",
            "Oprawa TRUSYS FLEX CHAIN KIT": "Zestaw łańcuchów do TruSys Flex Ledvance",
            "Oprawa Oprawa EMERGENCY KIT COMPACT BATTERY BAT LI 6.4 3.2AH": "Akumulator 6,4V 3,2Ah Li-Ion Ledvance",
            "Ręczna jednostka sterująca MANUAL CONTROL UNIT TOUCH DALI-2 TW TOUCH DALI-2 TW":
                "Ręczna jednostka sterująca TOUCH DALI-2 TW Ledvance",
        }
        for old, new in cases.items():
            self.assertEqual(build_name(row(old), RULES, {})["nazwa"], new)

    def test_bare_size_units(self):
        # bez jednostki: Ledvance "PLANON 120X30" = cm, GTV "SAKURA 120x85" = mm (2026-10-21)
        self.assertEqual(parse_deco_size("PLANON 120X30", "Panel LED", bare_cm=True), "120x30cm")
        self.assertEqual(parse_deco_size("SAKURA, ALUM., 120X85, IP20", "Oprawa sufitowa"), "12x8,5cm")

    def test_tubes_720_and_special_spectrum(self):
        res = build_name(row("Tuba LED T8 EM V 720 7W 830"), RULES, {})
        self.assertEqual(res["nazwa"], "Tuba LED T8 72cm 7W 3000K EM Ledvance")
        res = build_name(row("LEDTUBE T8 EM FOOD P 900 7.9W 833 LEDV"), RULES, {})
        self.assertEqual(res["nazwa"], "Tuba LED T8 90cm 7,9W EM do żywności Ledvance")
        self.assertNotIn("cct", res["braki"])


class LedvanceDecisions20261021Test(unittest.TestCase):
    """Decyzje uzytkownika po przegladzie v19 (reports/Ekonom-Ledvance_do_decyzji_2026-10-21.md)."""

    def test_candle_bw_is_plain_flame(self):
        # pkt 19: "BW" wyciete, zostaje "plomyk"
        name = build_name(row("Lampa LED CLASSIC BW P 3.4W 827 Clear E14"), RULES, {})["nazwa"]
        self.assertIn("płomyk przezroczysta", name)
        self.assertNotIn("BW", name)

    def test_dimming_by_package_name(self):
        # pkt 7: 3 STEP DIM = 3-stopniowa zamiast "sciemnialna"; Duo Click i GlowDim nazwa z opakowania
        self.assertEqual(parse_dim("FLAT LAMP 3 STEP DIM 3.5W 827"), "3-stopniowa")
        self.assertEqual(parse_dim("STAR+ CL A60 DIM DUO CLICK FILAMENT 7W"), "Duo Click")
        self.assertEqual(parse_dim("SUPERSTAR PAR16 GLOWDIM 4.5W"), "ściemnialna GlowDim")
        self.assertEqual(parse_dim("CLASSIC A DIM 5.9W"), "ściemnialna")
        name = build_name(row("Lampa LED FLAT LAMP 3 STEP DIM 3.5W 827 Frosted E14"), RULES, {})["nazwa"]
        self.assertIn("matowa 3-stopniowa", name)
        self.assertNotIn("ściemnialna", name)

    def test_tinted_glass_in_english(self):
        # pkt 6: Gold / Smoke / Amber po angielsku; barwione szklo zastepuje "przezroczysta", "matowa" zostaje
        self.assertEqual(parse_glass("VINTAGE 1906 CL EDISON SZKLO PRZEZROCZYSTE GOLD 22"), "Gold")
        self.assertEqual(parse_glass("VINTAGE 1906 GLOBE 4.5W 818 SMOKE E27"), "Smoke")
        self.assertEqual(parse_glass("NAV LED FILAMENT AMBER V 1600LM"), "Amber")
        self.assertEqual(parse_glass("CLASSIC GOLD FROSTED E27"), "matowa Gold")
        self.assertEqual(parse_glass("HALOSPOT MIRROR GOLD"), "lustrzana złota")

    def test_street_optics_as_code(self):
        # pkt 15: kod optyki ulicznej zawsze jako kod (RV20ST i RW20ST to rozne rozsyly), tez sklejony z barwa
        self.assertEqual(parse_optics("STREETLIGHT AREA 12 W 12W 727 RV20ST GY"), "opt. RV20ST")
        self.assertEqual(parse_optics("SLCOMPS6V025W 730RV40STNWAC2BTONEUELEDV"), "opt. RV40ST")
        self.assertEqual(parse_optics("STREETLIGHT AREA SPD MDV52W740RV20STGY4X1EUELEDV"), "opt. RV20ST")
        self.assertEqual(parse_optics("OPRAWA LED 7W 36ST"), "36°")  # Kobi bez zmian

    def test_print_sky_pattern_number_and_set(self):
        # pkt 13: numer wzoru przy nazwie wlasnej; SETn = komplet n paneli ("op. n szt.")
        ean = "4099854535413"
        rules = {**RULES, "_card_families": {ean: "Print Sky"}, "_card_family_raw": {ean: "PANEL PRINT SKY"}}
        res = build_name(row("Oprawa PANEL PRINT SKY1 600 V 33W LEDV", ean), rules, {})
        self.assertTrue(res["nazwa"].endswith("Ledvance Print Sky 1"), res["nazwa"])
        res = build_name(row("Oprawa PANEL PRINT SKY 600 V 33W SET9", ean), rules, {})
        res["row"], res["stara"] = row("Oprawa PANEL PRINT SKY 600 V 33W SET9", ean), "Oprawa PANEL PRINT SKY 600 V 33W SET9"
        self.assertTrue(res["nazwa"].endswith("Ledvance Print Sky"), res["nazwa"])
        apply_pack_size([res], RULES, {})
        self.assertTrue(res["nazwa"].endswith("Ledvance Print Sky op. 9 szt."), res["nazwa"])

    def test_decor_spot_points_and_remote_in_luminaires(self):
        # pkt 14: Decor Spot N-punktowy; pkt 9: "z pilotem" tez w oprawach parametrycznych
        name = build_name(row("Oprawa DECOR SPOT CORK BLACK 3x3.4W 927 DIM"), RULES, {})["nazwa"]
        self.assertTrue(name.startswith("Downlight LED 3-punktowy 3x3,4W"), name)
        name = build_name(row("Oprawa DECOR SPOT CORK BLACK 1x3.4W 927 DIM"), RULES, {})["nazwa"]
        self.assertNotIn("punktowy", name)
        name = build_name(row("Oprawa Office Line Double Wrap 0.6M 20W 840 DIM RC"), RULES, {})["nazwa"]
        self.assertIn("4000K z pilotem", name)

    def test_onoff_only_next_to_dali_twin(self):
        # pkt 11: ON-OFF zostaje tylko przy blizniaku DALI / 1DIM
        res = build_name(row("Oprawa DAMP PROOF ECO 3TH ON-OFF 33W 857 IP65 GRAY"), RULES, {})
        self.assertIn("ON-OFF", res["nazwa"])  # build_name zapisuje, decyduje `drop_lone_onoff`
        results = [{"nazwa": "Oprawa uliczna LED 25W 3000K opt. RV40ST ON-OFF Ledvance Compact", "uwagi": []},
                   {"nazwa": "Oprawa uliczna LED 25W 3000K opt. RV40ST 1DIM Ledvance Compact", "uwagi": []},
                   {"nazwa": "Oprawa hermetyczna LED 33W 5700K IP65 ON-OFF szara Ledvance Damp Proof Eco", "uwagi": []}]
        drop_lone_onoff(results)
        self.assertIn("ON-OFF", results[0]["nazwa"])
        self.assertEqual(results[2]["nazwa"], "Oprawa hermetyczna LED 33W 5700K IP65 szara Ledvance Damp Proof Eco")

    def test_hqi_daylight_kelvin_from_card(self):
        # pkt 18: HQI D / DL - K z karty; karta ponizej 5000 K przeczy D / DL -> WARNING
        ean = "4008321678348"
        card = {"exact": True, "cct": 5350, "url": "u", "name": "HQI-TS 70 W/D", "title": "HQI-TS 70 W/D"}
        res = build_name(row("Lampa wyładowcza HQI TS 70W D UVS", ean), {**RULES, "_gtin_cards": {ean: card}}, {})
        self.assertEqual(res["nazwa"], "Lampa metalohalogenkowa HQI-TS 70W 5350K Ledvance")
        self.assertEqual(res["status"], "OK")
        card = {**card, "cct": 4400, "name": "HQI-TS 1000 W/NDL/S UVS"}
        res = build_name(row("Lampa wyładowcza HQI TS 70W D UVS", ean), {**RULES, "_gtin_cards": {ean: card}}, {})
        self.assertEqual(res["status"], "WARNING")
        res = build_name(row("Lampa wyładowcza HQI TS 70W D UVS", ean), RULES, {})
        self.assertIn("barwa dzienna", res["nazwa"])  # bez karty bez zmian

    def test_gtin_card_of_other_product(self):
        # pkt 2: inny trzonek / ksztalt / dlugosc tuby = inny produkt; zapis trzonka "GX24Q.-3" to ten sam GX24Q-3
        def other(old, page):
            from build_ekonom_names import SOCKET_RE, parse_power
            o, p = parse_view(old), parse_view(page)
            return gtin_other_product(o, p, set(SOCKET_RE.findall(o)), set(SOCKET_RE.findall(p)),
                                      parse_power(o)[1], parse_power(p)[1], RULES)
        self.assertIn("trzonek", other("LED CLASSIC A DIM P 3.4W 827 Clear E27", "LED Classic P 40 DIM 3.4W 827 Frosted B22d"))
        self.assertEqual(other("DULUX LED T/E26 HF V 10W 830 GX24Q.-3LEDV", "DULUX LED T/E26 HF 10W 830 GX24Q-3"), "")
        self.assertIn("kształt", other("LED CLASSIC P V 25 2.5W/6500K E27", "LED Recycled Plastic STICK 100 11W 827 E27"))
        self.assertIn("długość", other("Tuba LED T8 EM PLUS E 600 8W 830", "LED TUBE T8 EM PLUS 1200 16W 830"))
        self.assertEqual(other("PARATHOM PRO DIM AR111 75 24° 1,7W 927 G53", "PPRO AR111 75 24° 11.7W 927 G53"), "")

    def test_camera_luminaire_from_card(self):
        # pkt 21: kamera z oswietleniem = rodzaj oprawy z kamera + W / lm / K z karty GTIN
        ean = "4058075564480"
        card = {"exact": True, "power": 15.0, "lumen": 950.0, "cct": 3000, "powers": [15.0], "lumens": [950.0],
                "url": "u", "name": "SMART+ Wifi Camera Round", "title": "SMART+ Wifi Camera Round", "family": ""}
        rules = {**RULES, "_gtin_cards": {ean: card}}
        res = build_name(row("Oprawa SMART+ OUTD WIFI ROUND CAMERA DG", ean), rules, {})
        self.assertEqual(res["nazwa"], "Kinkiet LED z kamerą Smart+ WiFi 15W 950lm 3000K grafitowy Ledvance Round")
        self.assertTrue(build_name(row("Oprawa SMART+ OUTD WIFI FLOOD CAMERA CON DG"), RULES, {})["nazwa"]
                        .startswith("Naświetlacz LED z kamerą Smart+"))
        self.assertEqual(build_name(row("Oprawa SMART OUTD WIFI GLOBE CAM TRACE E27"), RULES, {})["nazwa"],
                         "Kamera Smart+ WiFi do oprawy E27 Ledvance Globe Trace")

    def test_retrofit_is_a_distinguisher(self):
        # pkt 17: "Retrofit" odroznia blizniaki (wczesniej szum w `stop_patterns`)
        results = [built("Lampa LED Retrofit CLASSIC A DIM 5.9W 827 Clear E27", "4099854465055"),
                   built("Lampa LED CLASSIC A DIM 5.9W 827 Clear E27", "4099854463891")]
        self.assertEqual(results[0]["nazwa"], results[1]["nazwa"])
        add_all_distinguishers(results, RULES, {}, {})
        self.assertIn("Retrofit", results[0]["nazwa"])
        self.assertNotEqual(results[0]["nazwa"], results[1]["nazwa"])



class JoinedNameFormatTest(unittest.TestCase):
    """"Zestawienie" z Ekonoma (2026-10-26): "Nazwa towaru" = "<40 znakow> | <reszta>", bez kolumny "Nazwa dodatkowa"."""

    HEADER = "LP;Kod towaru;Nazwa towaru;J.M.;Ilość;Cena C3 - Cena Katalogowa;Cena Internet sklep;Kod producenta\r\n"
    EKONOM_HEADER = "LP;Kod towaru;Nazwa towaru;Nazwa dodatkowa;Kod producenta;JM;Cena netto;Waluta\r\n"

    def test_split_on_read_and_write_in_ekonom_layout(self):
        """Wynik idzie w ukladzie Ekonoma z "Nazwa dodatkowa", nie w ukladzie zestawienia (2026-10-27)."""
        with tempfile.TemporaryDirectory() as tmp:
            src = Path(tmp) / "in.csv"
            src.write_bytes((self.HEADER + '1;"4058075420991/LDV";"Naświetlacz FLOOD LED PERFORMANCE 20W 30 | 00K SYM";'
                             '"szt";0;71,82;;"4058075420991"\r\n2;"x";"LEDTUBE T5 10W 840 | ";"szt";;;;"1"\r\n')
                            .encode("cp1250"))
            fields, rows = read_csv(src)
            self.assertNotIn("Nazwa dodatkowa", fields)
            self.assertEqual(full_name(rows[0]), "Naświetlacz FLOOD LED PERFORMANCE 20W 3000K SYM")
            self.assertEqual(rows[1]["Nazwa dodatkowa"], "")
            name = "Naświetlacz LED 20W 2200lm 3000K 100° biały Ledvance Performance"
            rows[0]["Nazwa towaru"], rows[0]["Nazwa dodatkowa"] = split_for_ekonom(name)
            out = Path(tmp) / "out.csv"
            write_ekonom_csv(out, fields, rows)
            text = out.read_bytes().decode("cp1250")
            self.assertTrue(text.startswith(self.EKONOM_HEADER))
            self.assertNotIn(" | ", text)
            fields2, rows2 = read_csv(out)
            self.assertEqual(fields2, EKONOM_FIELDS)
            self.assertEqual(full_name(rows2[0]), name)
            self.assertLessEqual(len(rows2[0]["Nazwa towaru"]), 40)
            self.assertEqual((rows2[0]["JM"], rows2[0]["Cena netto"], rows2[0]["Waluta"]), ("szt", "71,82", "PLN"))
            self.assertEqual(rows2[1]["Cena netto"], "")


class LedvanceZestawienie20261026Test(unittest.TestCase):
    """Zestawienie Ledvance z 2026-10-26: stare kartoteki Osram z kodem na poczatku nazwy."""

    def test_order_reference_is_not_power(self):
        # kod zamowieniowy z karty GTIN "LVPR16501204,5W" (PAR16 50 120 4,5W) - wczesniej 16501204,5 W
        self.assertEqual(parse_power("LVPR16501204,5W"), ("", None))

    def test_short_lumen_codes(self):
        self.assertIn("1200lm", built("Naświetlacz FLOODLIGHT 10 8W 1K2LM 830 PS SY100 BK")["nazwa"])
        self.assertIn("10000lm", built("Naświetlacz FLOODLIGHT 100 69W 10KLM 830 PS SY100 BK")["nazwa"])

    def test_old_osram_codes(self):
        res = built("L15/76 NATURA ŚWIETLÓWKA OSRAM")
        self.assertEqual(body(res["nazwa"]), "Świetlówka 15W Natura")
        self.assertEqual(res["status"], "OK")
        self.assertEqual(body(built("DS7/41-827 G23 ŚWIETL. KOMP. OSRAM")["nazwa"]), "Świetlówka kompaktowa G23 7W 2700K")
        self.assertEqual(body(built("HCI T 70/930 WDL G12 LAMPA MET-HAL OSRAM")["nazwa"]),
                         "Lampa metalohalogenkowa HCI-T 70W G12 3000K")
        self.assertEqual(body(built("1613 LIN 35W S14D LINESTRA ŻARÓWKA LINIOWA OSRAM")["nazwa"]),
                         "Żarówka liniowa S14d 35W Linestra")
        self.assertEqual(body(built("64502 150W 230V GX6,35 ŻAR. HAL. OSRAM")["nazwa"]), "Żarówka halogenowa GX6.35 150W")
        self.assertTrue(built("ST8V-0.6M 7,6W/865 EM SUBSTITUBE VALUE OSRAM")["nazwa"].startswith("Tuba LED T8 60cm 7,6W"))

    def test_car_bulbs_are_osram(self):
        res = built("64210NL 55W 12V H7 ŻARÓWKA +150%SAMOCHODOWA")
        apply_package_brand([res], RULES, {})
        self.assertEqual(res["nazwa"], "Żarówka samochodowa H7 Night Breaker Laser +150% 55W 12V Osram")
        flashlight = built("Latarka samochodowa FLASHLIGHT CAR EMERGENCY")
        apply_package_brand([flashlight], RULES, {})
        self.assertTrue(flashlight["nazwa"].endswith("Ledvance"))

    def test_osram_with_ean_suffix(self):
        res = built("64432S HALOSTAR 35W 12V GY6,35 OSRAM 7426")
        apply_package_brand([res], RULES, {})
        self.assertEqual(res["nazwa"], "Żarówka halogenowa GY6.35 35W 12V Osram")

    def test_ballast_for_three_or_four_lamps(self):
        self.assertEqual(parse_ballast("STATECZNIK ELEKTRONICZNY QUICKTRONIC FIT8 3X18/4X18"), "3/4x18W")
        self.assertEqual(body(built("Statecznik elektroniczny Quicktronic FIT8 3x 4x18")["nazwa"]),
                         "Statecznik elektroniczny 3/4x18W")


SIGNIFY = {**load_yaml(DEFAULT_RULES)["suppliers"]["signify"], "params": True}


def signify_row(name, code="929003691502"):
    return {"Nazwa towaru": name[:40], "Nazwa dodatkowa": name[40:], "Kod towaru": "872016924864900/PHP",
            "Kod producenta": code}


def signify_built(name, card=None):
    from build_ekonom_names import signify_card_params
    feed = {"872016924864900": signify_card_params(card)} if card else {}
    return build_name(signify_row(name), SIGNIFY, {}, feed)


class SignifyTest(unittest.TestCase):
    """Signify (2026-11-03): karta lighting.philips.pl po 12NC jest zrodlem prawdy, jak pliki Parametry Kanlux."""

    def test_card_params_mapping(self):
        from build_ekonom_names import signify_card_params
        p = signify_card_params({"catalog_name": "DN142B 10S/830", "name": "CoreLine Downlight gen5, D150 mm, 1000 lm",
                                 "params": {"Strumień Świetlny": "2 900 lm 3 100 lm", "Zużycie energii": "31 W 22 W",
                                            "Kod stopnia ochrony": "IP20/44 [Ochrona]", "Całkowita średnica": "162 mm",
                                            "Zunifikowany wskaźnik ograniczenia olśnienia CEN": "19"}})
        self.assertEqual(p["Całkowity znamionowy strumień świetlny [lm]"], "2900/3100")
        self.assertEqual(p["Moc znamionowa [W]"], "22/31")
        self.assertEqual(p["Stopień IP"], "44")
        self.assertEqual(p["Średnica [mm]"], "150")  # rozmiar z nazwy karty, nie obudowa 162 mm
        self.assertEqual(p["UGR"], "<19")
        strip = signify_card_params({"catalog_name": "MAS LEDstrip 8.3W965 950LM/M 5M",
                                     "params": {"Strumień świetlny": "950 lm", "Napięcie wejściowe": "24 V"}})
        self.assertEqual(strip["Całkowity znamionowy strumień świetlny [lm]"], "1 m = 950")
        self.assertEqual(strip["Napięcie znamionowe [V]"], "24")

    def test_raw_catalog_lamp_from_card(self):
        card = {"catalog_name": "MAS LEDBulb D 4-60W E27 827 A60 CL G UE",
                "category": "Lampy i tuby LED > Żarówki LED > MASTER UltraEfficient LED bulb",
                "params": {"Podstawa-nasadka": "E27", "Strumień świetlny": "840 lm", "Zużycie energii": "4 W",
                           "Skorelowana temperatura barwowa (Nom)": "2700 K", "Kształt bańki": "A60",
                           "Z możliwością przyciemniania": "Tylko z niektórymi ściemniaczami"}}
        res = signify_built("MAS LEDBulb D 4-60W E27 827 A60 CL G UE", card)
        self.assertEqual(res["nazwa"], "Żarówka LED E27 4W 840lm 2700K A60 przezroczysta ściemnialna Philips MASTER")
        self.assertEqual(res["status"], "OK")

    def test_equivalent_power_is_not_power(self):
        # "LED classic 40W" = zamiennik 40 W; bez karty moc rzeczywista nieznana
        res = signify_built("Świeczki / kulki LED LED classic 40W P45 E14 WW CL ND RFSRT4")
        self.assertEqual(res["nazwa"], "Żarówka LED E14 2700K kulka przezroczysta zamiennik 40W Philips")
        self.assertEqual(res["status"], "DO_SPRAWDZENIA")

    def test_wrong_kind_in_ekonom_fixed_by_card_category(self):
        card = {"catalog_name": "TL-D Colored 58W Green 1SL/25",
                "category": "Konwencjonalne lampy i tuby > Świetlówki i zapłonniki do świetlówek > TL-D Colored",
                "params": {"Podstawa-nasadka": "G13", "Zużycie energii": "58,5 W"}}
        self.assertEqual(signify_built("Tuba LED 58W G13", card)["nazwa"], "Świetlówka T8 150cm 58W zielona Philips")

    def test_fluorescent_keeps_nominal_power(self):
        card = {"catalog_name": "MASTER TL5 HO 80W/865", "params": {"Zużycie energii": "79,8 W", "Podstawa-nasadka": "G5"}}
        res = signify_built("Świetlówka liniowe MASTER TL5 HO 80W/865 SLV/40", card)
        self.assertEqual(res["nazwa"], "Świetlówka T5 145cm 80W 6500K HO Philips MASTER")

    def test_fixture_family_code_and_line_last(self):
        card = {"catalog_name": "DN142B 10S/830 PSD-E UGR19",
                "category": "Oświetlenie wewnętrzne > Oprawy downlight > CoreLine Downlight gen5",
                "name": "Philips CoreLine Downlight gen5, 10.7 W, D150 mm, 1000 lm, 3000 K, DALI, UGR19",
                "params": {"Zużycie energii": "10,7 W", "Strumień Świetlny": "1 000 lm",
                           "Skorelowana temperatura barwowa (Nom)": "3000 K", "Kod stopnia ochrony": "IP20/44 [Ochrona]",
                           "Interfejs sterownika": "DALI", "Kolor Korpusu": "Biel"}}
        res = signify_built("Downlight LED 10.7W 1150lm 3000K IP20 DALI EL-ready UGR19 Philips DIM DN142B Coreline", card)
        # UGR19 z nazwy w Ekonomie (karta w tescie go nie podaje); "fi150" -> "150 mm" dopiero w main (`diameter_mm`)
        self.assertEqual(res["nazwa"], "Downlight LED fi150 10,7W 1000lm 3000K IP44 UGR19 DALI biały Philips DN142B CoreLine")
        self.assertIn("lumen z karty lighting.philips.pl: 1000lm (w Ekonomie: 1150lm)", res["uwagi"])

    def test_proper_name_and_street_optics(self):
        # 2026-11-03: nazwa wlasna rodziny z kategorii karty przed kodem, optyka i klasa z nazwy katalogowej karty
        card = {"catalog_name": "BRP102 LED55/740 II DM",
                "category": "Oprawy zewnętrzne > Oświetlenie drogowe i uliczne > CoreLine Malaga LED",
                "params": {"Zużycie energii": "39 W", "Strumień Świetlny": "4 600 lm",
                           "Skorelowana temperatura barwowa (Nom)": "4000 K", "Kod stopnia ochrony": "IP65 [Ochrona]"}}
        res = signify_built("Oprawa drogowa BRP102 LED55/740 II DM", card)
        self.assertEqual(res["nazwa"], "Oprawa uliczna LED 39W 4600lm 4000K IP65 opt. DM klasa II Philips Malaga BRP102 CoreLine")

    def test_floodlight_optics_short_near_end(self):
        # wytyczne 2026-11-04: "optyka symetryczna" -> "opt. sym.", blizej konca nazwy (za kolorem, przed marka)
        card = {"catalog_name": "BVP111 LED49-4S/730 S", "category": "Coreline Tempo Small Gen2",
                "params": {"Zużycie energii": "31 W", "Strumień Świetlny": "4 263 lm",
                           "Skorelowana temperatura barwowa (Nom)": "3000 K", "Kod stopnia ochrony": "IP66 [Ochrona]",
                           "Kolor Korpusu": "Szary"}}
        res = signify_built("Naświetlacz LED 31W 4263lm 3000K IP66 Philips BVP111 Coreline", card)
        self.assertEqual(res["nazwa"], "Naświetlacz LED 31W 4263lm 3000K IP66 szary opt. sym. Philips BVP111 CoreLine")
        card["catalog_name"] = "BVP110 LED49-4S/740 A"
        res = signify_built("Naświetlacz LED 31W 4263lm 3000K IP66 Philips BVP110 Coreline", card)
        self.assertIn("szary opt. asym. Philips", res["nazwa"])
        # High Bay: waska / szeroka optyka tez skrotem na koncu, przed marka
        res = signify_built("Highbay LED 100W 10000lm 4000K IP65 WB PILA BY007P")
        self.assertIn("IP65 opt. szer. PILA BY007P", res["nazwa"])

    def test_leaflet_pdf_parsing(self):
        # karta PDF docs.signify.com dla kodu, ktorego nie ma juz na stronie (angielska: separator tysiecy ",")
        sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))
        from fetch_signify_leaflets import parse_leaflet
        text = ("LL212X 53S/840 1x PSU WB 5x1.5 HE\n910925867495\nCoreLine Trunking gen2, 33 W, L1130 mm, 5300 lm,\n"
                "Light Technical\nLuminous Flux\n5,300 lm\nCorrelated Colour Temperature\n(Nom)\n4000 K\n"
                "Power Consumption\n33 W\nOverall length\n1130 mm\n")
        rec = parse_leaflet("910925867495", text, "https://x/FP/en_GB/910925867495_EU.en_GB.PROF.FP.pdf", "en_GB")
        self.assertEqual(rec["catalog_name"], "LL212X 53S/840 1x PSU WB 5x1.5 HE")
        self.assertEqual(rec["category"], "CoreLine Trunking gen2")
        self.assertEqual(rec["params"]["Strumień Świetlny"], "5300 lm")
        self.assertEqual(rec["params"]["Skorelowana temperatura barwowa (Nom)"], "4000 K")
        self.assertEqual(rec["params"]["Zużycie energii"], "33 W")


class KanluxZestawienie20261104Test(unittest.TestCase):
    """Zestawienie Kanlux "kart" z 2026-11-04 (1156 nowych kartotek) i karty kanlux.com."""

    def kx(self, name, code="1", rec=None):
        return build_name(kanlux_row(name, code), KANLUX, {}, {code: {**rec, "_name": name}} if rec else {})

    def test_new_breaker_codes(self):
        self.assertEqual(parse_breaker("KMBN6-C10/1"), "C10 1P+N 6kA")
        self.assertEqual(parse_breaker("KMCCB-F-125A-4P"), "125A 4P")
        self.assertEqual(parse_breaker("KMCCB-F-250A/RL/R"), "")  # wyzwalacz - akcesorium wylacznika
        self.assertEqual(parse_breaker("KMI-R-3/125A"), "125A 3P")
        self.assertEqual(parse_breaker("KCS-40A/1P"), "40A 1P")
        self.assertEqual(parse_breaker("KMC-20-11"), "20A 1NO+1NC")
        self.assertEqual(parse_breaker("KMC-20-04"), "20A 4NC")
        self.assertEqual(parse_breaker("KSPD-T1T2 275/100 1PN"), "T1+T2 1P+N")
        self.assertEqual(parse_breaker("KSPD-TT 275/160 3P+N"), "T2 TT 3P+N")
        self.assertEqual(parse_breaker("KSD-T2 275/40 M"), "T2 wkładka")
        self.assertEqual(parse_breaker("KRO6-1M-2/B10/30-A"), "B10 2P 30mA typ A 6kA")
        res = self.kx("KRO10-2/C16/30 Wyłącznik róznicowo-prądowy, z członem nadmiarowym, 10kA, 2P")
        self.assertEqual(res["nazwa"], "Wyłącznik RCBO C16 2P 30mA typ AC 10kA KRO10-2/C16/30 Kanlux")
        res = self.kx("KMI-3/63A Rozłącznik izolacyjny, modułowy")
        self.assertEqual(res["nazwa"], "Rozłącznik izolacyjny 63A 3P KMI-3/63A Kanlux")
        self.assertEqual(res["status"], "OK")

    def test_lumen_pairs_with_spaces(self):
        # "2500(WW) 2700(NW)" sklejalo sie w 25002700 - spacja tylko jako separator tysiecy
        rec = {"Strumień świetlny oprawy [lm]": "2500(WW) 2700(NW) / 4200(WW) 4400(NW)"}
        self.assertEqual(kx_lumen(rec, KANLUX["params_lumen_keys"]), "2500-4400lm")
        self.assertEqual(kx_lumen({"Strumień świetlny oprawy [lm]": "37 000"}, KANLUX["params_lumen_keys"]), "37000lm")

    def test_empty_card_key_is_not_brand(self):
        # karta kanlux.com ma pusty klucz z ikona klasy energetycznej ("" = "E") - marka "E" zamiast Kanlux
        rec = {"": "E", "Moc znamionowa [W]": "7", "Całkowity znamionowy strumień świetlny [lm]": "806",
               "Temperatura barwowa [K]": "2700", "Trzonek": "B22", "Źródło światła": "A60"}
        res = self.kx("XLEDIMW A60 B22 7W-WW Źródło światła LED", "37246", rec)
        self.assertTrue(res["nazwa"].endswith("Kanlux XLEDIMW"), res["nazwa"])
        with tempfile.TemporaryDirectory() as tmp:
            Path(tmp, "37246.json").write_text('{"code": "37246", "found": true, "name": "X", "params": {"": "E"}}',
                                               encoding="utf-8")
            params, _ = load_params_web(Path(tmp))
            self.assertNotIn("", params["37246"])

    def test_series_cleanup(self):
        self.assertIn("GRIP", self.kx("BRAVO S/P GRIP 12030 Uchwyt montażowy BRAVO")["nazwa"])
        self.assertIn("CTRL 12/24V", self.kx("CTRL 12/24V MONO/CCT Kontroler do taśm LED")["nazwa"])
        self.assertIn("1-10V", self.kx("BLINGO DRV40W1-10V BK Zasilacz LED")["nazwa"])
        self.assertTrue(self.kx("SPR.CTC-55/MR16 Akcesorium do opraw ścienno-sufitowych")["nazwa"]
                        .endswith("SPR.CTC-55/MR16"))
        self.assertNotIn("(", self.kx("AL-MS-WW-MAT-W-NT (35 Oprawa liniowa LED")["nazwa"])
        self.assertNotIn("11,8WCCT", self.kx("AZPO LED 11,8WCCT-O-W Plafoniera LED")["nazwa"])

    def test_dustproof_is_hermetic(self):
        # "Oprawa pyłoszczelna" z typu Kanlux = "Oprawa hermetyczna" (decyzja uzytkownika 2026-11-09);
        # zamiana w `name_replace` (stosowana w main), takze w odmianie
        def replaced(name):
            for item in KANLUX["name_replace"]:
                name = re.sub(item["match"], item["repl"], name)
            return name

        res = self.kx("DICHT 4LED PI 236/PS Oprawa pyłoszczelna pod tuby LED")
        self.assertTrue(res["nazwa"].startswith("Oprawa pyłoszczelna pod tuby"), res["nazwa"])
        self.assertTrue(replaced(res["nazwa"]).startswith("Oprawa hermetyczna pod tuby"))
        self.assertEqual(replaced("Akcesorium oprawy pyłoszczelnej Kanlux MAH PRO SENSOR C"),
                         "Akcesorium oprawy hermetycznej Kanlux MAH PRO SENSOR C")
        self.assertEqual(replaced("Pyłoszczelna oprawa"), "Hermetyczna oprawa")

    def test_glass_tube_with_type_tuba(self):
        res = self.kx("T8 LED GLASSv3 18W-NW Tuba LED 4000K 2700lm")
        self.assertEqual(res["nazwa"], "Tuba LED T8 120cm 18W 2700lm 4000K Kanlux GLASSv3")
        self.assertTrue(self.kx("T8 LED GLASSv3 24W-NW LAMPA LED")["nazwa"].startswith("Tuba LED T8 150cm 24W"))

    def test_split_exactly_at_40_before_preposition(self):
        with tempfile.TemporaryDirectory() as tmp:
            src = Path(tmp) / "in.csv"
            src.write_bytes(("LP;Kod towaru;Nazwa towaru;Kod producenta\r\n"
                             '1;"2161/KAN";"HLDR-E27-F     Oprawka do źródeł światła | z blaszką";"2161"\r\n'
                             '2;"28501/KAN";"BRAVO S/P GRIP 12030 Uchwyt montażowy BR | AVO";"28501"\r\n')
                            .encode("cp1250"))
            _, rows = read_csv(src)
            self.assertEqual(full_name(rows[0]), "HLDR-E27-F Oprawka do źródeł światła z blaszką")
            self.assertEqual(full_name(rows[1]), "BRAVO S/P GRIP 12030 Uchwyt montażowy BRAVO")


ELGOTECH = load_yaml(DEFAULT_RULES)["suppliers"]["elgotech"]


def elgotech_name(name, code):
    row = {"Nazwa towaru": name[:40], "Nazwa dodatkowa": name[40:], "Kod producenta": code, "Kod towaru": f"{code}/EGT"}
    return build_name(row, ELGOTECH, {})


class Elgotech20261110Test(unittest.TestCase):
    """Przedluzacze Elgotech wg schematu V3; kod producenta wycinany z nazwy, skroty z/u, b/u, 2P+Z."""

    def test_home_extension_lead_v3_order_and_line(self):
        res = elgotech_name("EPS-502W-2 Przedłużacz 5xz/u z wyłącz. CZARNY 2m", "EPS-502W-2")
        self.assertEqual(res["nazwa"], "Przedłużacz 5 gn. 2m z uziemieniem z wyłącznikiem czarny Economic Elgotech")
        res = elgotech_name("PSN-201,5W-1,5-2 Przedłużacz przenośny 2.-krotny, 2x2P.+Z czarny z podświetlanym "
                            "wyłącznikiem 250V~/16A 3x1,5mm2 PVC- 1,5m z zalewaną wtyczką kątową UNISCHUKO i otworem "
                            "do montażu, Classic PSN-201,5W-1,5-2", "PSN-201,5W-1,5-2")
        self.assertEqual(res["nazwa"],
                         "Przedłużacz 2 gn. 1,5m 3x1,5mm2 z uziemieniem z wyłącznikiem czarny Classic Elgotech")

    def test_bez_uziemienia_is_not_z_uziemieniem(self):
        res = elgotech_name("Przedłużacz PK-1010 bez uziemienia, przew.  2x1mm? 10m", "PK-1010")
        # katalog: PK-10 to przedluzacz ogrodowy z pomaranczowym przewodem
        self.assertEqual(res["nazwa"], "Przedłużacz ogrodowy 1 gn. 10m 2x1mm2 bez uziemienia pomarańczowy Elgotech")
        self.assertIn("bez uziemienia", elgotech_name("GN-4 Gniazdo cztekrotne bez uziemienia", "GN-4")["nazwa"])

    def test_drum_material_ip_and_small_from_code(self):
        res = elgotech_name("PZB-1111-40G/2,5 Przedł.bębn.4xIP-44 z/u H05RR 3x2,5 GUMA 40m", "PZB-1111-40G/2,5")
        self.assertEqual(res["nazwa"], "Przedłużacz bębnowy metalowy 4 gn. 40m 3x2,5mm2 guma IP44 z uziemieniem Elgotech")
        res = elgotech_name("PZB1-40-20G Przed.bęb. 4xz/u H05RR 3x1,5 (GUMA) 20m", "PZB1-40-20G")
        self.assertTrue(res["nazwa"].startswith("Przedłużacz bębnowy mały 4 gn. 20m"))
        res = elgotech_name("PSPZ1-12-20 Przedł.bęb.2xIP-44,1xsił. H07RN5x2,5GUMA 20m", "PSPZ1-12-20")
        self.assertEqual(res["nazwa"],
                         "Przedłużacz bębnowy 2 gn. + 1 gn. siłowe 20m 5x2,5mm2 guma H07RN-F IP44 Elgotech")

    def test_color_from_code_only_after_model_segment(self):
        self.assertEqual(elgotech_name("WT-20S-2 Wtyczka standard 16A z uziemieniem", "WT-20S-2")["nazwa"],
                         "Wtyczka standardowa z uziemieniem 16A czarna Elgotech")
        # "GN-2" = 2 gniazda, nie kolor
        self.assertEqual(elgotech_name("GN-2 Gniazdo dwukrotne bez uziemienia", "GN-2")["nazwa"],
                         "Gniazdo przenośne 2 gn. bez uziemienia Elgotech")

    def test_ip_at_end_survives_code_cleanup(self):
        self.assertEqual(elgotech_name("WT-30G Wtyczka unischuko z uziemieniem IP-44", "WT-30G")["nazwa"],
                         "Wtyczka Unischuko z uziemieniem IP44 Elgotech")

    def test_code_length_conflict_goes_to_review(self):
        res = elgotech_name("PSN-407-1,5 Przedłużacz przenośny 4-krotny, 4x2P+Z biały 250V~/16A 3x1,5mm2 PVC- 6m z "
                            "zalewaną wtyczką kątową UNISCHUKO i otworami do montażu, Classic PSN-406-1,5", "PSN-407-1,5")
        self.assertEqual(res["status"], "DO_SPRAWDZENIA")
        self.assertIn("6m", res["nazwa"])

    def test_catalog_fills_cross_and_color_and_flags_conflict(self):
        rules = {**ELGOTECH, "_catalog": {"PSN-303": {"rest": "3x1 mm2 biały"},
                                          "PSN-503W-2": {"rest": "3x1 mm2 czarny"},
                                          "R-20BS-24": {"rest": "biały"}}}
        row = {"Nazwa towaru": "PSN-303 Przedłużacz 3xz/u 3m", "Nazwa dodatkowa": "", "Kod producenta": "PSN-303",
               "Kod towaru": "PSN-303/EGT"}
        self.assertEqual(build_name(row, rules, {})["nazwa"],
                         "Przedłużacz 3 gn. 3m 3x1mm2 z uziemieniem biały Classic Elgotech")
        row = {"Nazwa towaru": "PSN-503W-2 Przedłużacz 5xz/u biały z wyłącznikiem 3m", "Nazwa dodatkowa": "",
               "Kod producenta": "PSN-503W-2", "Kod towaru": "PSN-503W-2/EGT"}
        self.assertEqual(build_name(row, rules, {})["status"], "DO_SPRAWDZENIA")
        # blad tabeli katalogu (zdjecie: czarno-czerwony) - bez konfliktu
        row = {"Nazwa towaru": "R-20BS-24 ROZGAŁĘŹNIK 2X Z/U IP44 CZARNO-CZERWONY 16A", "Nazwa dodatkowa": "",
               "Kod producenta": "R-20BS-24", "Kod towaru": "R-20BS-24/EGT"}
        self.assertEqual(build_name(row, rules, {})["status"], "OK")

    def test_surge_model_garden_and_splitter_shape(self):
        self.assertTrue(elgotech_name("PSF3-501 Przed. przeciwprzepięciowy 5xz/u i wył., szary 1,5m", "PSF3-501")
                        ["nazwa"].endswith("szary PSF3 Elgotech"))
        self.assertEqual(elgotech_name("Przedłużacz PS-1010-5 z uziemieniem, przew. 3x1mm? 10m", "PS-1010-5")["nazwa"],
                         "Przedłużacz ogrodowy 1 gn. 10m 3x1mm2 z uziemieniem pomarańczowy Elgotech")
        self.assertTrue(elgotech_name("RT-30CA Rozgałęźnik 3GN. z uziemieniem", "RT-30CA")["nazwa"]
                        .startswith("Rozgałęźnik trójkątny 3 gn."))
        self.assertTrue(elgotech_name("PZB-1111-40G Przedł.bębn.4xIP-44 z/u H05RR 3x1,5 GUMA 40m", "PZB-1111-40G")
                        ["nazwa"].startswith("Przedłużacz bębnowy metalowy 4 gn. 40m"))

    def test_catalog_extra_before_bare_length_and_small_coiled(self):
        rules = {**ELGOTECH, "_catalog": {"PSN-501,5": {"rest": "3x1 mm2 biały"},
                                          "PWP-113-1,5": {"rest": "5x1,5 mm2 H07RN-F"}}}
        row = {"Nazwa towaru": "PSN-501,5 Przedłużacz 5xz/u 1,5", "Nazwa dodatkowa": "", "Kod producenta": "PSN-501,5",
               "Kod towaru": "PSN-501,5/EGT"}
        self.assertEqual(build_name(row, rules, {})["nazwa"],
                         "Przedłużacz 5 gn. 1,5m 3x1mm2 z uziemieniem biały Classic Elgotech")
        row = {"Nazwa towaru": "PWP-113-1,5 Przedłużacz 2xgn.jednofaz.,1", "Nazwa dodatkowa": "xgn.siłowe 1,5m",
               "Kod producenta": "PWP-113-1,5", "Kod towaru": "PWP-113-1,5/EGT"}
        self.assertEqual(build_name(row, rules, {})["nazwa"],
                         "Przedłużacz 2 gn. + 1 gn. siłowe 1,5m 5x1,5mm2 guma H07RN-F Elgotech")
        self.assertTrue(elgotech_name("PSZ-306 Przedłużacz zwijany 3xz/u, MAŁY 6m", "PSZ-306")["nazwa"]
                        .startswith("Przedłużacz zwijany mały 3 gn. 6m"))
        self.assertEqual(elgotech_name("G-10G Gniazdo z uziemieniem IP-44", "G-10G")["nazwa"],
                         "Gniazdo przenośne 1 gn. z uziemieniem IP44 Elgotech")
        # "gniazda okrągłe" to rodzaj gniazd, nie obudowa
        self.assertFalse(elgotech_name("R-30-2 Rozgałęźnik 3-krotny 3x2P+Z gniazda okrągłe 250V~/16A czarny R-30-2",
                                       "R-30-2")["nazwa"].startswith("Rozgałęźnik okrągły"))


PLASTROL = load_yaml(DEFAULT_RULES)["suppliers"]["plastrol"]


def plastrol_name(name, code):
    row = {"Nazwa towaru": name[:40], "Nazwa dodatkowa": name[40:], "Kod producenta": code, "Kod towaru": f"{code}/PLST"}
    return build_name(row, PLASTROL, {})["nazwa"]


class Plastrol20261110Test(unittest.TestCase):
    """Plastrol: przedluzacze wg schematu V3, linie bebnow, X1 / X3 = budowlane 1 / 3 gn., KEL = marka Plastrolu."""

    def test_drum_line_and_cable(self):
        self.assertEqual(plastrol_name("Przedłużacz bębnowy HEAVY LINE: 4x2p+Z: IP44: 16A: H07BQ-F: 3x2,5: 40m: "
                                       "NIEBIESKI", "W-99275"),
                         "Przedłużacz bębnowy 4 gn. 40m 3x2,5mm2 PUR IP44 z uziemieniem Heavy Line Plastrol")
        self.assertEqual(plastrol_name("Przedłużacz bębnowy z automatycznym zwijaczem 14+2m IP20, 3×1,5mm",
                                       "W-100518"),
                         "Przedłużacz bębnowy automatyczny 14+2m 3x1,5mm2 IP20 Plastrol")

    def test_construction_x3_colour_of_plug_not_cable(self):
        self.assertEqual(plastrol_name("Przedłużacz  X3, przewód w oponie gumowej H05RR-F czarna 3x1,5mm2, IP 54, "
                                       "wtyczka i gniazdo w kolorze czarno-czerwonym, przewód kolor czarny", "W-98985"),
                         "Przedłużacz budowlany 3 gn. 3x1,5mm2 guma IP54 czarno-czerwony X3 Plastrol")

    def test_two_pole_is_without_earth(self):
        self.assertEqual(plastrol_name("Przedłużacz jednogniazdowy ECO 2p, 10A,  2500W, z przewodem OMY 2x1mm2, "
                                       "10m POMARAŃCZOWY", "W-98937"),
                         "Przedłużacz 1 gn. 10m 2x1mm2 bez uziemienia pomarańczowy Eco Plastrol")

    def test_kel_brand_and_usb(self):
        self.assertEqual(plastrol_name("Przedłużacz stały z sześciokrotnym gniazdem wtyczkowym 2p+Z oraz gniazdami "
                                       "USB A, USB C, 2m, SCHUKO  KEL", "W-00354"),
                         "Przedłużacz 6 gn. + USB A+C 2m Schuko KEL")

    def test_plug_model_after_lost_space(self):
        self.assertEqual(plastrol_name("Wtyczka rozbieralna kątowa płaska WT-11PBIAŁA", "W-99969"),
                         "Wtyczka rozbieralna kątowa płaska biała WT-11P Plastrol")


PCE = load_yaml(DEFAULT_RULES)["suppliers"]["pce"]


def pce_name(name, code):
    row = {"Nazwa towaru": name[:40], "Nazwa dodatkowa": name[40:], "Kod producenta": code, "Kod towaru": f"{code}/PCE"}
    return build_name(row, PCE, {})


class Pce20261110Test(unittest.TestCase):
    """PCE: zlacza CEE (bieguny 3P / 4P / 5P, seria), rozdzielnice (wyjscia, zabezpieczenia, nazwa), bebny wg V3."""

    def test_cee_poles_and_series(self):
        self.assertEqual(pce_name("WTYCZKA PRZENOŚNA SHARK TT IP44 16A 400V 3P+Z+N BEZŚRUBOWA", "015-6TT")["nazwa"],
                         "Wtyczka przenośna 16A 5P 400V IP44 Shark TT PCE")
        # ostatnie IP wygrywa ("... IP44 16A 24V 2P IP67"), 24V/~ = AC
        self.assertEqual(pce_name("WTYCZKI PRZENOŚNE TWIST 24V  IP44 16 A 24V/~50/60Hz 2P IP67", "0622v")["nazwa"],
                         "Wtyczka przenośna 16A 2P 24V AC IP67 Twist PCE")
        # "3P+Z" + ".+N" na cieciu i "400V" + "3P+Z" bez spacji
        self.assertEqual(pce_name("WTYCZKA ODBIORNIKOWA IP67 125A 400V 3P+Z.+N ", "545-6")["nazwa"],
                         "Wtyczka odbiornikowa 125A 5P 400V IP67 PCE")
        self.assertEqual(pce_name("GNIAZDO STAŁE bez dławika IP67 125A 400V3P+Z ", "144-6")["nazwa"],
                         "Gniazdo stałe 125A 4P 400V IP67 bez dławika PCE")

    def test_board_outputs_protection_and_name(self):
        self.assertEqual(pce_name("ROZDZIELNICA STACJONARNA Z TWORZYWA TYP „ADAMOW” 10M IP44 Wejście M25 - wyjście "
                                  "32/5xB32 2xGSx2xB16 FI40/4/0,03", "9036314W")["nazwa"],
                         "Rozdzielnica stacjonarna 32/5 2x230V z RCD 40A IP44 Adamów PCE")
        # "16/5 GS" = jedno gniazdo 230V, nie piec
        self.assertEqual(pce_name("ROZDZIELNICA STACJONARNA Z TWORZYWA TYP „KOŁO”  IP44 Wejście M20 - wyjście 16/5 GS "
                                  "R.0-1 16A 3P ", "9769215W")["nazwa"],
                         "Rozdzielnica stacjonarna 16/5 1x230V z rozłącznikiem 0-1 IP54 Koło PCE")  # IP54 z nazwy producenta
        # zasilanie wtyczka 32/5 nie jest wyjsciem; zapis "2*16*B" = 2x16/5 z wylacznikami
        self.assertEqual(pce_name("ROZDZIELNICA PRZENOŚNA Z TWORZYWA TYP „PIESZYCE”  IP44 Wejście 32/5+1,5M - wyjście "
                                  "3x32/5", "9430009Wk")["nazwa"],
                         "Rozdzielnica przenośna 3x32/5 z wtyczką 32/5 IP44 Pieszyce PCE")
        self.assertEqual(pce_name("ROZDZIELNICA STACJONARNA Z ZABEZPIECZENIAMI IP54 OPOLE 32/5*B 2*16*B 5GS*4B FI63, "
                                  "IP54", "909110004W")["nazwa"],
                         "Rozdzielnica stacjonarna 32/5 2x16/5 5x230V z RCD 63A IP54 Opole PCE")

    def test_drums_v3(self):
        self.assertEqual(pce_name("PRZEDŁUŻACZ ZWIJANY JEDNOFAZOWY TYP „XREEL” - WYJŚCIE 4xGS Z TER. IP44 Wejście "
                                  "20 mb H05RR-F 3x1,5(OW) wtyczka gumowa 16A IP54 ", "92501H48123")["nazwa"],
                         "Przedłużacz bębnowy 4 gn. 20m 3x1,5mm2 guma IP44 z uziemieniem z termikiem XREEL PCE")
        self.assertEqual(pce_name("PRZEDŁUŻACZ ZWIJANY TRÓJFAZOWE TYP „KŁODZKO”  IP44 Wejście 30 mb H05RR-F 5x1,5 "
                                  "(OW) wtyczka 16/5 – wyjście 2*16/5 GS TERMIK IP44", "9331G32265")["nazwa"],
                         "Przedłużacz bębnowy trójfazowy 2x16/5 1x230V 30m 5x1,5mm2 guma IP44 z termikiem Kłodzko PCE")

    def test_single_phase_earth_types(self):
        self.assertEqual(pce_name("WTYCZKA JEDNOFAZOWA GUMOWA 10/16A 2P+Z IP44 Wtyczka z bolcem \"uni-schuko\" pełna "
                                  "guma ", "0522-sr")["nazwa"],
                         "Wtyczka pełnogumowa Unischuko 16A IP44 PCE")
        self.assertEqual(pce_name("GNIAZDO PRZENOŚNE JEDNOFAZOWE GUMOWE 10/16A 2 P+Z TOP TAURUS 2 IP44 Gniazdo gumowe "
                                  "z zab. bocznym VDE  (z klapką)", "25712-sr")["nazwa"],
                         "Gniazdo przenośne gumowe Schuko 16A IP54 z zatyczką Taurus2 PCE")
        # producent: 25712 = z zatyczka i IP54 (Ekonom: z klapka, IP44)

    def test_ekonom_errors_fixed_from_manufacturer_with_note(self):
        # wtyczka tablicowa 2P+Z 6h = 230V (karta producenta), Ekonom pisze 400V
        res = pce_name("WTYCZKA TABLICOWA  IP44 16A 400V 2P+Z", "613-6")
        self.assertEqual(res["nazwa"], "Wtyczka tablicowa 16A 3P 230V IP44 PCE")
        self.assertEqual(res["status"], "WARNING")
        self.assertTrue(any("w Ekonomie 400V" in u for u in res["uwagi"]))
        # 75132-6 = IP67 (pcelectric.at); bez poprawki mial nazwe identyczna z 7513-6
        res = pce_name("GNIAZDO Z WYŁĄCZNIKIEM 0-1 DUŻE BLOKADA MECHANICZNA IP44 16A 230V 2P+Z IP44", "75132-6")
        self.assertEqual(res["nazwa"], "Gniazdo z wyłącznikiem 0-1 i blokadą 16A 3P 230V IP67 duże PCE")
        # puszka -171 opisana w Ekonomie wymiarami -161
        res = pce_name("PUSZKA SERII „LIGHT”  PUSZKA INSTALACYJNA LIGHT POKRYWA PEŁNA, DŁAWNICE 162x116x76 "
                       "(6xM25, 4xM40) IP66/67 ", "EV-171")
        self.assertEqual(res["nazwa"], "Puszka instalacyjna 155x179x100mm z dławnicami IP66/67 Light PCE")
        res = pce_name("GNIAZDO Z WYŁĄCZNIKIEM 0-1 DUŻE BLOKADA MECHANICZNA IP44 16A 400V 3P+Z IP67", "75142-6")
        self.assertEqual(res["status"], "OK")


PAWBOL = load_yaml(DEFAULT_RULES)["suppliers"]["pawbol"]


def pawbol_name(name, code, maker=None):
    """Nazwa z Ekonoma + opis producenta (cennik Pawbol 2026); `name_replace` jak w main()."""
    row = {"Nazwa towaru": name[:40], "Nazwa dodatkowa": name[40:], "Kod producenta": code,
           "Kod towaru": f"{code}/PAW", "_feed_key": pawbol_code_key(code)}
    feed = {row["_feed_key"]: {"_name": maker, "_source": "cennik Pawbol 2026"}} if maker else {}
    res = build_name(row, {**PAWBOL, "_withdrawn": set()}, {}, feed)
    for item in PAWBOL.get("name_replace", []):
        res["nazwa"] = re.sub(item["match"], item["repl"], res["nazwa"])
    return res


class Pawbol20261110Test(unittest.TestCase):
    """Pawbol: nazwa z cennika producenta, rozdzielnice jak PCE, bebny wg V3, BETA z kodu, bledy cennika z kodu."""

    def test_code_key_and_loader(self):
        self.assertEqual(pawbol_code_key("H.7015B -25"), "H.7015B-25")
        with tempfile.TemporaryDirectory() as tmp:
            Path(tmp, "cennik_2026.json").write_text(
                '{"A.0001": {"name": "Puszka podtynkowa PK60, płytka 40mm"}, "B.1": {"name": "Produkt wycofany"},'
                ' "B.2": {"name": "Produkt wycofany"}}', encoding="utf-8")
            Path(tmp, "produkty.json").write_text(
                '{"B.2": {"name": "Rozdzielnica R-BOX 150: 2x250V", "groups": ["R-BOX 150"]}}', encoding="utf-8")
            names, withdrawn = load_pawbol_names(Path(tmp))
        self.assertEqual(names["A.0001"]["_source"], "cennik Pawbol 2026")
        self.assertEqual(names["B.2"]["_name"], "Rozdzielnica R-BOX 150: 2x250V")  # wycofany w cenniku, jest na stronie
        self.assertEqual(withdrawn, {"B.1"})

    def test_board_from_price_list_and_from_ekonom(self):
        res = pawbol_name("R-BOX 240R 4S, 1x32A/5p, 1x16A/5p, 2x250V, puste okno", "B.1701R",
                          "Rozdzielnica R-BOX 240 4S: 1x32A/5P, 1x16A/5P, 2x250V, 400V, IP44")
        self.assertEqual(res["nazwa"], "Rozdzielnica 32/5 16/5 2x230V z rączką IP44 R-BOX 240 4S Pawbol")
        self.assertIn("nazwa z: cennik Pawbol 2026", res["uwagi"])
        # kartoteka spoza cennika: zapis Ekonoma "2x32/5", moduly za wyjsciami, RCD 4/40/0,03
        res = pawbol_name("R-BOX 380R 1x16/5, 2x32/5, 4x230V 13S zabezp. 2xM.01-B32/3, 1xM.01-B16/3, "
                          "1xM.02-4/40/0,03 IP44", "B.18.380-2ZRP")
        self.assertEqual(res["nazwa"],
                         "Rozdzielnica 2x32/5 16/5 4x230V z wył. nadprądowymi i RCD 40A IP44 R-BOX 380 13S Pawbol")
        # przelacznik "1xL/P 40A" i 2 podstawy (-PS) / 1 podstawa (-PS_1)
        res = pawbol_name("R-BOX 240, 1x32A/5p, 1x16A/5p, 2x250V, wył. L/P", "B.1095W-PS",
                          "Rozdzielnica R-BOX 240: 1x32A/5P, 1x16A/5P, 2x250V, 1xL/P 40A, 400V, IP44, 1x termik "
                          "+ 2x podstawka")
        self.assertEqual(res["nazwa"], "Rozdzielnica 32/5 16/5 2x230V z przełącznikiem L-0-P z termikiem "
                                       "z 2 podstawami IP44 R-BOX 240 Pawbol")

    def test_drums_v3_single_and_three_phase(self):
        self.assertEqual(pawbol_name("Przedłużacz PB H05RR F 3x1,5 - 25m, 4xGZ16A (250V), IP 44", "B.1116-300G-25",
                                     "Przedłużacz na bębnie plastikowym, fi300, przewód H05RR-F 3x1,5mm2 - 25m, "
                                     "4x250V, 250V, 1x termik, czarny, IP44")["nazwa"],
                         "Przedłużacz bębnowy fi300 4 gn. 25m 3x1,5mm2 guma IP44 z uziemieniem z termikiem Pawbol")
        self.assertEqual(pawbol_name("Przedłużacz PBS H05RR F 5x1,5 - 25m", "B.111015-25",
                                     "Przedłużacz na bębnie plastikowym, fi340, przewód H05RR-F 5x1,5mm2 - 25m, "
                                     "1x16A/5P, 2x250V, 400V, 1x termik, biały, IP44")["nazwa"],
                         "Przedłużacz bębnowy trójfazowy fi340 16/5 2x230V 25m 5x1,5mm2 guma IP44 z uziemieniem "
                         "z termikiem Pawbol")

    def test_beta_layout_from_code_overrides_price_list(self):
        # cennik: "gniazdo BETA pojedyncze", kod GG i Ekonom: podwojne
        res = pawbol_name("Wyłącznik BETA1, pojedynczy, jednobiegunowy z gniazdem BETA1, podwójnym", "D.3114BGGW/1",
                          "Wyłącznik BETA pojedynczy, jednobiegunowy i gniazdo BETA pojedyncze, białe, biała klapka, "
                          "IP44")
        self.assertEqual(res["nazwa"], "Wyłącznik jednobiegunowy z gniazdem podwójnym biały klapka biała IP44 "
                                       "Beta Pawbol")
        self.assertEqual(res["status"], "WARNING")
        self.assertEqual(pawbol_name("Gniazdo herm.BETA1,3x2P+Z,IP44,czarne ,schuko", "D.3715SCGGG/1",
                                     "Gniazdo BETA potrójne schuko, czarne, czarna klapka, IP44")["nazwa"],
                         "Gniazdo potrójne Schuko czarne klapka czarna IP44 Beta Pawbol")
        # stara kartoteka bez koloru obudowy ("dymnym, ob. brąz") - kolor z kodu D.32..
        res = pawbol_name("Wyłącznik BETA1, podwójny, jednobiegunowy+świecznikowy z gniazdem BETA1, 2P+Z, IP44, "
                          "dymnym, ob. brąz*", "D.3214GWW2/1")
        self.assertEqual(res["nazwa"], "Wyłącznik podwójny jednobiegunowy + świecznikowy z gniazdem brązowy "
                                       "klapka dymna IP44 Beta Pawbol")
        self.assertTrue(any("brak w cenniku" in u for u in res["uwagi"]))

    def test_luminaires(self):
        self.assertEqual(pawbol_name("Oprawa PANTERA LED", "D.3198SRM-BN-13W",
                                     "Oprawa PANTERA, LED, 13W, klosz mleczny, pierścień srebrny, z czujnikiem ruchu, "
                                     "odporna na UV, 230V, IK10, IP54")["nazwa"],
                         "Plafon LED 13W 4000K klosz mleczny srebrny z czujnikiem ruchu IP54 Pantera Pawbol")
        # E27: moc to limit zarowki - bez W; OWAL 60 / 100 to rozmiar
        self.assertEqual(pawbol_name("Oprawa kanałowa z siatką metalową owalna 60W, IP 44", "D.3166",
                                     "Oprawa OWAL, E27, 60W, klosz przeźroczysty szklany, podstawa biała, siatka "
                                     "metalowa, 230V, IK06, IP44")["nazwa"],
                         "Oprawa kanałowa E27 klosz przezroczysty szklany z siatką metalową biała IP44 Owal 60 Pawbol")
        # wycofana BZ (bez cennika): barwa slowem, K nieznane
        self.assertEqual(pawbol_name("Oprawa oświetleniowa IRYS LED czarny pierścień, klosz mleczny b.zimny  13W",
                                     "D.31101CM-BZ-13W")["nazwa"],
                         "Plafon LED 13W barwa zimna klosz mleczny czarny Irys Pawbol")

    def test_boxes_and_price_list_errors(self):
        self.assertEqual(pawbol_name("Puszka do pustych ścian ProDuo podwójna", "A.0092K",
                                     "Puszka do pustych ścian ProDuo podwójna, głęboka 60mm, fi68, łączeniowa, z "
                                     "wkrętami, wiatroszczelna, pomarańczowa, IP30")["nazwa"],
                         "Puszka do pustych ścian głęboka 60mm fi68 łączeniowa podwójna z wkrętami wiatroszczelna "
                         "pomarańczowa ProDuo Pawbol")
        self.assertEqual(pawbol_name("Puszka hermetyczna 120x80x50 dł.d.przeźr.IP65 bezhalogenowa", "S-BOX 206-P",
                                     "Puszka hermetyczna S-BOX 120x80x50, 6 dławików, szara, z pokrywą "
                                     "przeźroczystą, IP65")["nazwa"],
                         "Puszka hermetyczna 120x80x50mm 6 dławików z pokrywą przezroczystą szara IP65 S-BOX Pawbol")
        # cennik: LZ 10/Ż "zielona", B.1923 "wtyczka tablicowa" - poprawione z kodu / strony, z uwaga
        res = pawbol_name("Listwa zaciskowa niska podstawa 10 - dzielona 2x5 żółta", "E.4132",
                          "Listwa zaciskowa ochronna dzielona LZ 10/Ż, niska podstawa, 10x16mm2, 32A, 400V, zielona, "
                          "IP20")
        self.assertIn("żółta", res["nazwa"])
        self.assertTrue(any("cennik Pawbol 2026: zielona" in u for u in res["uwagi"]))
        self.assertEqual(pawbol_name("Wtyczka siłowa 32A 3P IP67", "B.1923",
                                     "Wtyczka tablicowa prosta, 32A/3P, 250V, niebieska, IP67")["nazwa"],
                         "Wtyczka przenośna 32A 3P 230V IP67 Pawbol")
        # cecha tylko w nazwie z Ekonoma (cennik: oba maszty "2m")
        self.assertEqual(pawbol_name("Maszt odgromowy wolnostojący h=2m fi16/10 podst. 15kg", "R.8312",
                                     "Maszt odgromowy, wolnostojący, 2m")["nazwa"],
                         "Maszt odgromowy wolnostojący 2m podstawa 15kg Pawbol")


FELEKTRO = load_yaml(DEFAULT_RULES)["suppliers"]["felektro"]


def felektro_name(name, code, card=None):
    """Nazwa z Ekonoma + rekord karty Onninen (`load_onninen_cards`); marka i `name_replace` jak w main()."""
    row = {"Nazwa towaru": name[:40], "Nazwa dodatkowa": name[40:], "Kod producenta": code, "Kod towaru": f"{code}/FEL"}
    feed = {code: {"_url": "https://onninen.pl/produkt/x", "_slug": "", **card}} if card else {}
    res = build_name(row, FELEKTRO, {}, feed)
    move_brand_tail([res], FELEKTRO)
    for item in FELEKTRO.get("name_replace", []):
        res["nazwa"] = re.sub(item["match"], item["repl"], res["nazwa"])
    return res


class FElektro20261110Test(unittest.TestCase):
    """F-Elektro: rozdzielnice budowlane (wyjscia, zabezpieczenia, przelacznik, zasilanie, seria z modulami),
    przedluzacze bebnowe wg V3, IP i przelacznik z karty Onninen, opis z Ekonoma dla osprzetu."""

    def test_boards(self):
        self.assertEqual(felektro_name("Rozdzielnica R-240 - gniazda 2x32A 5p, 2x230V, L/P (32A), wtyczka 32A 5p, "
                                       "OPD 5x4mm2 -1,5m", "F3.0147")["nazwa"],
                         "Rozdzielnica 2x32/5 2x230V z przełącznikiem L-0-P 32A z wtyczką 32/5 na przewodzie 5x4mm2 1,5m "
                         "R-240 F-Elektro")
        self.assertEqual(felektro_name("Rozdzielnica LAGO 6M - gniazda 2x230V z zabezpieczeniami R2P25A typ AC, "
                                       "1FB16Ax2 /FE/", "F3.3598")["nazwa"],
                         "Rozdzielnica 2x230V z wył. nadprądowymi i RCD 25A LAGO 6M F-Elektro")
        # wtyk odbiornikowy w obudowie to zasilanie, nie wyjscie; "odbornikowy" = literowka Ekonoma
        self.assertEqual(felektro_name("Rozdzielnica POLYBOX - wtyk odbornikowy 32A 5p, gniazda 16A 5p, 2x32A 5p",
                                       "F3.0616")["nazwa"],
                         "Rozdzielnica 16/5 2x32/5 z wtykiem 32/5 POLYBOX F-Elektro")
        # "-Z" = przelacznik zamykany; "2x250V" = gniazda 230V
        self.assertEqual(felektro_name("Rozdzielnica RS-KOMBI - gniazda 32A 5p, 2x230V, L/P (32A-Z)", "F3.1890")["nazwa"],
                         "Rozdzielnica 32/5 2x230V z przełącznikiem zamykanym L-0-P 32A RS-KOMBI F-Elektro")
        self.assertEqual(felektro_name("Rozdzielnica ULISSE 6M MAXI - gniazda 16A 5p, 32A 5p, 2x250V", "74180.8")["nazwa"],
                         "Rozdzielnica 16/5 32/5 2x230V ULISSE 6M MAXI F-Elektro")

    def test_card_ip_and_switch(self):
        # wersja IP65 ma w Ekonomie te sama nazwe bez " - gniazda" - IP z tytulu karty Onninen, uwaga ze zrodlem
        res = felektro_name("Rozdzielnica LAGO 10M 32A 5p, 2x230V", "F3.0322",
                            {"IP": "IP65"})
        self.assertEqual(res["nazwa"], "Rozdzielnica 32/5 2x230V IP65 LAGO 10M F-Elektro")
        self.assertEqual(res["status"], "OK")
        self.assertIn("IP65 z tytułu karty Onninen", res["uwagi"][0])
        # RB-300: Ekonom zgubil przelacznik, tytul karty go ma
        self.assertEqual(felektro_name("Rozdzielnica budowlana RB-300 12M 2x16A5p, 2x230V", "F3.0544",
                                       {"_slug": "F-ELEKTRO-ROZDZIELNICA-BUDOWLANA-RB-300-12M-2X16A5P-2X230V-L-P-32A-F3-0544"}
                                       )["nazwa"],
                         "Rozdzielnica budowlana 2x16/5 2x230V z przełącznikiem L-0-P 32A RB-300 12M F-Elektro")
        # IP tylko z generowanego opisu karty - WARNING
        res = felektro_name("Gniazdo przemysłowe przenośne 16A 3p /FE/", "F7.0108",
                            {"IP_opis": "IP67"})
        self.assertEqual(res["nazwa"], "Gniazdo przemysłowe przenośne 16A 3P IP67 F-Elektro")
        self.assertEqual(res["status"], "WARNING")
        # w rozdzielnicach IP z opisu nie wchodzi (opis podaje IP obudowy takze przy wersji z gniazdami IP44)
        self.assertEqual(felektro_name("Rozdzielnica LAGO 10M - gniazda 2x16A 5p", "F3.0853", {"IP_opis": "IP65"})["nazwa"],
                         "Rozdzielnica 2x16/5 LAGO 10M F-Elektro")
        # przelacznik zamykany z nazwy wygrywa z tytulem karty bez "-Z"
        self.assertEqual(felektro_name("Rozdzielnica RS-KOMBI - gniazda 16A 5p, 230V, L/P (16A-Z)", "F3.1893",
                                       {"_slug": "F-ELEKTRO-ROZDZIELNICA-RS-KOMBI-GNIAZDA-16A-5P-230V-L-P-16A-F3-1893"}
                                       )["nazwa"],
                         "Rozdzielnica 16/5 1x230V z przełącznikiem zamykanym L-0-P 16A RS-KOMBI F-Elektro")

    def test_card_loader(self):
        with tempfile.TemporaryDirectory() as tmp:
            Path(tmp, "adresy.json").write_text(
                '{"F3-0322": {"url": "u1", "slug": "F-ELEKTRO-rozdzielnica-lago-10m-gniazda-32a-5p-2x230v-ip65-F3-0322"},'
                ' "F7-0108": {"url": "u2", "slug": "F-ELEKTRO-gniazdo-przemyslowe-przenosne-16a-3p-fe-F7-0108"}}',
                encoding="utf-8")
            Path(tmp, "karty.json").write_text(
                '{"F7-0108": {"url": "u2", "ip_opis": ["IP67"]}, "F7-0013": {"url": "u3", "ip_opis": ["IP44", "IP67"]}}',
                encoding="utf-8")
            rows = [{"Kod producenta": c, "Kod towaru": f"{c}/FEL"} for c in ("F3.0322", "F7.0108", "F7.0013", "F1.0001")]
            feed = load_onninen_cards(Path(tmp), rows)
        self.assertEqual(feed["F3.0322"]["IP"], "IP65")
        self.assertEqual(feed["F7.0108"]["IP_opis"], "IP67")
        self.assertNotIn("IP_opis", feed["F7.0013"])  # opis z dwoma IP nie rozstrzyga
        self.assertNotIn("F1.0001", feed)

    def test_drums_and_extensions(self):
        self.assertEqual(felektro_name("Przedłużacz MAKARA-460 OPD 5x4 2x32A 5p, 230V, wtyczka 32A 5p - 40m",
                                       "14800-64")["nazwa"],
                         "Przedłużacz bębnowy trójfazowy 2x32/5 1x230V 40m 5x4mm2 guma H07RN-F MAKARA-460 F-Elektro")
        self.assertEqual(felektro_name("Przedłużacz FILO-230 OWY 3x1,5 4x230V -30m", "14508-45")["nazwa"],
                         "Przedłużacz bębnowy 4 gn. 30m 3x1,5mm2 PVC FILO-230 F-Elektro")
        self.assertEqual(felektro_name("Rozgałęźnik NERO-240 3x230V +termik +kontrolka", "14650")["nazwa"],
                         "Rozgałęźnik na bębnie 3 gn. z termikiem z kontrolką NERO-240 F-Elektro")
        self.assertEqual(felektro_name("Przedłużacz listwowy 4x230V OPD (H07RN-F) 3x2,5 /20m/", "34977-4")["nazwa"],
                         "Przedłużacz listwowy 4 gn. 20m 3x2,5mm2 guma H07RN-F F-Elektro")

    def test_accessories_from_ekonom_description(self):
        self.assertEqual(felektro_name("Reduktor wtyczka 16A 4p na gniazdo 32A 4p", "F7.0023")["nazwa"],
                         "Reduktor wtyczka 16A 4P na gniazdo 32A 4P F-Elektro")
        self.assertEqual(felektro_name("Szybkozłączka uniwersalna 0,14-4mm2 - 3-styki - opakowanie 10 szt",
                                       "F2.0630")["nazwa"],
                         "Szybkozłączka uniwersalna 0,14-4mm2 3 styki F-Elektro op. 10 szt.")
        self.assertEqual(felektro_name("Puszka hermet. V8, 400V, 80x45x41, biała.-klik", "F1.0262")["nazwa"],
                         "Puszka hermetyczna V8 400V 80x45x41 biała klik F-Elektro")
        self.assertEqual(felektro_name("Obudowa ULISSE 6M MIDI 355- gniazda mm x 135mm x 105mm", "74180.41")["nazwa"],
                         "Obudowa rozdzielnicy 355x135x105mm ULISSE 6M MIDI F-Elektro")


class ProperNames20261116Test(unittest.TestCase):
    """Przeglad nazw wlasnych 2026-11-16 (Kobi, GTV, Kanlux): seria / linia / producent nie gina z nazwy."""

    def test_kobi_series_in_parametric_types(self):
        feed = {"4666": {"_name": "Plafon LED NIVERA SQUARE 12W 3CCT IP54 czarny Kobi Premium", "Marka": "Kobi Premium"}}
        res = build_name(kobi_row("Plafon LED NIVERA SQUARE 12W 3CCT IP54 czarny Kobi Premium"), KOBI, {}, feed)
        self.assertTrue(res["nazwa"].endswith("Nivera Square Kobi Premium"), res["nazwa"])
        # rodzaj z nazwy feedu i kat "110ST" nie wchodza do serii; kody MHN zostaja WIELKIMI
        res = build_name(kobi_row("Naświetlacz LED DELYO MHN 30W 4000K IP65 Kobi Premium"), KOBI, {}, {})
        self.assertTrue(res["nazwa"].endswith("Delyo MHN Kobi Premium"), res["nazwa"])
        res = build_name(kobi_row("NEXTRACK NT4 MINI 20W 60ST 4000K CR90 CP"), KOBI, {}, {})
        self.assertTrue(res["nazwa"].endswith("Nextrack NT4 Mini Kobi"), res["nazwa"])
        # w dekor seria dalej niesie DALI (porzadki tylko w typach parametrycznych)
        self.assertNotIn("Dali", build_name(kobi_row("Panel LED CAPRI 40W 60x60 4000K DALI Kobi Pro"), KOBI, {}, {})["nazwa"])

    def test_kobi_proper_name_before_brand(self):
        # 2026-11-17: nazwa wlasna (seria) przed marka, marka na samym koncu
        res = build_name(kobi_row("Taśma LED TRAMO 320 COB 50m 6500K 12V IP20 Kobi Premium"), KOBI, {}, {})
        self.assertEqual(res["nazwa"], "Taśma LED COB 320LED/m 50m 6500K 12V IP20 Tramo Kobi Premium")
        res = build_name(kobi_row("Oprawa drogowa LED CYOTO 50W 4000K LX Kobi Premium"), KOBI, {}, {})
        self.assertEqual(res["nazwa"], "Oprawa uliczna LED 50W 4000K z czujnikiem Cyoto Kobi Premium")

    def test_kobi_review_20261117(self):
        # przeglad po "nazwa wlasna przed marka": zestaw z tubami, lacznik z cecha z Ekonoma, Windstar, akumulatory
        res = build_name(kobi_row("ZESTAW HERMETIC 2X120 + LED T8 18W 4000K"), KOBI, {}, {})
        self.assertEqual(res["nazwa"], "Oprawa hermetyczna z tubami LED 2x120cm 2x18W 4000K Hermetic Kobi")
        res = build_name(kobi_row("Zestaw hermetyczny HERMETIC G2 + 2x LED T8 18W 120cm 4000K KOBI"), KOBI, {}, {})
        self.assertIn("z tubami LED 2x120cm 2x18W", res["nazwa"])
        feed = {"4666": {"_name": "Łącznik do szynoprzewodu 3-obwodowy T PRAWY biały Kobi", "Marka": "Kobi"}}
        res = build_name(kobi_row("ŁĄCZNIK T ZEWNĘTRZN BIAŁY DO NEXTRACK NT"), KOBI, {}, feed)
        self.assertEqual(res["nazwa"], "Łącznik do szynoprzewodu 3-obwodowy T zewnętrzny prawy biały Nextrack NT Kobi")
        feed = {"4666": {"_name": "Grzejnik konwektorowy SONDO 2000W TURBO biały Kobi Windstar", "Marka": "Kobi Windstar"}}
        res = build_name(kobi_row("GRZEJNIK KONWEKTOROWY SONDO 2000W TURBO"), KOBI, {}, feed)
        self.assertEqual(res["nazwa"], "Grzejnik konwektorowy 2000W biały Sondo Turbo Kobi Windstar")
        feed = {"4666": {"_name": "Wentylator wieżowy LISSE 2 ST 45W czarny Kobi Windstar", "Marka": "Kobi Windstar"}}
        res = build_name(kobi_row("Wentylator wieżowy LISSE 2 ST 45W czarny Kobi Windstar"), KOBI, {}, feed)
        self.assertEqual(res["nazwa"], "Wentylator wieżowy 45W czarny Lisse 2 ST Kobi Windstar")
        res = build_name(kobi_row("AKUM. 18650 7,4V 5,4AH (SOLAR STREET 15W"), KOBI, {}, {})
        self.assertEqual(res["nazwa"], "Akumulator 18650 7,4V 5,4Ah do Solar Street 15W Kobi")
        res = build_name(kobi_row("KOSTKA HALOGENOWA Z WYSIĘGNIKIEM - K006"), KOBI, {}, {})
        self.assertEqual(res["nazwa"], "Kostka halogenowa z wysięgnikiem K006 Kobi")
        # "barwa NEUTRALNA BIAŁA" to barwa swiatla - nie seria i nie kolor obudowy
        res = build_name(kobi_row("Oprawa LED NEXFORCE N2 52W barwa NEUTRALNA BIAŁA"), KOBI, {}, {})
        self.assertNotRegex(res["nazwa"], r"(?i)neutraln|biał")

    def test_kobi_accessory_system_from_ekonom_name(self):
        feed = {"4666": {"_name": "Szynoprzewód 3-obwodowy 2m czarny Kobi", "Marka": "Kobi"}}
        res = build_name(kobi_row("SZYNA 3-FAZOWA 2M CZARNA DO NEXTRACK NT"), KOBI, {}, feed)
        self.assertTrue(res["nazwa"].endswith("Nextrack NT Kobi"), res["nazwa"])

    def test_gtv_bulb_line_and_wall_luminaire_model(self):
        res = build_name(gtv_row("ŹRÓDŁO ŚWIATŁA LED, FILAMENT, SPIN, C35, E14, 3,0W, 165lm, 2200K, AC220-240V, 360ST"),
                         GTV, {})
        self.assertTrue(res["nazwa"].endswith("GTV Spin"), res["nazwa"])
        res = build_name(gtv_row("OPRAWA ŚCIENNA JEDNOKIERUNKOWA, TIAGO, GU10, max. 35W, IP54, AC220-240V, 50/60Hz, INOX"),
                         GTV, {})
        self.assertIn("Tiago", res["nazwa"])

    def test_gtv_driver_maker_and_voltage_range(self):
        res = build_name(gtv_row("ZASILACZ STAŁOPRĄDOWY DALI REGULOWANY 550-1000mA 3-54V SELV KLASA 1, BOKE"), GTV, {})
        self.assertIn("GTV Boke", res["nazwa"])
        self.assertNotRegex(res["nazwa"], r"\b3-(\s|$)")
        res = build_name(gtv_row("ZASILACZ STAŁOPRĄDOWY DALI-PUSH-1-10V REGULOWANY 250-700mA 3-42V SELV KLASA 2, BOKE"),
                         GTV, {})
        self.assertIn("DALI-PUSH-1-10V", res["nazwa"])

    def test_kanlux_tube_and_breaker_line(self):
        res = build_name(kanlux_row("T8 LED GLASSv4 18W-NW Tuba LED 4000k 2880lm", "33212"), KANLUX, {}, {})
        self.assertTrue(res["nazwa"].endswith("Kanlux GLASSv4"), res["nazwa"])
        # linia za marka (rzeczownik z karty kanlux.com); gdy rzeczownik z nazwy w Ekonomie juz ja niesie, nie powtarza sie
        res = build_name(kanlux_row("KRD10-2/63/30-AC Wyłącznik różnicowoprądowy IDEAL TS 2P", "38732"), KANLUX, {}, {})
        self.assertEqual(res["nazwa"].count("IDEAL TS"), 1, res["nazwa"])
