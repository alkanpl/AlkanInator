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
    package_from_pdfs,
    lumen_from_window,
    lumens_from_pdf_codes,
    lumens_from_pdfs,
    move_brand_tail,
    diameter_in_mm,
    parse_ballast,
    parse_breaker,
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
        self.assertEqual(res["nazwa"], "Plafon LED 24W 1540lm 4000K IP44 z czujnikiem ruchu i zmierzchu czarny Kobi Premium")

    def test_feed_with_other_product_is_ignored(self):
        feed = {"4666": {"_name": "Żarówka LED GS 13W E27 6500K Kobi", "Strumień świetlny |lm|": "1400"}}
        res = build_name(kobi_row("LED GS 13W E27 6000K"), KOBI, {}, feed)
        self.assertNotIn("1400lm", res["nazwa"])
        self.assertTrue(any("inny produkt" in u for u in res["uwagi"]))

    def test_extension_cord_and_psu(self):
        res = build_name(kobi_row("Przedłużacz LINEA 3gn/5m/Zu+W czarny Kobi"), KOBI, {}, {})
        self.assertEqual(res["nazwa"], "Przedłużacz 3gn 5m Zu+W czarny Kobi")
        res = build_name(kobi_row("ZASILACZ DESKTOP 12V 72W 6,0A"), KOBI, {}, {})
        self.assertEqual(res["nazwa"], "Zasilacz LED desktop 12V 72W Kobi")

    def test_track_spot_degrees_st(self):
        res = build_name(kobi_row("Reflektor szynowy LED NEXTRACK CORE 20W 3CCT 36ST biały Kobi Pro"), KOBI, {}, {})
        self.assertEqual(res["nazwa"], "Reflektor szynowy 20W CCT 36° biały Kobi Pro")

    def test_caps_noun_does_not_eat_model(self):
        # rzeczownik ze slownika ma inna dlugosc niz tekst - wczesniej "AKUM. 18650" dawalo "Akumulator Kobi 0"
        res = build_name(kobi_row("AKUM. 18650 3,7V 3AH (SOLAR MHC 5W)"), KOBI, {}, {})
        self.assertTrue(res["nazwa"].startswith("Akumulator 18650 3,7V") and res["nazwa"].endswith(" Kobi"))
        res = build_name(kobi_row("SMART SOCKET 16A WIFI"), KOBI, {}, {})
        self.assertEqual(res["nazwa"], "Gniazdko Smart 16A WiFi Kobi")
        res = build_name(kobi_row("KOSTKA HALOGENOWA Z WYSIĘGNIKIEM - K006"), KOBI, {}, {})
        self.assertEqual(res["nazwa"], "Kostka halogenowa z wysięgnikiem - K006 Kobi")

    def test_brand_before_preposition(self):
        res = build_name(kobi_row("Rozeta sufitowa do LED KOLINE K2 biała Kobi Pro"), KOBI, {}, {})
        self.assertEqual(res["nazwa"], "Rozeta sufitowa do KOLINE K2 biała Kobi Pro")
        res = build_name(kobi_row("Adapter do szynoprzewodu 3-obwodowy biały Kobi"), KOBI, {}, {})
        self.assertEqual(res["nazwa"], "Adapter do szynoprzewodu 3-obwodowy biały Kobi")
        # przyimek bedacy czescia rodzaju oprawy zostaje w rzeczowniku
        res = build_name(kobi_row("Oprawa do nadbudowania OH12 1xGX53 biała Kobi"), KOBI, {}, {})
        self.assertEqual(res["nazwa"], "Oprawa do nadbudowania 1xGX53 biała Kobi OH12")

    def test_lumen_not_glued_to_ip(self):
        res = build_name(kobi_row("Kinkiet ogrodowy LED QUERK 2x2W 3000K IP54 550lm czarny LED2B"), KOBI, {}, {})
        self.assertEqual(res["nazwa"], "Kinkiet ogrodowy 2x2W 550lm 3000K IP54 czarny LED2B Querk")

    def test_tube_length_and_socket_fixture(self):
        res = build_name(kobi_row("LED T8 9W 60CM 4000K LED2B"), KOBI, {}, {})
        self.assertEqual(res["nazwa"], "Tuba LED T8 60cm 9W 4000K LED2B")
        # oprawa na wymienne zrodlo nie ma wlasnego strumienia - brak lm to nie blad
        res = build_name(kobi_row("Reflektor szynowy LAITILINE 1xGU10 biały Laitica"), KOBI, {}, {})
        self.assertEqual(res["nazwa"], "Reflektor szynowy 1xGU10 biały Laitica")
        self.assertEqual(res["status"], "OK")
        res = build_name(kobi_row("Kinkiet GRANDE 1xR7s Art of Light"), KOBI, {}, {})
        self.assertEqual(res["nazwa"], "Kinkiet 1xR7s Art of Light Grande")

    def test_extension_cord_without_earth(self):
        res = build_name(kobi_row("Przedłużacz LINEA G 1gn/20m/Bu/2x1mm pomarańczowy Kobi"), KOBI, {}, {})
        self.assertEqual(res["nazwa"], "Przedłużacz 1gn 20m Bu 2x1mm2 pomarańczowy Kobi")
        feed = {"4666": {"_name": "Przedłużacz LINEA PRO CUBE 3gn+3USB/1USB-C/Zu+W czarny Kobi", "Marka": "Kobi"}}
        res = build_name(kobi_row("Przedłużacz LINEA PRO CUBE 3gn+3USB/1USB.-C/Zu.+W czarny Kobi"), KOBI, {}, feed)
        self.assertEqual(res["nazwa"], "Przedłużacz 3gn+3USB/1USB-C Zu+W czarny Kobi")

    def test_version_number_gets_model_word(self):
        results = []
        for name in ["LED SOLARIS 2 5,5W 4000K", "LED SOLARIS 3 5,5W 4000K",
                     "QUAZAR 11 1XGU10 CZARNA", "QUAZAR 15 1XGU10 CZARNA"]:
            res = build_name(kobi_row(name), KOBI, {}, {})
            res["stara"] = name
            results.append(res)
        add_version_suffixes(results, KOBI)
        self.assertEqual(results[0]["nazwa"], "Plafon LED 5,5W 4000K Kobi SOLARIS 2")
        self.assertEqual(results[1]["nazwa"], "Plafon LED 5,5W 4000K Kobi SOLARIS 3")
        # seria jest juz w nazwie - numer idzie za nia, bez powtorzenia serii
        self.assertEqual(results[2]["nazwa"], "Kinkiet ogrodowy 1xGU10 czarny Kobi Quazar 11")
        self.assertEqual(results[3]["nazwa"], "Kinkiet ogrodowy 1xGU10 czarny Kobi Quazar 15")


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
        self.assertEqual(res["nazwa"], "Przedłużacz 3 gn. z uziemieniem z wyłącznikiem 1,0mm2 1,5m Bemko")
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
        self.assertEqual(res["nazwa"], "Uchwyt regulowany do opraw z uchwytem FI60mm na słupach z FI60mm Bemko")
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
