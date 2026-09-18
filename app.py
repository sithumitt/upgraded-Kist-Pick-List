"""Picklist PDF -> Sinhala Word (.docx) converter (Streamlit app).

Run with:  streamlit run picklist_converter.py
Needs:     pip install streamlit pdfplumber python-docx pandas
"""
import hashlib
import inspect
import io
import re
from pathlib import Path

import pandas as pd
import pdfplumber
import streamlit as st
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, Twips

FONT_SIZE = Pt(16)

# Column names used in the Word table and the on-screen preview
COL_ITEM = "භාණ්ඩය"
COL_CASES = "කේස්"
COL_PIECES = "කෑලි"

# Product Mapping Dictionary
PRODUCT_MAPPING = {
    "aloe vera drink 1l": "කෝමාරිකා බීම ලීටර් 1",
    "aloe vera drink 200ml": "කෝමාරිකා බීම 200",
    "aloe vera drink 500ml": "කෝමාරිකා බීම 500",
    "assortment 150gm": "ඇසෝඩ්මන්ට් බිස්කට් 150",
    "assortment 300gm": "ඇසෝඩ්මන්ට් බිස්කට් 300",
    "butter cookies 120gm": "බටර් කුකීස් 120",
    "cheese cuts tin 210gm": "චීස් බිස්කට් රතු ටින්",
    "chilli paste 60g": "චිලී පේස්ට් 60",
    "choco shorties 220gm": "චොකෝ ශෝටීස් 220",
    "chocolate chips cookies 120gm": "චොකලට් චිප් කුකීස් 120",
    "chocolate chips cookies 65gm": "චොකලට් චිප් කුකීස් 65",
    "chocolate cream 500gm": "චොකලට් ක්‍රීම්",
    "chocolate wafer 40gm": "චොකලට වේපස් 40",
    "chocolate wafer 90gm": "චොකලට වේපස් 90",
    "chocolate wafer 200gm": "චොකලට වේපස් 200",
    "chocolate marie 100gm": "චොකලට් මාරි 100GM",
    "chocolate wafer 360gm x 3pks": "චොකලට වේපස් 360",
    "cookie assortment blue 330gm": "කුකි ඇසෝඩ්මන්ට් නිල් 330",
    "coffee wafer 200gm": "කෝපි වේපර්ස්",
    "creamy choc 90gm": "ක්‍රීමි චොක් 90",
    "creamy choc 210gm": "ක්‍රීමි චොක් 210",
    "creamy choc 360gm x 4pks": "ක්‍රීමි චොක් 360",
    "creamy vanilla 360gm": "ක්‍රීම් වැනිලා 360",
    "falooda wafer 200gm": "ෆාලූඩා වේපස් 200",
    "faluda wafer 90gm": "ෆාලූඩා වේපස් 90",
    "faluda wafer 200gm": "ෆාලූඩා වේපස් 200",
    "ginger 80gm": "ඉඟුරු බිස්කට් 80",
    "ginger 250gm": "ඉඟුරු බිස්කට් 250",
    "green apple sparkling 250ml": "ඇපල් ස්පාක්ලින්",
    "kist cream cracker 125gm": "ක්‍රීම් ක්‍රැකර් 125",
    "kist cream cracker 500gm": "ක්‍රීම් ක්‍රැකර් 500",
    "kist gift assortment 400gm": "ගීෆ්ට් ඇසෝඩ්මන්ට් 400",
    "kithul treacle 340ml": "කිතුල් පැණි 340",
    "knuckles water 500ml": "වතුර 500",
    "knuckles water 1000ml": "වතුර 1000",
    "knuckles water 1500ml": "වතුර 1500",
    "lemon & mint nectar 1l": "ලෙමන් සහ මින්ට් බීම 1l",
    "lemon & mint nectar 200ml": "ලෙමන් සහ මින්ට් බීම 200",
    "lemon puff 100gm": "ලෙමන් පෆ් 100",
    "lemon puff 200gm": "ලෙමන් පෆ් 200",
    "magic choco fun 40gm": "චෝෆන් කහ පාට 40",
    "magic choco fun 100gm": "චෝෆන් කහ පාට 100",
    "magic choco nut 160gm": "චොකෝ නට් පොල් 160",
    "magic chocooh 100gm": "චොකො ඕ [සුදු පාට ]",
    "magic chokoblok chocolat 160gm": "චොකෝ බ්ලොක් චොකලට් 160",
    "magic chokoblok coffee 160gm": "චොකෝ බ්ලොක් කෝපි",
    "magic chokoblok vanilla 160gm": "චොකෝ බ්ලොක් වැනිලා",
    "magic chokstik choco 10gm": "චොකලට් චොක්ස්ටික්",
    "magic chokstik strawberry 10gm": "ස්ටෝබරි චොක්ස්ටික්",
    "magic chokstik vanilla 10gm": "වැනිලා චොක්ස්ටික්",
    "magic chonkz 120gm": "චෝන්ස් ලා නිල් 120",
    "magic chonkz 240gm": "චෝන්ස් ලා නිල් 240",
    "magic fingers 27gm": "ෆින්ගස් 27",
    "magic fingers 100gm": "ෆින්ගස් 100",
    "magic olo assortment 180gm": "ඔලෝ ඇසෝර්ට්මන් 180",
    "magic olo butter scotch 140gm": "ඕලෝ බටර් ස්කොච් 140",
    "magic olo chocolate 60gm": "ඕලෝ චොකලට් 60",
    "magic olo chocolate 140gm": "ඕලෝ චොකලට් 140",
    "magic olo classic 60gm": "ඔලෝ ක්ලැසික් 60",
    "magic olo classic 140gm": "ඔලෝ ක්ලැසික් 140",
    "magic olo strawberry 60gm": "ඔලෝ ස්ට්‍රෝබෙරි 60",
    "magic olo strawberry 140gm": "ඔලෝ ස්ට්‍රෝබෙරි 140",
    "magic olo white 140gm": "ඔලෝ වයිට් 140",
    "magic treats 90gm": "ටීට්ස් දම් පාට",
    "mango nectar 1l": "අඹ බීම 1l",
    "mango nectar 200ml": "අඹ බීම 200",
    "mango nectar 500ml": "අඹ බීම 500",
    "marie 50gm": "මාරි 50",
    "marie 100gm": "මාරි 100",
    "marie 200gm": "මාරි 200",
    "marie 500gm": "මාරි 500",
    "mayonnaise 200g pouch": "මයෝනාඊස් 200",
    "milk shorties 220gm": "කිරි ෂෝටීස් 220",
    "milki cookies 120gm": "මීල්කි කුකීස්",
    "mixed fruit jam cup 100g": "මිශ්‍ර ජෑම් කප් 100",
    "mixed fruit jam 200g": "මිශ්‍ර ජෑම් 200",
    "mixed fruit jam 300g": "මිශ්‍ර ජෑම් 300",
    "mixed fruit jam 510g": "මිශ්‍ර ජෑම් 510",
    "mixed fruit nectar 1l": "මිශ්‍ර පලතුරු බීම 1L",
    "mixed fruit nectar 200ml": "මිශ්‍ර පලතුරු බීම 200ML",
    "mixed fruit nectar 500ml": "මිශ්‍ර පලතුරු බීම 500ML",
    "nice 430gm": "නායිස් 430",
    "onion byte 30gm": "අනියන් බයිට් 30",
    "orange nectar 500ml": "දොඩම් බීම 500",
    "orange sparkling 250ml": "දොඩම් ස්පාක්ලින්",
    "passion fruit nectar 1l": "පැෂන්ෆෲට් නෙක්ටා 1L",
    "ride classic drink 250ml": "රයිට් නිල්",
    "ride redberry drink 250ml": "රයිට් රතු",
    "ride sugar free drink 250ml": "රයිට් සීනි නැති",
    "s/berry melon jam cup100g": "ස්ටෝබරි ජෑම් 100 C",
    "berry flv melon jam200g": "ස්ටෝබරි ජෑම් 200",
    "berry flv melon jam300g": "ස්ටෝබරි ජෑම් 300",
    "strawberry 200g": "ස්ටෝබරි ජෑම් 200",
    "strawberry 300g": "ස්ටෝබරි ජෑම් 300",
    "sesame cookies 120gm": "සෙසමිකුකීස්",
    "strawberry sparkling 250ml": "ස්ටෝබරි ස්පාක්ලින්",
    "strawberry wafer 40gm": "ස්ටෝබරි වේපස් 40",
    "strawberry wafer 90gm": "ස්ටෝබරි වේපස් 90",
    "strawberry wafer 200gm": "ස්ටෝබරි වේපස් 200",
    "strawberry wafer 360gm x 3pks": "ස්ටෝබරි වේපස් 360",
    "soya sauce squeeasy 180ml": "සෝස් පැකට් 180ml",
    "tomato sachet 15g": "සෝස් පැකට් 15",
    "tomato sauce 110gr - pouch": "සෝස් පැකට් 110",
    "tomato sauce 400g - pouch": "සෝස් පැක්ට් 400",
    "tomato sauce 400g ": "සෝස් වීදුරු බෝතලේ 400 ",
    "tomato sauce 200gr ": "සෝස් වීදුරු බෝතලේ 200 ",
    "vanilla wafer 40gm": "වැනිලා වේපස් 40",
    "vanilla wafer 90gm": "වැනිලා වේපස් 90",
    "vanilla wafer 200gm": "වැනිලා වේපස් 200",
    "vanilla wafer 360gm x 3pks": "වැනිලා වේපස් 360",
    "woodapple nectar 200ml": "දිවුල් බීම 200",
    "woodapple nectar 500ml": "දිවුල් බීම 500",
    "kist cream cracker 190gm": "ක්‍රීම් ක්‍රැකර් 190",
    "kist cream cracker 250gm": "ක්‍රීම් ක්‍රැකර් 250",
    "kist buddy cream cracker 80gm": "බඩි ක්‍රීම් ක්‍රැකර් 80",
    "kist buddy cream cracker 190gm": "බඩි ක්‍රීම් ක්‍රැකර් 190",
    "kist buddy cream cracker 70gm": "බඩි ක්‍රීම් ක්‍රැකර් 70",
    "chocolate puff 200gm": "චොකලට් පෆ් 200",
    "nice 100gm": "නායිස් 100",
    "shorties 270gm": "ශෝටීස් 270",
    "chocolate shorties 270gm": "චොකලට් ශෝටීස් 270",
    "teasty coffee 270gm": "ටේස්ටි කෝපි 270",
    "marie 300gm": "මාරි 300",
    "marie 70gm": "මාරි 70",
    "marie 400gm": "මාරි 400",
    "kist chocolate cream 100gm": "කිස්ට් චොකලට් ක්‍රීම් 100",
    "kist chocolate cream 400gm": "කිස්ට් චොකලට් ක්‍රීම් 400",
    "creamy choc (100g)": "ක්‍රීමි චොක් 100",
    "creamy choc (360g x 8pks)": "ක්‍රීමි චොක් 360 x 8",
    "kist chocolate cream 210gm": "කිස්ට් චොකලට් ක්‍රීම් 210",
    "creamy choc 490gm": "ක්‍රීමි චොක් 490",
    "cheese cuts 170gm": "චීස් කට්ස් 170",
    "onion byte 130gm": "අනියන් බයිට් 130",
    "cheese cutz 80gm": "චීස් කට්ස් 80",
    "hot chilli byte 25gm": "හොට් චිලී බයිට් 25",
    "cheese cutz 40gm": "චීස් කට්ස් 40",
    "wheels milk (60gm)": "වීල්ස් කිරි 60",
    "wheels milk 50gm": "වීල්ස් කිරි 50",
    "wheels milk 100gm": "වීල්ස් කිරි 100",
    "wheels chocolate (60gm)": "වීල්ස් චොකලට් 60",
    "wheels milk 240gm": "වීල්ස් කිරි 240",
    "wheels chocolate 100gm": "වීල්ස් චොකලට් 100",
    "wheels chocolate 50gm": "වීල්ස් චොකලට් 50",
    "milki cookies 300gm": "මීල්කි කුකීස් 300",
    "cookie assortment red 330gm": "කුකි ඇසෝඩ්මන්ට් රතු 330",
    "party carol 250gm": "පාටි කැරොල් 250",
    "choky magic 170gm": "චොකී මැජික් 170",
    "choky magic vanilla 170gm": "චොකී මැජික් වැනිලා 170",
    "choky magic coffee 170gm": "චොකී මැජික් කෝපි 170",
    "choky magic disc 100gm": "චොකී මැජික් ඩිස්ක් 100",
    "magic choky collection 185gm": "මැජික් චොකී කලෙක්ෂන් 185",
    "olo mint 140gm": "ඕලෝ මින්ට් 140",
    "olo banana 140gm": "ඕලෝ බනානා 140",
    "magic olo lemon 60gm": "ඔලෝ ලෙමන් 60",
    "vanilla wafer (360gm x 6pks)": "වැනිලා වේපස් 360 x 6",
    "kist knu water 5l": "කිස්ට් වතුර ලීටර් 5",
    "knuckles glass water 330ml": "වීදුරු බෝතල් වතුර 330",
    "knuckles glass water 500ml": "වීදුරු බෝතල් වතුර 500",
    "100% juice red apple 500ml": "රතු ඇපල් ජූස් 500",
    "100% juice orange 500ml": "දොඩම් ජූස් 500",
    "100% juice green apple 500ml": "කොළ ඇපල් ජූස් 500",
    "100% juice grape 500ml": "මිදි ජූස් 500",
    "orange sparkling drink 215ml": "දොඩම් ස්පාක්ලින් 215",
    "s/berry sparkling drink 215ml": "ස්ටෝබරි ස්පාක්ලින් 215",
    "lime sparkling drink 215ml": "දෙහි ස්පාක්ලින් 215",
    "apple sparkling drink 215ml": "ඇපල් ස්පාක්ලින් 215",
    "orange sparkling 4 pack": "දොඩම් ස්පාක්ලින් පැක් 4",
    "tonic sparkling drink 250ml": "ටොනික් ස්පාක්ලින් 250",
    "lime soda drink 250ml": "ලයිම් සෝඩා 250",
    "ginger beer 250ml": "ජින්ජර් බියර් 250",
    "ride classic double pack": "රයිට් නිල් ඩබල් පැක්",
    "ride celebrations pack": "රයිට් සෙලිබ්‍රේෂන් පැක්",
    "ride glow apple drink 250ml": "රයිට් ග්ලෝ ඇපල් 250",
    "hit ice blue 250ml": "හිට් අයිස් බ්ලූ 250",
    "economy tomato sauce 2lt": "ඉකොනොමි ටොමාටෝ සෝස් ලීටර් 2",
    "soya sauce 4.6kg": "සෝස් කිලෝග්‍රෑම් 4.6",
    "tomato sauce pack 4lt": "ටොමාටෝ සෝස් පැක් ලීටර් 4",
    "tomato sauce bsp 4.2lt": "ටොමාටෝ සෝස් ලීටර් 4.2",
    "economy tomato sauce 4.2lt": "ඉකොනොමි ටොමාටෝ සෝස් ලීටර් 4.2",
    "tomato sauce cate. pack 1.25kg": "ටොමාටෝ සෝස් කැටරින් පැක් 1.25",
    "tomato sachet 9g - 576 pcs": "සෝස් පැකට් 9 (576ක්)",
    "chili sauce 190gr": "චිලී සෝස් 190",
    "tomato sauce squeeasy 200g": "සෝස් ස්කීසි 200",
    "chilli sauce squeeasy 200g": "චිලී සෝස් ස්කීසි 200",
    "oyster sauce squeeasy 200g": "ඔයිස්ටර් සෝස් ස්කීසි 200",
    "bbq sauce squeeasy 200g": "බාබකියු සෝස් ස්කීසි 200",
    "devilled sauce squeeasy 200g": "ඩෙවිල්ඩ් සෝස් ස්කීසි 200",
    "kochchi sauce squeeasy 200g": "කොච්චි සෝස් ස්කීසි 200",
    "kottu sauce squeeasy 200g": "කොත්තු සෝස් ස්කීසි 200",
    "salsa sauce squeeasy 200g": "සල්සා සෝස් ස්කීසි 200",
    "chinese soya sauce 310ml": "චයිනීස් සෝයා සෝස් 310",
    "soya sauce 350ml": "සෝයා සෝස් 350",
    "hot chillie sauce 355g": "හොට් චිලී සෝස් 355",
    "tomato ketchup 375g": "ටොමාටෝ කැචප් 375",
    "chillie & garlic sauce 375g": "චිලී ඇන්ඩ් ගාර්ලික් සෝස් 375",
    "chillie sauce 375g": "චිලී සෝස් 375",
    "sweet and sour sauce 395g": "ස්වීට් ඇන්ඩ් සවර් සෝස් 395",
    "devilled sauce 375g": "ඩෙවිල්ඩ් සෝස් 375",
    "tomato sauce squeezable 370g": "ටොමාටෝ සෝස් ස්කීසි 370",
    "oyster sauce 375g": "ඔයිස්ටර් සෝස් 375",
    "mango bbq sauce 400g": "මැන්ගෝ බාබකියු සෝස් 400",
    "chilli sauce squeezable 375g": "චිලී සෝස් ස්කීසි 375",
    "tomato puree 415g": "ටොමාටෝ පියුරි 415",
    "soya sauce 725ml": "සෝයා සෝස් 725",
    "chillie sauce 835g": "චිලී සෝස් 835",
    "tomato sauce 865g": "ටොමාටෝ සෝස් 865",
    "mustard cream 150g": "මස්ටර්ඩ් ක්‍රීම් 150",
    "ambarella chutney 250g": "ඇඹරැල්ලා චට්නි 250",
    "mango chutney 250g": "අඹ චට්නි 250",
    "coconut treacle 340ml": "පොල් පැණි 340",
    "kithul treacle 170ml": "කිතුල් පැණි 170",
    "ambarella chutney 450g": "ඇඹරැල්ලා චට්නි 450",
    "mango chutney 460g": "අඹ චට්නි 460",
    "golden syrup 470gr": "ගෝල්ඩන් සිරප් 470",
    "coconut treacle 740ml": "පොල් පැණි 740",
    "kithul treacle 740ml": "කිතුල් පැණි 740",
    "mango chutney 890g": "අඹ චට්නි 890",
    "milca full cream 150g": "මිල්කා ෆුල් ක්‍රීම් 150",
    "bonlac skimmed milk 200g": "බොන්ලැක් ස්කිම්ඩ් මිල්ක් 200",
    "bonlac skimmed milk 400g": "බොන්ලැක් ස්කිම්ඩ් මිල්ක් 400",
    "milca full cream 400g": "මිල්කා ෆුල් ක්‍රීම් 400",
    "chicken flv noodles 85g": "චිකන් නූඩ්ල්ස් 85",
    "vegetable flv noodles 85g": "වෙස්ටබල් නූඩ්ල්ස් 85",
    "prawn flv noodles 85g": "ප්‍රෝන් නූඩ්ල්ස් 85",
    "dry noodles 400g": "ඩ්‍රයි නූඩ්ල්ස් 400",
}

# Used only when no key in PRODUCT_MAPPING matched (e.g. "STRAWBERRY JAM 300G")
STRAWBERRY_JAM_FALLBACK = {
    "200": "ස්ටෝබරි ජෑම් 200",
    "300": "ස්ටෝබරි ජෑම් 300",
}

BATCH_RE = re.compile(r"\b([A-Z]{2}\d)\b")
NUMBER_RE = re.compile(r"^\d+(?:[.,]\d+)?$")
DATE_RE = re.compile(
    r"^(?:\d{1,2}[/.\-])?\d{1,2}[/.\-]\d{4}$|^\d{4}[/.\-]\d{1,2}(?:[/.\-]\d{1,2})?$"
)
# A real product row starts with: row number + product code, e.g. "35 F15522602 ..."
ROW_RE = re.compile(r"^(\d+)\s+([A-Za-z]{1,3}\d{4,})\s+(.*)$")
# Lines that start a page header (the table header repeats on every page and turns
# the table back on)
PAGE_HEADER_STARTS = ("picklist", "printed by", "warehouse:")
# Column-unit / currency header rows such as "(LKR)" and "CA/KG EA/GM CA/KG EA/GM"
UNIT_HEADER_RE = re.compile(r"ca/kg|ea/gm|\(lkr\)")
BOILERPLATE_RE = re.compile(
    r"\b(?:product code|product description|select product|filter|mrp|conv|"
    r"selling qty|sampling qty|total qty|batch|expiry date|page|date|"
    r"grand total|sub total)\b"
)


# --------------------------------------------------------------------------
# Text matching
# --------------------------------------------------------------------------
def clean_text_for_matching(text):
    if not text:
        return ""
    cleaned = re.sub(r'["()]', "", text)  # quotes and brackets: "(60gm)" == "60gm"
    cleaned = cleaned.replace("/", " ").replace(".", " ").replace("-", " ")
    return " ".join(cleaned.lower().split())


def _build_matchers(mapping):
    """Compile every key once and sort LONGEST FIRST, so that a more specific
    name ("chocolate shorties 270gm") always wins over a shorter name that is
    contained in it ("shorties 270gm")."""
    matchers = []
    for english_key, sinhala_val in mapping.items():
        key = clean_text_for_matching(english_key)
        pattern = re.compile(r"(?<![a-z0-9])" + re.escape(key) + r"(?!\d)")
        matchers.append((len(key), pattern, sinhala_val))
    matchers.sort(key=lambda m: m[0], reverse=True)
    return [(pattern, sinhala_val) for _, pattern, sinhala_val in matchers]


MATCHERS = _build_matchers(PRODUCT_MAPPING)


def strawberry_jam_fallback(cleaned_line):
    if "strawberry" not in cleaned_line:
        return None
    if "wafer" in cleaned_line or "sparkling" in cleaned_line:
        return None
    for size, sinhala_val in STRAWBERRY_JAM_FALLBACK.items():
        if re.search(rf"(?<!\d){size}\s?g", cleaned_line):
            return sinhala_val
    return None


def find_sinhala_name(cleaned_line):
    for pattern, sinhala_val in MATCHERS:
        if pattern.search(cleaned_line):
            return sinhala_val
    return strawberry_jam_fallback(cleaned_line)


def is_boilerplate(line):
    if not line.strip() or len(line.strip()) < 3:
        return True
    lower_line = line.lower()
    return bool(UNIT_HEADER_RE.search(lower_line) or BOILERPLATE_RE.search(lower_line))


# --------------------------------------------------------------------------
# Line parsing
# --------------------------------------------------------------------------
def _fmt_qty(token):
    token = token.replace(",", "")
    return re.sub(r"\.0+$", "", token)  # "5.00" -> "5"


def extract_quantities(line):
    """Return (cases, pieces) = the last two whole-number tokens before the
    batch code. Returns ("?", "?") when they can't be found, instead of
    silently writing 0."""
    batch = BATCH_RE.search(line)
    text = line[: batch.start()] if batch else line
    tokens = [t for t in text.split() if not DATE_RE.match(t)]
    numbers = [t for t in tokens if NUMBER_RE.match(t)]
    if len(numbers) < 2:
        return "?", "?"
    return _fmt_qty(numbers[-2]), _fmt_qty(numbers[-1])


def extract_no_and_description(raw_line):
    """Extract (No, English description) from a product row.
    "35 F15522602 FALUDA WAFER 90GM 108.00 24 0 6 0 0 0 6 CG1 01/08/2028"
        -> ("35", "FALUDA WAFER 90GM")
    The product code, MRP, quantities, batch and expiry date are dropped."""
    line = raw_line.strip()
    match = ROW_RE.match(line)
    if match:
        item_no, _code, rest = match.groups()
    else:
        item_no, rest = "", line

    # Strip the trailing MRP / quantities / batch code / expiry date
    tokens = rest.replace("`", "").split()
    while tokens and (
        NUMBER_RE.match(tokens[-1])
        or DATE_RE.match(tokens[-1])
        or re.fullmatch(r"[A-Z]{2}\d", tokens[-1])
    ):
        tokens.pop()

    return item_no, " ".join(tokens).upper()


def parse_picklist(pdf_bytes):
    """Read the PDF and return (rows, unmatched, pdf_items, english_rows).

    rows          - converted items for the Sinhala file: {Sinhala name, cases, pieces}
    unmatched     - product rows with no entry in PRODUCT_MAPPING: (no, description)
    pdf_items     - total number of product rows found in the PDF's item table
    english_rows  - EVERY product row as (name as printed in the PDF, cases, pieces),
                    for the English file (no mapping needed)
    """
    rows = []
    unmatched = []
    english_rows = []
    pdf_items = 0
    in_target_table = False

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            text_content = page.extract_text()
            if not text_content:
                continue

            for line in text_content.split("\n"):
                normalized_line = line.replace('"', "").strip()
                lower_line = normalized_line.lower()

                if lower_line.startswith(PAGE_HEADER_STARTS) or any(
                    k in lower_line for k in ("invoice", "customer name", "sales route")
                ):
                    in_target_table = False
                    continue

                if any(k in lower_line for k in ("product description", "selling qty", "total qty")):
                    in_target_table = True
                    continue

                # A real product row ("35 F15522602 FALUDA WAFER 90GM ...") is never
                # skipped as boilerplate, so it is always counted.
                looks_like_row = bool(ROW_RE.match(normalized_line))
                if not in_target_table or (not looks_like_row and is_boilerplate(normalized_line)):
                    continue

                sinhala_val = find_sinhala_name(clean_text_for_matching(normalized_line))
                if not (sinhala_val or looks_like_row):
                    continue  # page headers, unit rows, totals ...

                pdf_items += 1
                qty1, qty2 = extract_quantities(normalized_line)
                item_no, description = extract_no_and_description(normalized_line)
                description = description or normalized_line.upper()
                english_rows.append((description, qty1, qty2))

                if sinhala_val:
                    rows.append({COL_ITEM: sinhala_val, COL_CASES: qty1, COL_PIECES: qty2})
                else:
                    unmatched.append((item_no, description))

    return rows, unmatched, pdf_items, english_rows


# --------------------------------------------------------------------------
# Word document
# --------------------------------------------------------------------------
# Column widths in twips (1 inch = 1440), copied from the reference Word file:
#   item 4878 | cases 2160 | pieces 2322   (table width 9360)
MAIN_TABLE_WIDTHS = (Twips(4878), Twips(2160), Twips(2322))
MISSING_TABLE_WIDTHS = (Twips(1440), Twips(7920))  # same total width, 9360


def format_run(run, bold=False):
    """Set 16pt (and bold). Sinhala is a complex script, so Word reads the
    complex-script properties (szCs / bCs), not just sz / b. Set both."""
    run.font.size = FONT_SIZE
    rpr = run._element.get_or_add_rPr()
    if rpr.find(qn("w:szCs")) is None:
        rpr.find(qn("w:sz")).addnext(OxmlElement("w:szCs"))
    rpr.find(qn("w:szCs")).set(qn("w:val"), str(int(FONT_SIZE.pt * 2)))

    if bold:
        run.font.bold = True
        if rpr.find(qn("w:bCs")) is None:
            rpr.find(qn("w:b")).addnext(OxmlElement("w:bCs"))


def write_cell(cell, text, bold=False):
    cell.text = str(text).strip()
    for paragraph in cell.paragraphs:
        for run in paragraph.runs:
            format_run(run, bold)


def set_column_widths(table, widths):
    """Fixed layout with exact widths on the grid, on every cell and on the table
    itself - the same structure as the reference Word file. Call after all rows
    exist (rows added later don't inherit per-cell widths)."""
    table.autofit = False  # <w:tblLayout w:type="fixed"/>
    for idx, width in enumerate(widths):
        table.columns[idx].width = width
        for cell in table.columns[idx].cells:
            cell.width = width
    tbl_w = table._tbl.tblPr.find(qn("w:tblW"))
    if tbl_w is not None:
        tbl_w.set(qn("w:type"), "dxa")
        tbl_w.set(qn("w:w"), str(sum(w.twips for w in widths)))


def add_items_table(doc, headers, body_rows):
    """3-column table (item | cases | pieces) with the reference column widths."""
    table = doc.add_table(rows=1, cols=3)
    table.style = "Table Grid"
    for cell, text in zip(table.rows[0].cells, headers):
        write_cell(cell, text, bold=True)
    for item, cases, pieces in body_rows:
        cells = table.add_row().cells
        write_cell(cells[0], item)
        write_cell(cells[1], cases)
        write_cell(cells[2], pieces)
    set_column_widths(table, MAIN_TABLE_WIDTHS)
    return table


def _save(doc):
    stream = io.BytesIO()
    doc.save(stream)
    return stream.getvalue()


def build_docx(rows, unmatched):
    """Sinhala Word file (+ a table of items that are not in PRODUCT_MAPPING)."""
    doc = Document()
    doc.add_heading("පික් ලිස්ට් එකේ බඩු", level=1)
    add_items_table(
        doc,
        (COL_ITEM, COL_CASES, COL_PIECES),
        [(r[COL_ITEM], r[COL_CASES], r[COL_PIECES]) for r in rows],
    )

    # Items that were not found in PRODUCT_MAPPING, as a [No, Description] table
    if unmatched:
        doc.add_paragraph()
        heading_para = doc.add_paragraph()
        run_h = heading_para.add_run("Unmatched Items:")
        format_run(run_h, bold=True)

        missing_table = doc.add_table(rows=1, cols=2)
        missing_table.style = "Table Grid"
        write_cell(missing_table.rows[0].cells[0], "No", bold=True)
        write_cell(missing_table.rows[0].cells[1], "Description", bold=True)

        for item_no, desc in unmatched:
            r_cells = missing_table.add_row().cells
            write_cell(r_cells[0], item_no)
            write_cell(r_cells[1], desc)

        set_column_widths(missing_table, MISSING_TABLE_WIDTHS)

    return _save(doc)


def build_docx_en(english_rows):
    """English Word file: every item, names exactly as printed in the PDF."""
    doc = Document()
    doc.add_heading("Picklist Items", level=1)
    add_items_table(doc, ("Item", "Cases", "Pieces"), english_rows)
    return _save(doc)


# --------------------------------------------------------------------------
# Item count summary (PDF vs Word)
# --------------------------------------------------------------------------
def count_word_items(docx_bytes):
    """Number of item rows in the converted table of the generated Word file.
    The file is re-opened and counted, so this is what is really inside it."""
    if not docx_bytes:
        return 0
    table = Document(io.BytesIO(docx_bytes)).tables[0]
    return max(len(table.rows) - 1, 0)  # minus the header row


def build_summary(pdf_items, docx_bytes, docx_en_bytes=None):
    """pdf_items     - total items in the input PDF
    word_items    - total items in the (Sinhala) Word file
    missing       - items that did not make it into the Sinhala Word file (0 when none)
    english_items - total items in the English Word file"""
    word_items = count_word_items(docx_bytes)
    return {
        "pdf_items": pdf_items,
        "word_items": word_items,
        "missing": max(pdf_items - word_items, 0),
        "english_items": count_word_items(docx_en_bytes),
    }


# --------------------------------------------------------------------------
# Streamlit UI  (soft, glass-style look)
# --------------------------------------------------------------------------
CSS = """
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=Noto+Sans+Sinhala:wght@400;500;600;700&display=swap');

:root {
  --ink: #2f3356;
  --muted: #7377a0;
  --glass: rgba(255, 255, 255, 0.55);
  --glass-strong: rgba(255, 255, 255, 0.78);
  --edge: rgba(255, 255, 255, 0.85);
  --shadow: 0 10px 40px rgba(112, 120, 190, 0.16);
  --radius: 24px;
  --lav: #e7eaff;   --lav-ink: #5561c7;
  --mint: #dcf6ec;  --mint-ink: #2c8a68;
  --peach: #ffe9dd; --peach-ink: #c0623a;
}

html, body, .stApp {
  font-family: 'Inter', 'Noto Sans Sinhala', 'Iskoola Pota', 'Nirmala UI', system-ui, sans-serif;
  color-scheme: light;
}
.stApp {
  background:
    radial-gradient(900px 520px at 6% -8%, #d8e1ff 0%, rgba(216, 225, 255, 0) 60%),
    radial-gradient(800px 480px at 100% 2%, #ffe0ef 0%, rgba(255, 224, 239, 0) 60%),
    radial-gradient(900px 600px at 55% 112%, #d2f4ea 0%, rgba(210, 244, 234, 0) 60%),
    #f5f6fc;
  background-attachment: fixed;
  color: var(--ink);
}
.stApp p, .stApp label, .stApp li { color: var(--ink); }

/* hide default Streamlit chrome so it feels like a real app */
#MainMenu, footer, [data-testid="stToolbar"], [data-testid="stDecoration"] { display: none !important; }
header[data-testid="stHeader"] { background: transparent !important; }
.block-container { max-width: 940px; padding: 2.4rem 1.5rem 4rem; }

/* ---------- hero ---------- */
.hero {
  display: flex; align-items: center; gap: 1.2rem;
  background: var(--glass);
  backdrop-filter: blur(18px) saturate(140%); -webkit-backdrop-filter: blur(18px) saturate(140%);
  border: 1px solid var(--edge); border-radius: var(--radius);
  box-shadow: var(--shadow); padding: 1.5rem 1.7rem; margin-bottom: 1.1rem;
}
.hero-icon {
  flex: 0 0 auto; width: 64px; height: 64px; border-radius: 20px;
  display: flex; align-items: center; justify-content: center; font-size: 30px;
  background: linear-gradient(145deg, #dfe5ff, #f1e6ff);
  box-shadow: inset 0 1px 0 rgba(255, 255, 255, 0.9), 0 8px 20px rgba(143, 160, 255, 0.25);
}
.hero-title { font-size: 1.55rem; font-weight: 700; line-height: 1.3; color: var(--ink); }
.hero-sub { margin-top: 0.25rem; font-size: 0.98rem; color: var(--muted); }

/* ---------- steps ---------- */
.steps { display: flex; align-items: center; justify-content: center; gap: 0.6rem; flex-wrap: wrap; margin: 0.2rem 0 1.1rem; }
.step {
  display: flex; align-items: center; gap: 0.55rem; padding: 0.42rem 0.95rem 0.42rem 0.45rem;
  border-radius: 999px; font-size: 0.9rem; font-weight: 500; color: var(--muted);
  background: rgba(255, 255, 255, 0.45); border: 1px solid var(--edge);
}
.step-dot {
  width: 24px; height: 24px; border-radius: 50%; display: inline-flex; align-items: center; justify-content: center;
  font-size: 0.78rem; font-weight: 600; background: #e9ebf7; color: var(--muted);
}
.step.active { color: var(--lav-ink); background: var(--glass-strong); box-shadow: 0 6px 18px rgba(143, 160, 255, 0.22); }
.step.active .step-dot { background: linear-gradient(135deg, #8fa0ff, #b79cff); color: #fff; }
.step.done { color: var(--mint-ink); }
.step.done .step-dot { background: var(--mint); color: var(--mint-ink); }
.step-line { width: 26px; height: 2px; border-radius: 2px; background: rgba(143, 160, 255, 0.28); }

/* ---------- file uploader ---------- */
[data-testid="stFileUploader"] label,
[data-testid="stFileUploader"] label p { color: var(--ink) !important; font-weight: 600; }
[data-testid="stFileUploaderDropzone"] {
  background: var(--glass) !important;
  backdrop-filter: blur(16px); -webkit-backdrop-filter: blur(16px);
  border: 2px dashed #b9c1f2 !important; border-radius: var(--radius) !important;
  padding: 1.7rem 1.5rem !important; transition: all 0.25s ease;
}
[data-testid="stFileUploaderDropzone"]:hover {
  border-color: #8f9bf0 !important; background: var(--glass-strong) !important; box-shadow: var(--shadow);
}
[data-testid="stFileUploaderDropzone"] *,
[data-testid="stFileUploaderDropzone"] small { color: var(--muted) !important; }
[data-testid="stFileUploaderDropzone"] button {
  border-radius: 999px !important; border: 0 !important; padding: 0.5rem 1.3rem !important; font-weight: 600;
  background: linear-gradient(135deg, #7385f2, #9d80f7) !important;
  box-shadow: 0 8px 20px rgba(143, 160, 255, 0.35);
}
[data-testid="stFileUploaderDropzone"] button,
[data-testid="stFileUploaderDropzone"] button * { color: #ffffff !important; }
[data-testid="stFileUploaderFile"], [data-testid="stFileChip"] {
  background: var(--glass-strong) !important; border: 1px solid var(--edge) !important; border-radius: 16px !important;
  box-shadow: 0 4px 14px rgba(112, 120, 190, 0.10);
}
[data-testid="stFileUploaderFile"] *, [data-testid="stFileChip"] * { color: var(--ink) !important; }
[data-testid="stFileChip"] > div:first-child { background: var(--lav) !important; border-radius: 12px !important; }
[data-testid="stFileChip"] svg { color: var(--lav-ink) !important; }
[data-testid="stFileChipName"] { font-weight: 600; }
/* remove (x) and add (+) buttons: soft, not the big gradient pill */
[data-testid="stFileChipDeleteBtn"] button, [data-testid="stFileUploaderDeleteBtn"] button,
[data-testid="stBaseButton-borderlessIcon"] {
  background: rgba(255, 255, 255, 0.85) !important; box-shadow: none !important; padding: 0.25rem 0.6rem !important;
}
[data-testid="stFileChipDeleteBtn"] button *, [data-testid="stFileUploaderDeleteBtn"] button *,
[data-testid="stBaseButton-borderlessIcon"], [data-testid="stBaseButton-borderlessIcon"] * {
  color: var(--lav-ink) !important;
}
[data-testid="stSpinner"] * { color: var(--muted) !important; }

/* ---------- status pills ---------- */
.status {
  display: flex; align-items: center; gap: 0.6rem; padding: 0.8rem 1.15rem; margin-bottom: 0.7rem;
  border-radius: 18px; font-size: 0.97rem; font-weight: 500;
  backdrop-filter: blur(14px); -webkit-backdrop-filter: blur(14px);
  border: 1px solid var(--edge); box-shadow: 0 6px 22px rgba(112, 120, 190, 0.10);
}
.status.ok   { background: rgba(220, 246, 236, 0.75); color: var(--mint-ink); }
.status.warn { background: rgba(255, 233, 221, 0.80); color: var(--peach-ink); }

/* ---------- cards ---------- */
.glass {
  background: var(--glass);
  backdrop-filter: blur(18px) saturate(140%); -webkit-backdrop-filter: blur(18px) saturate(140%);
  border: 1px solid var(--edge); border-radius: var(--radius);
  box-shadow: var(--shadow); padding: 1.25rem 1.4rem;
}
.section-title { font-size: 1.05rem; font-weight: 600; color: var(--ink); margin: 1.2rem 0 0.7rem; }
.section-hint { font-size: 0.88rem; color: var(--muted); margin: -0.3rem 0 0.7rem; }

.stat-grid { display: grid; grid-template-columns: repeat(3, minmax(0, 1fr)); gap: 1rem; }
.stat {
  border-radius: var(--radius); padding: 1.15rem 1.3rem 1.05rem; border: 1px solid var(--edge);
  backdrop-filter: blur(18px); -webkit-backdrop-filter: blur(18px);
  box-shadow: var(--shadow); transition: transform 0.2s ease, box-shadow 0.2s ease;
}
.stat:hover { transform: translateY(-3px); box-shadow: 0 16px 44px rgba(112, 120, 190, 0.22); }
.stat-icon {
  width: 40px; height: 40px; border-radius: 14px; display: flex; align-items: center; justify-content: center;
  font-size: 19px; background: rgba(255, 255, 255, 0.75); margin-bottom: 0.7rem;
}
.stat-num { font-size: 2.5rem; font-weight: 700; line-height: 1.05; }
.stat-label { margin-top: 0.4rem; font-size: 0.95rem; font-weight: 600; color: var(--ink); }
.stat-sub { font-size: 0.8rem; color: var(--muted); margin-top: 0.1rem; }
.stat.lav  { background: linear-gradient(150deg, rgba(231, 234, 255, 0.92), rgba(255, 255, 255, 0.5)); }
.stat.lav .stat-num { color: var(--lav-ink); }
.stat.mint { background: linear-gradient(150deg, rgba(220, 246, 236, 0.92), rgba(255, 255, 255, 0.5)); }
.stat.mint .stat-num { color: var(--mint-ink); }
.stat.ok   { background: linear-gradient(150deg, rgba(220, 246, 236, 0.92), rgba(255, 255, 255, 0.5)); }
.stat.ok .stat-num { color: var(--mint-ink); }
.stat.warn { background: linear-gradient(150deg, rgba(255, 233, 221, 0.95), rgba(255, 255, 255, 0.5)); }
.stat.warn .stat-num { color: var(--peach-ink); }

/* ---------- tables ---------- */
.table-wrap {
  max-height: 430px; overflow: auto; border-radius: 18px;
  border: 1px solid var(--edge); background: rgba(255, 255, 255, 0.5);
}
.table-wrap::-webkit-scrollbar { width: 8px; height: 8px; }
.table-wrap::-webkit-scrollbar-thumb { background: rgba(143, 160, 255, 0.35); border-radius: 8px; }
table.glass-table { width: 100%; border-collapse: separate; border-spacing: 0; font-size: 0.95rem; color: var(--ink); }
table.glass-table thead th {
  position: sticky; top: 0; z-index: 1; text-align: left !important; padding: 0.75rem 1.05rem;
  font-weight: 600; color: var(--lav-ink); background: rgba(233, 236, 255, 0.96);
}
table.glass-table th, table.glass-table td { border: 0 !important; }
table.glass-table td { padding: 0.6rem 1.05rem; border-top: 1px solid rgba(160, 168, 220, 0.18) !important; }
table.glass-table tbody tr:nth-child(even) td { background: rgba(255, 255, 255, 0.4); }
table.glass-table tbody tr:hover td { background: rgba(200, 208, 255, 0.28); }

/* ---------- download button ---------- */
.stDownloadButton button, [data-testid="stDownloadButton"] button {
  background: linear-gradient(135deg, #7385f2 0%, #9d80f7 100%) !important;
  border: 0 !important; border-radius: 18px !important; padding: 0.95rem 1.4rem !important;
  font-weight: 600; font-size: 1rem; box-shadow: 0 12px 28px rgba(115, 133, 242, 0.38);
  transition: transform 0.15s ease, box-shadow 0.15s ease;
}
.stDownloadButton button:hover, [data-testid="stDownloadButton"] button:hover {
  transform: translateY(-2px); box-shadow: 0 16px 34px rgba(143, 160, 255, 0.5);
}
.stDownloadButton button *, [data-testid="stDownloadButton"] button * { color: #ffffff !important; }

/* second (English) download button: soft teal so the two options are easy to tell apart */
.st-key-dl_en button,
[data-testid="stColumn"]:nth-child(2) [data-testid="stDownloadButton"] button,
[data-testid="column"]:nth-child(2) .stDownloadButton button {
  background: linear-gradient(135deg, #4fb69a 0%, #63a4e6 100%) !important;
  box-shadow: 0 12px 28px rgba(79, 182, 154, 0.35) !important;
}
.st-key-dl_en button:hover,
[data-testid="stColumn"]:nth-child(2) [data-testid="stDownloadButton"] button:hover,
[data-testid="column"]:nth-child(2) .stDownloadButton button:hover {
  box-shadow: 0 16px 34px rgba(79, 182, 154, 0.5) !important;
}
.dl-hint { text-align: center; font-size: 0.85rem; color: var(--muted); margin-top: 0.55rem; }

@media (max-width: 720px) {
  .stat-grid { grid-template-columns: 1fr; }
  .hero { flex-direction: column; text-align: center; }
}
"""


def html_table(df):
    """DataFrame -> styled, scrollable HTML table (text is escaped, so PDF content can't inject HTML)."""
    html = df.to_html(index=False, border=0, classes="glass-table", escape=True)
    return '<div class="table-wrap">' + re.sub(r">\s+<", "><", html) + "</div>"


def hero_html():
    return (
        '<div class="hero"><div class="hero-icon">📋</div><div>'
        '<div class="hero-title">පික් ලිස්ට් එකේ බඩු පරිවර්තකය</div>'
        '<div class="hero-sub">ඔබේ Picklist PDF එක සිංහල Word ගොනුවක් බවට ක්ෂණිකව පරිවර්තනය කරන්න</div>'
        "</div></div>"
    )


def steps_html(active):
    labels = ["PDF තෝරන්න", "පරිවර්තනය", "Word බාගත කරන්න"]
    parts = []
    for i, label in enumerate(labels, start=1):
        state = "done" if i < active else ("active" if i == active else "")
        mark = "✓" if i < active else str(i)
        parts.append(f'<div class="step {state}"><span class="step-dot">{mark}</span>{label}</div>')
        if i < len(labels):
            parts.append('<div class="step-line"></div>')
    return '<div class="steps">' + "".join(parts) + "</div>"


def stat_card(tone, icon, number, label_si, label_en):
    return (
        f'<div class="stat {tone}"><div class="stat-icon">{icon}</div>'
        f'<div class="stat-num">{number}</div>'
        f'<div class="stat-label">{label_si}</div>'
        f'<div class="stat-sub">{label_en}</div></div>'
    )


def summary_html(summary):
    missing = summary["missing"]
    return (
        '<div class="stat-grid">'
        + stat_card("lav", "📄", summary["pdf_items"], "PDF එකේ මුළු භාණ්ඩ ගණන", "Total items in PDF")
        + stat_card("mint", "📝", summary["word_items"], "Word එකේ මුළු භාණ්ඩ ගණන", "Total items in Word file")
        + stat_card(
            "ok" if missing == 0 else "warn",
            "✅" if missing == 0 else "⚠️",
            missing,
            "මගහැරුණු භාණ්ඩ ගණන",
            "Missing items" if missing else "Nothing missing",
        )
        + "</div>"
    )


def unmatched_html(unmatched):
    df = pd.DataFrame(unmatched, columns=["No", "Description"])
    return (
        '<div class="section-title">⚠️ නාමාවලියේ නැති අයිතම</div>'
        '<div class="section-hint">මේවා PRODUCT_MAPPING එකට එක් කළ විට Word ගොනුවට ඇතුළත් වේ.</div>'
        + html_table(df)
    )


@st.cache_data(show_spinner=False)
def process_pdf(pdf_bytes, code_version):
    """Cached, so clicking the download button doesn't re-parse the PDF."""
    rows, unmatched, pdf_items, english_rows = parse_picklist(pdf_bytes)
    docx_bytes = build_docx(rows, unmatched) if rows else None
    docx_en_bytes = build_docx_en(english_rows) if english_rows else None
    summary = build_summary(pdf_items, docx_bytes, docx_en_bytes)
    return rows, unmatched, docx_bytes, docx_en_bytes, summary


# Changes whenever this file changes, so a cached result from an older version of
# the code can never be shown again (st.cache_data only tracks the function's own
# source, not the helpers it calls). Must NOT start with "_" or it is ignored.
CODE_VERSION = hashlib.md5(Path(__file__).read_bytes()).hexdigest()


def stretch_kwargs():
    # Newer Streamlit deprecates use_container_width in favour of width="stretch"
    if "width" in inspect.signature(st.download_button).parameters:
        return {"width": "stretch"}
    return {"use_container_width": True}


DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"


def render_downloads(docx_si, docx_en, english_count):
    """Two download buttons: Sinhala Word file and English Word file."""
    label_si = "📥 සිංහල Word ගොනුව බාගත කරන්න (Sinhala)"
    label_en = "📥 ඉංග්‍රීසි Word ගොනුව බාගත කරන්න (English)"
    if docx_si and docx_en:
        col_si, col_en = st.columns(2)
        with col_si:
            st.download_button(label_si, data=docx_si, file_name="පික්_ලිස්ට්_එකේ_බඩු.docx",
                               mime=DOCX_MIME, key="dl_si", **stretch_kwargs())
        with col_en:
            st.download_button(label_en, data=docx_en, file_name="picklist_items_english.docx",
                               mime=DOCX_MIME, key="dl_en", **stretch_kwargs())
    elif docx_en:  # nothing matched the Sinhala catalog, but the English file still works
        st.download_button(label_en, data=docx_en, file_name="picklist_items_english.docx",
                           mime=DOCX_MIME, key="dl_en", **stretch_kwargs())
    if docx_en:
        st.markdown(
            f'<div class="dl-hint">ඉංග්‍රීසි ගොනුවේ PDF එකේ ඇති සියලුම භාණ්ඩ {english_count} ම ඇතුළත් වේ.</div>',
            unsafe_allow_html=True,
        )


def main():
    st.set_page_config(page_title="පික් ලිස්ට් පරිවර්තකය", page_icon="📋", layout="centered")
    st.markdown(f"<style>{CSS}</style>", unsafe_allow_html=True)
    st.markdown(hero_html(), unsafe_allow_html=True)

    steps_slot = st.empty()
    steps_slot.markdown(steps_html(1), unsafe_allow_html=True)

    uploaded_file = st.file_uploader(
        "පරිවර්තනය සඳහා PDF ගොනුවක් තෝරන්න (Select PDF File)", type=["pdf"]
    )
    if uploaded_file is None:
        return

    steps_slot.markdown(steps_html(2), unsafe_allow_html=True)
    with st.spinner("දත්ත විශ්ලේෂණය කරමින් පවතී..."):
        rows, unmatched, docx_bytes, docx_en_bytes, summary = process_pdf(uploaded_file.getvalue(), CODE_VERSION)
    steps_slot.markdown(steps_html(3 if rows else 2), unsafe_allow_html=True)

    matched_count = len(rows)

    if matched_count == 0:
        st.markdown(
            '<div class="status warn">⚠️ දෝෂයකි: අප්ලෝඩ් කරන ලද PDF ගොනුවේ අදාළ වගුව තුළ '
            "කිසිදු භාණ්ඩයක් අපගේ නාමාවලිය සමඟ ගැළපුණේ නැත.</div>",
            unsafe_allow_html=True,
        )
        st.markdown(summary_html(summary), unsafe_allow_html=True)
        st.write("")
        render_downloads(None, docx_en_bytes, summary["english_items"])
        if unmatched:
            st.markdown(unmatched_html(unmatched), unsafe_allow_html=True)
        return

    st.markdown(
        f'<div class="status ok">🎉 සාර්ථකයි! ගැළපෙන භාණ්ඩ පේළි {matched_count} ක් සාර්ථකව පරිවර්තනය කරන ලදී.</div>',
        unsafe_allow_html=True,
    )
    if summary["missing"]:
        st.markdown(
            f'<div class="status warn">⚠️ නාමාවලියේ නැති භාණ්ඩ {summary["missing"]} ක් මගහැරී ඇත. පහත ලැයිස්තුව බලන්න.</div>',
            unsafe_allow_html=True,
        )
    unreadable = sum(1 for r in rows if "?" in (r[COL_CASES], r[COL_PIECES]))
    if unreadable:
        st.markdown(
            f'<div class="status warn">⚠️ ප්‍රමාණ කියවා ගැනීමට නොහැකි වූ පේළි {unreadable} ක් ඇත '
            "(වගුවේ ? ලෙස සලකුණු කර ඇත). කරුණාකර PDF එක සමඟ පරීක්ෂා කරන්න.</div>",
            unsafe_allow_html=True,
        )

    st.markdown(
        '<div class="section-title">📊 සංසන්දන වාර්තාව (Comparison Summary)</div>' + summary_html(summary),
        unsafe_allow_html=True,
    )

    st.write("")
    render_downloads(docx_bytes, docx_en_bytes, summary["english_items"])

    st.markdown(
        '<div class="section-title">🔎 දත්ත පෙරදසුන (Data Preview)</div>' + html_table(pd.DataFrame(rows)),
        unsafe_allow_html=True,
    )

    if unmatched:
        st.markdown(unmatched_html(unmatched), unsafe_allow_html=True)


if __name__ == "__main__":
    main()
