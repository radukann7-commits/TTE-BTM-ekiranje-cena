import os
import html
import re
import xlrd
import requests
import xml.etree.ElementTree as ET

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter


# ============================================================
# PODEŠAVANJA
# ============================================================

BTM_FILE = "BTM_export.xls"
OUTPUT_FILE = "BTM_TTE_poredjenje.xlsx"

TTE_ETAIL_API_KEY = os.environ.get("TTE_ETAIL_API_KEY")
TTE_HAVIT_API_KEY = os.environ.get("TTE_HAVIT_API_KEY")

ETAIL_URL = (
    "https://tte.rs/api/sr/products"
    "?output_type=xml&api_key=" + str(TTE_ETAIL_API_KEY)
)

HAVIT_URL = (
    "https://tte.rs/api/sr/products"
    "?output_type=xml&api_key=" + str(TTE_HAVIT_API_KEY)
)

# Ovi brendovi se izbacuju iz poređenja.
# JBL, Samsung i Apple OSTAJU.
EXCLUDED_BRANDS = {
    "bavin",
    "bavitel",
    "havit",
    "powerology",
    "green lion",
    "porodo",
}


# ============================================================
# POMOĆNE FUNKCIJE
# ============================================================

def clean_text(value):
    if value is None:
        return ""

    value = html.unescape(str(value))
    value = re.sub(r"<[^>]+>", " ", value)
    value = value.replace("\xa0", " ")
    value = value.replace("\n", " ")
    value = value.replace("\r", " ")
    value = value.replace("\t", " ")
    value = re.sub(r"\s+", " ", value)
    return value.strip()


def normalize_text(value):
    value = clean_text(value).lower()
    replacements = {
        "č": "c",
        "ć": "c",
        "š": "s",
        "ž": "z",
        "đ": "d",
    }
    for a, b in replacements.items():
        value = value.replace(a, b)
    return " ".join(re.sub(r"[^a-z0-9]+", " ", value).split())


def normalize_ean(value):
    if value is None:
        return ""

    value = str(value).strip()
    if value.endswith(".0"):
        value = value[:-2]

    return re.sub(r"\D", "", value)


def number_from_text(value):
    if value is None:
        return None

    if isinstance(value, (int, float)):
        return float(value)

    text = clean_text(value)
    if not text:
        return None

    text = (
        text.replace("RSD", "")
        .replace("rsd", "")
        .replace("DIN", "")
        .replace("din", "")
        .replace(" ", "")
    )

    if re.fullmatch(r"\d{1,3}(\.\d{3})+", text):
        text = text.replace(".", "")
    elif re.fullmatch(r"\d{1,3}(,\d{3})+", text):
        text = text.replace(",", "")
    else:
        text = text.replace(",", ".")

    try:
        return float(text)
    except Exception:
        return None


def format_price(value):
    if value is None:
        return ""
    value = round(float(value), 2)
    return int(value) if value.is_integer() else value


def get_xml_value(product, tag):
    element = product.find(tag)
    if element is None:
        return ""
    return clean_text(element.text)


def find_column(headers, possible_names):
    normalized_headers = [normalize_text(h) for h in headers]
    normalized_names = [normalize_text(x) for x in possible_names]

    for index, header in enumerate(normalized_headers):
        if header in normalized_names:
            return index

    for index, header in enumerate(normalized_headers):
        if any(name in header for name in normalized_names):
            return index

    return None


# ============================================================
# BTM
# ============================================================

def find_btm_header(sheet):
    best_row = None
    best_score = 0

    for row_index in range(min(sheet.nrows, 50)):
        values = [
            clean_text(sheet.cell_value(row_index, col)).lower()
            for col in range(sheet.ncols)
        ]
        row_text = " | ".join(values)

        score = 0
        if "sifra" in row_text or "šifra" in row_text:
            score += 3
        if "naziv" in row_text:
            score += 3
        if "ean" in row_text or "barkod" in row_text:
            score += 3
        if "cena" in row_text:
            score += 2

        if score > best_score:
            best_score = score
            best_row = row_index

    if best_row is None:
        raise Exception("Ne mogu da pronađem zaglavlje BTM tabele.")

    return best_row


def read_btm_xls(filename):
    print("=" * 70)
    print("ČITANJE BTM XLS")
    print("=" * 70)

    if not os.path.exists(filename):
        raise Exception(f"BTM fajl ne postoji: {filename}")

    workbook = xlrd.open_workbook(filename, formatting_info=False)
    sheet = workbook.sheet_by_index(0)
    header_row = find_btm_header(sheet)

    headers = [
        clean_text(sheet.cell_value(header_row, col))
        for col in range(sheet.ncols)
    ]

    ean_col = find_column(
        headers,
        ["EAN", "Barkod", "Bar code", "Barcode", "EAN kod"]
    )
    name_col = find_column(
        headers,
        ["Naziv", "Naziv artikla", "Artikal naziv", "Product name", "Name"]
    )
    price_col = find_column(
        headers,
        ["Cena", "Cena RSD", "Cena bez PDV", "Cena NETO", "Neto cena", "Price"]
    )
    code_col = find_column(
        headers,
        ["Šifra", "Sifra", "Artikal šifra", "Product code", "Code"]
    )

    if ean_col is None:
        raise Exception("Ne mogu da pronađem BTM EAN/Barkod kolonu.")
    if name_col is None:
        raise Exception("Ne mogu da pronađem BTM kolonu sa nazivom artikla.")
    if price_col is None:
        raise Exception("Ne mogu da pronađem BTM kolonu sa cenom.")

    products = []

    for row in range(header_row + 1, sheet.nrows):
        ean = normalize_ean(sheet.cell_value(row, ean_col))
        name = clean_text(sheet.cell_value(row, name_col))
        price = number_from_text(sheet.cell_value(row, price_col))
        code = (
            clean_text(sheet.cell_value(row, code_col))
            if code_col is not None
            else ""
        )

        # Bez EAN-a ne ulazi u automatsko poređenje.
        if not ean or not name or price is None:
            continue

        # BTM cena se tretira kao NETO cena, bez PDV-a.
        products.append({
            "ean": ean,
            "name": name,
            "code": code,
            "net_price": price,
        })

    print("BTM artikala sa EAN-om:", len(products))
    return products


# ============================================================
# TTE API
# ============================================================

def download_xml(url, source_name):
    print("=" * 70)
    print("PREUZIMANJE:", source_name)
    print("=" * 70)

    response = requests.get(
        url,
        timeout=120,
        headers={"User-Agent": "Mozilla/5.0"}
    )
    response.raise_for_status()

    if not response.content:
        raise Exception(f"{source_name}: XML je prazan.")

    try:
        ET.fromstring(response.content)
    except Exception as e:
        print(response.content[:500])
        raise Exception(f"{source_name}: odgovor nije ispravan XML: {e}")

    return response.content


def parse_tte_xml(xml_data, source_name):
    root = ET.fromstring(xml_data)
    products = {}

    for product in root.findall(".//product"):
        ean = normalize_ean(get_xml_value(product, "ean"))
        if not ean:
            continue

        brand = get_xml_value(product, "brand")
        if normalize_text(brand) in EXCLUDED_BRANDS:
            continue

        price_text = get_xml_value(product, "neto_price")
        if not price_text:
            price_text = get_xml_value(product, "net_price")
        if not price_text:
            price_text = get_xml_value(product, "price")

        net_price = number_from_text(price_text)
        if net_price is None:
            continue

        if ean not in products:
            products[ean] = {
                "ean": ean,
                "article_number": get_xml_value(product, "article_number"),
                "name": get_xml_value(product, "name"),
                "brand": brand,
                "net_price": net_price,
            }

    print(source_name, "artikala sa EAN-om posle filtera:", len(products))
    return products


# ============================================================
# EXCEL REZULTAT
# ============================================================

def main():
    print("")
    print("=" * 70)
    print("TTE / BTM POREĐENJE - FINAL")
    print("=" * 70)
    print("POVEZIVANJE: ISKLJUČIVO PO EAN-u")
    print("CENE: BTM NETO + TTE NETO")
    print("IZLAZ: SAMO ARTIKLI SA RAZLIKOM")
    print("IZBAČENI BRENDOVI:", ", ".join(sorted(EXCLUDED_BRANDS)))

    if not TTE_ETAIL_API_KEY:
        raise Exception("Nedostaje GitHub Secret: TTE_ETAIL_API_KEY")
    if not TTE_HAVIT_API_KEY:
        raise Exception("Nedostaje GitHub Secret: TTE_HAVIT_API_KEY")

    btm_products = read_btm_xls(BTM_FILE)

    etail_xml = download_xml(ETAIL_URL, "ETAIL")
    havit_xml = download_xml(HAVIT_URL, "HAVIT")

    etail_products = parse_tte_xml(etail_xml, "ETAIL")
    havit_products = parse_tte_xml(havit_xml, "HAVIT")

    results = []
    matched_etail = 0
    matched_havit = 0

    for btm in btm_products:
        ean = btm["ean"]
        etail = etail_products.get(ean)
        havit = havit_products.get(ean)

        if etail:
            matched_etail += 1
        if havit:
            matched_havit += 1

        if not etail and not havit:
            continue

        btm_net = btm["net_price"]

        etail_net = etail["net_price"] if etail else None
        havit_net = havit["net_price"] if havit else None

        etail_diff = (
            round(etail_net - btm_net, 2)
            if etail_net is not None else None
        )
        havit_diff = (
            round(havit_net - btm_net, 2)
            if havit_net is not None else None
        )

        # Ako je cena ista u oba postojeća cenovna ranga, ne prikazuj artikal.
        has_difference = (
            (etail_diff is not None and abs(etail_diff) >= 0.01)
            or
            (havit_diff is not None and abs(havit_diff) >= 0.01)
        )

        if not has_difference:
            continue

        results.append([
            ean,
            btm["code"],
            btm["name"],
            format_price(btm_net),
            format_price(havit_net),
            format_price(havit_diff),
            format_price(etail_net),
            format_price(etail_diff),
            havit["article_number"] if havit else (etail["article_number"] if etail else ""),
            havit["name"] if havit else (etail["name"] if etail else ""),
            havit["brand"] if havit else (etail["brand"] if etail else ""),
        ])

    # Najveće apsolutne razlike prve.
    results.sort(
        key=lambda row: max(
            abs(float(row[5])) if row[5] != "" else 0,
            abs(float(row[7])) if row[7] != "" else 0,
        ),
        reverse=True,
    )

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "BTM vs TTE"

    headers = [
        "EAN",
        "BTM šifra",
        "BTM naziv",
        "BTM cena NETO",
        "HAVIT cena NETO",
        "Razlika HAVIT",
        "ETAIL cena NETO",
        "Razlika ETAIL",
        "TTE šifra",
        "TTE naziv",
        "Brend",
    ]

    sheet.append(headers)

    for cell in sheet[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for row in results:
        sheet.append(row)

    for row in sheet.iter_rows(min_row=2):
        for col in [4, 5, 6, 7, 8]:
            cell = row[col - 1]
            if isinstance(cell.value, (int, float)):
                cell.number_format = "#,##0.00"

    widths = {
        1: 16,
        2: 16,
        3: 50,
        4: 18,
        5: 18,
        6: 16,
        7: 18,
        8: 16,
        9: 18,
        10: 50,
        11: 18,
    }

    for col, width in widths.items():
        sheet.column_dimensions[get_column_letter(col)].width = width

    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions

    workbook.save(OUTPUT_FILE)

    print("")
    print("=" * 70)
    print("ZAVRŠENO")
    print("=" * 70)
    print("BTM artikala:", len(btm_products))
    print("ETAIL poklapanja po EAN-u:", matched_etail)
    print("HAVIT poklapanja po EAN-u:", matched_havit)
    print("ARTIKALA SA RAZLIKOM:", len(results))
    print("IZLAZ:", OUTPUT_FILE)


if __name__ == "__main__":
    main()
