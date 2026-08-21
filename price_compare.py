import os
import re
import xml.etree.ElementTree as ET
from difflib import SequenceMatcher

import requests
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter


# ============================================================
# TTE API
# ============================================================

ETAIL_URL = (
    "https://tte.rs/api/sr/products"
    "?output_type=xml"
    "&api_key=1GQTUjZo7rlFIlsxMyidtVyh"
)

HAVIT_URL = (
    "https://tte.rs/api/sr/products"
    "?output_type=xml"
    "&api_key=jnak1b82M19JEf6jMkz6SGxo"
)


# ============================================================
# PODEŠAVANJA
# ============================================================

BTM_FILE = "BTM_export.xlsx"
OUTPUT_FILE = "BTM_TTE_poredjenje.xlsx"

MIN_MATCH_SCORE = 0.72


# ============================================================
# POMOĆNE FUNKCIJE
# ============================================================

def clean_text(value):
    if value is None:
        return ""

    value = str(value).strip().lower()

    replacements = {
        "č": "c",
        "ć": "c",
        "š": "s",
        "đ": "d",
        "ž": "z",
    }

    for old, new in replacements.items():
        value = value.replace(old, new)

    value = re.sub(r"[^a-z0-9]+", " ", value)
    value = re.sub(r"\s+", " ", value).strip()

    return value


def normalize_ean(value):
    if value is None:
        return ""

    value = str(value).strip()

    if value.endswith(".0"):
        value = value[:-2]

    value = re.sub(r"\D", "", value)

    return value


def normalize_article_number(value):
    if value is None:
        return ""

    value = str(value).strip()

    if value.endswith(".0"):
        value = value[:-2]

    return value


def similarity(a, b):
    a = clean_text(a)
    b = clean_text(b)

    if not a or not b:
        return 0

    return SequenceMatcher(None, a, b).ratio()


def get_xml_value(product, tag):
    element = product.find(tag)

    if element is None:
        return ""

    return element.text.strip() if element.text else ""


# ============================================================
# PREUZIMANJE XML-a
# ============================================================

def download_xml(url, filename):
    print("=" * 60)
    print("PREUZIMAM SVEŽ XML")
    print(filename)
    print("=" * 60)

    response = requests.get(
        url,
        timeout=120,
        headers={
            "User-Agent": "Mozilla/5.0"
        }
    )

    response.raise_for_status()

    content = response.content

    if not content:
        raise Exception(f"XML je prazan: {filename}")

    with open(filename, "wb") as f:
        f.write(content)

    print("Preuzet:", len(content), "bytes")

    return content


# ============================================================
# ČITANJE TTE XML-a
# ============================================================

def parse_tte_xml(xml_content, source_name):
    print()
    print("=" * 60)
    print("OBRAĐUJEM", source_name)
    print("=" * 60)

    root = ET.fromstring(xml_content)

    products = []

    for product in root.iter("product"):

        article_number = normalize_article_number(
            get_xml_value(product, "article_number")
        )

        name = get_xml_value(product, "name")

        ean = normalize_ean(
            get_xml_value(product, "ean")
        )

        netto_price_raw = get_xml_value(
            product,
            "netto_price"
        )

        price_raw = get_xml_value(
            product,
            "price"
        )

        try:
            netto_price = float(
                netto_price_raw.replace(",", ".")
            ) if netto_price_raw else None
        except:
            netto_price = None

        try:
            price = float(
                price_raw.replace(",", ".")
            ) if price_raw else None
        except:
            price = None

        if not article_number and not name:
            continue

        products.append({
            "source": source_name,
            "article_number": article_number,
            "name": name,
            "ean": ean,
            "netto_price": netto_price,
            "price": price,
        })

    print("Pronađeno artikala:", len(products))

    return products


# ============================================================
# ČITANJE BTM EXCELA
# ============================================================

def find_header_row(ws):
    possible_headers = {
        "ean",
        "naziv",
        "name",
        "sifra",
        "šifra",
        "cena",
        "price",
    }

    for row in ws.iter_rows(min_row=1, max_row=min(20, ws.max_row)):

        values = []

        for cell in row:
            if cell.value is not None:
                values.append(clean_text(cell.value))

        matches = sum(
            1 for value in values
            if value in possible_headers
        )

        if matches >= 2:
            return row[0].row

    return 1


def find_column(headers, names):

    normalized_headers = {
        clean_text(header): index
        for index, header in headers.items()
        if header is not None
    }

    for name in names:
        key = clean_text(name)

        if key in normalized_headers:
            return normalized_headers[key]

    return None


def read_btm_excel(filename):

    print()
    print("=" * 60)
    print("ČITAM BTM EXPORT")
    print("=" * 60)

    wb = load_workbook(
        filename,
        data_only=True
    )

    ws = wb.active

    header_row = find_header_row(ws)

    headers = {}

    for cell in ws[header_row]:
        if cell.value is not None:
            headers[cell.column] = cell.value

    ean_col = find_column(
        headers,
        ["EAN", "BARCODE", "BAR KOD"]
    )

    name_col = find_column(
        headers,
        ["Naziv", "Name", "Product Name", "Artikal"]
    )

    price_col = find_column(
        headers,
        ["Cena", "Price", "Prodajna cena"]
    )

    btm_code_col = find_column(
        headers,
        ["Šifra", "Sifra", "Code", "SKU"]
    )

    if name_col is None:
        raise Exception(
            "Ne mogu da pronađem kolonu sa nazivom artikla u BTM Excelu."
        )

    if price_col is None:
        raise Exception(
            "Ne mogu da pronađem kolonu sa cenom u BTM Excelu."
        )

    btm_products = []

    for row in range(header_row + 1, ws.max_row + 1):

        name = ws.cell(
            row=row,
            column=name_col
        ).value

        if not name:
            continue

        ean = ""

        if ean_col:
            ean = normalize_ean(
                ws.cell(
                    row=row,
                    column=ean_col
                ).value
            )

        price = ws.cell(
            row=row,
            column=price_col
        ).value

        try:
            price = float(price)
        except:
            continue

        btm_code = ""

        if btm_code_col:
            btm_code = normalize_article_number(
                ws.cell(
                    row=row,
                    column=btm_code_col
                ).value
            )

        btm_products.append({
            "btm_name": str(name).strip(),
            "btm_ean": ean,
            "btm_price": price,
            "btm_code": btm_code,
        })

    print("BTM artikala:", len(btm_products))

    return btm_products


# ============================================================
# POVEZIVANJE ARTIKALA
# ============================================================

def match_product(btm, tte_products):

    btm_ean = btm["btm_ean"]
    btm_name = btm["btm_name"]

    # --------------------------------------------------------
    # 1. EAN - najjače povezivanje
    # --------------------------------------------------------

    if btm_ean:

        for product in tte_products:

            if product["ean"] == btm_ean:
                return product, 1.0, "EAN"

    # --------------------------------------------------------
    # 2. TTE šifra ako se nalazi u BTM nazivu
    # --------------------------------------------------------

    for product in tte_products:

        code = product["article_number"]

        if not code:
            continue

        if code in btm_name:
            return product, 0.98, "TTE šifra"

    # --------------------------------------------------------
    # 3. Naziv
    # --------------------------------------------------------

    best_product = None
    best_score = 0

    for product in tte_products:

        score = similarity(
            btm_name,
            product["name"]
        )

        if score > best_score:
            best_score = score
            best_product = product

    if best_product and best_score >= MIN_MATCH_SCORE:
        return best_product, best_score, "Naziv"

    return None, 0, ""


# ============================================================
# PRAVLJENJE EXCELA
# ============================================================

def create_excel(rows):

    print()
    print("=" * 60)
    print("PRAVIM KONAČAN EXCEL")
    print("=" * 60)

    wb = Workbook()

    ws = wb.active
    ws.title = "Poredjenje"

    headers = [
        "TTE šifra",
        "TTE naziv",
        "BTM naziv",
        "BTM cena",
        "Havit cena",
        "Etail spec cena",
        "Razlika Havit",
        "Razlika Etail",
    ]

    ws.append(headers)

    # Zaglavlje
    for cell in ws[1]:
        cell.font = Font(
            bold=True
        )
        cell.alignment = Alignment(
            horizontal="center"
        )

    for row in rows:

        ws.append([
            row["tte_code"],
            row["tte_name"],
            row["btm_name"],
            row["btm_price"],
            row["havit_price"],
            row["etail_price"],
            row["havit_difference"],
            row["etail_difference"],
        ])

    # --------------------------------------------------------
    # BOLD Havit i Etail CENE
    # --------------------------------------------------------

    for row in range(2, ws.max_row + 1):

        ws.cell(
            row=row,
            column=5
        ).font = Font(bold=True)

        ws.cell(
            row=row,
            column=6
        ).font = Font(bold=True)

    # --------------------------------------------------------
    # Format cena
    # --------------------------------------------------------

    for row in range(2, ws.max_row + 1):

        for col in range(4, 9):

            ws.cell(
                row=row,
                column=col
            ).number_format = '#,##0'

    # --------------------------------------------------------
    # Auto širina
    # --------------------------------------------------------

    for column in ws.columns:

        max_length = 0

        for cell in column:

            if cell.value is not None:

                length = len(
                    str(cell.value)
                )

                if length > max_length:
                    max_length = length

        width = min(
            max_length + 2,
            60
        )

        ws.column_dimensions[
            get_column_letter(
                column[0].column
            )
        ].width = width

    ws.freeze_panes = "A2"

    wb.save(OUTPUT_FILE)

    print()
    print("GOTOVO:")
    print(OUTPUT_FILE)


# ============================================================
# GLAVNI PROGRAM
# ============================================================

def main():

    print()
    print("==============================================")
    print("TTE + BTM POREĐENJE")
    print("==============================================")

    # --------------------------------------------------------
    # Sveži XML
    # --------------------------------------------------------

    etail_xml = download_xml(
        ETAIL_URL,
        "Etail_spec.xml"
    )

    havit_xml = download_xml(
        HAVIT_URL,
        "Havit.xml"
    )

    # --------------------------------------------------------
    # Obrada XML-a
    # --------------------------------------------------------

    etail_products = parse_tte_xml(
        etail_xml,
        "Etail spec"
    )

    havit_products = parse_tte_xml(
        havit_xml,
        "Havit"
    )

    # --------------------------------------------------------
    # BTM
    # --------------------------------------------------------

    if not os.path.exists(BTM_FILE):

        raise Exception(
            f"BTM fajl nije pronađen: {BTM_FILE}"
        )

    btm_products = read_btm_excel(
        BTM_FILE
    )

    # --------------------------------------------------------
    # SVI TTE proizvodi
    # --------------------------------------------------------

    all_tte = {}

    for product in etail_products:

        code = product["article_number"]

        if code:
            all_tte.setdefault(
                code,
                {
                    "code": code,
                    "name": product["name"],
                    "etail": product
                }
            )

    for product in havit_products:

        code = product["article_number"]

        if code:

            if code not in all_tte:

                all_tte[code] = {
                    "code": code,
                    "name": product["name"],
                    "etail": None
                }

            all_tte[code]["havit"] = product

    # --------------------------------------------------------
    # BTM → TTE
    # --------------------------------------------------------

    results = []

    for btm in btm_products:

        havit_match, havit_score, havit_method = match_product(
            btm,
            havit_products
        )

        etail_match, etail_score, etail_method = match_product(
            btm,
            etail_products
        )

        # Ako nije pronađen ni u jednom TTE XML-u,
        # NE UBACUJEMO ga u rezultat.

        if not havit_match and not etail_match:
            continue

        # Prioritet naziva
        tte_product = (
            havit_match
            if havit_match
            else etail_match
        )

        tte_code = tte_product["article_number"]
        tte_name = tte_product["name"]

        havit_price = (
            havit_match["netto_price"]
            if havit_match
            else None
        )

        etail_price = (
            etail_match["netto_price"]
            if etail_match
            else None
        )

        havit_difference = None
        etail_difference = None

        if havit_price is not None:
            havit_difference = (
                havit_price -
                btm["btm_price"]
            )

        if etail_price is not None:
            etail_difference = (
                etail_price -
                btm["btm_price"]
            )

        results.append({
            "tte_code": tte_code,
            "tte_name": tte_name,
            "btm_name": btm["btm_name"],
            "btm_price": btm["btm_price"],
            "havit_price": havit_price,
            "etail_price": etail_price,
            "havit_difference": havit_difference,
            "etail_difference": etail_difference,
        })

    # --------------------------------------------------------
    # Sortiranje po TTE šifri
    # --------------------------------------------------------

    results.sort(
        key=lambda x: x["tte_code"]
    )

    print()
    print("Ukupno rezultata:", len(results))

    create_excel(results)


if __name__ == "__main__":
    main()
