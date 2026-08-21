import xlrd
import os
import re
import xml.etree.ElementTree as ET
from difflib import SequenceMatcher

import requests
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter


# ============================================================
# TTE API - KLJUČEVI SE UZIMAJU IZ GITHUB SECRETS
# ============================================================

TTE_ETAIL_API_KEY = os.environ["TTE_ETAIL_API_KEY"]
TTE_HAVIT_API_KEY = os.environ["TTE_HAVIT_API_KEY"]

ETAIL_URL = (
    "https://tte.rs/api/sr/products"
    "?output_type=xml&api_key="
    + TTE_ETAIL_API_KEY
)

HAVIT_URL = (
    "https://tte.rs/api/sr/products"
    "?output_type=xml&api_key="
    + TTE_HAVIT_API_KEY
)


# ============================================================
# FAJLOVI
# ============================================================

BTM_FILE = "BTM_export.xls"
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

    return re.sub(r"\D", "", value)


def normalize_code(value):
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


def xml_value(product, tag):
    element = product.find(tag)

    if element is None:
        return ""

    return element.text.strip() if element.text else ""


# ============================================================
# DOWNLOAD XML
# ============================================================

def download_xml(url, filename):

    print("=" * 60)
    print("PREUZIMAM SVEŽ XML:", filename)
    print("=" * 60)

    response = requests.get(
        url,
        timeout=120,
        headers={
            "User-Agent": "Mozilla/5.0"
        }
    )

    print("HTTP status:", response.status_code)
    print("Content-Type:", response.headers.get("content-type"))

    response.raise_for_status()

    if not response.content:
        raise Exception("XML je prazan: " + filename)

    with open(filename, "wb") as f:
        f.write(response.content)

    print("Preuzeto bytes:", len(response.content))

    return response.content


# ============================================================
# ČITANJE TTE XML
# ============================================================

def parse_tte_xml(xml_content, source_name):

    print()
    print("=" * 60)
    print("OBRAĐUJEM:", source_name)
    print("=" * 60)

    root = ET.fromstring(xml_content)

    products = []

    for product in root.iter("product"):

        code = normalize_code(
            xml_value(product, "article_number")
        )

        name = xml_value(
            product,
            "name"
        )

        ean = normalize_ean(
            xml_value(product, "ean")
        )

        netto_raw = xml_value(
            product,
            "netto_price"
        )

        try:
            netto_price = float(
                netto_raw.replace(",", ".")
            ) if netto_raw else None
        except:
            netto_price = None

        if not code and not name:
            continue

        products.append({
            "source": source_name,
            "code": code,
            "name": name,
            "ean": ean,
            "netto_price": netto_price
        })

    print("Artikala pronađeno:", len(products))

    return products


# ============================================================
# ČITANJE BTM EXCELA
# ============================================================

def clean_header(value):
    return clean_text(value)


def find_column(headers, possible_names):

    for column_number, header in headers.items():

        normalized = clean_header(header)

        for name in possible_names:

            if normalized == clean_header(name):
                return column_number

    return None


def read_btm_excel(filename):

    print()
    print("=" * 60)
    print("ČITAM BTM:", filename)
    print("=" * 60)

    wb = load_workbook(
        filename,
        data_only=True
    )

    ws = wb.active

    header_row = None

    for row in range(1, min(ws.max_row, 20) + 1):

        values = [
            clean_header(cell.value)
            for cell in ws[row]
            if cell.value is not None
        ]

        if (
            any(v in values for v in ["naziv", "name", "artikal"])
            and
            any(v in values for v in ["cena", "price"])
        ):
            header_row = row
            break

    if header_row is None:
        header_row = 1

    headers = {}

    for cell in ws[header_row]:

        if cell.value is not None:
            headers[cell.column] = cell.value

    name_col = find_column(
        headers,
        [
            "Naziv",
            "Name",
            "Product Name",
            "Artikal"
        ]
    )

    price_col = find_column(
        headers,
        [
            "Cena",
            "Price",
            "Prodajna cena"
        ]
    )

    ean_col = find_column(
        headers,
        [
            "EAN",
            "Barcode",
            "Bar kod"
        ]
    )

    if name_col is None:
        raise Exception(
            "BTM Excel: nije pronađena kolona sa nazivom."
        )

    if price_col is None:
        raise Exception(
            "BTM Excel: nije pronađena kolona sa cenom."
        )

    products = []

    for row in range(
        header_row + 1,
        ws.max_row + 1
    ):

        name = ws.cell(
            row=row,
            column=name_col
        ).value

        if not name:
            continue

        price = ws.cell(
            row=row,
            column=price_col
        ).value

        try:
            price = float(price)
        except:
            continue

        ean = ""

        if ean_col:

            ean = normalize_ean(
                ws.cell(
                    row=row,
                    column=ean_col
                ).value
            )

        products.append({
            "name": str(name).strip(),
            "price": price,
            "ean": ean
        })

    print("BTM artikala:", len(products))

    return products


# ============================================================
# POVEZIVANJE BTM → TTE
# ============================================================

def match_product(btm, tte_products):

    btm_name = btm["name"]
    btm_ean = btm["ean"]

    # --------------------------------------------------------
    # 1. EAN
    # --------------------------------------------------------

    if btm_ean:

        for product in tte_products:

            if product["ean"] == btm_ean:
                return product, 1.0, "EAN"

    # --------------------------------------------------------
    # 2. TTE šifra u BTM nazivu
    # --------------------------------------------------------

    for product in tte_products:

        code = product["code"]

        if code and code in btm_name:
            return product, 0.98, "Šifra"

    # --------------------------------------------------------
    # 3. Naziv
    # --------------------------------------------------------

    best = None
    best_score = 0

    for product in tte_products:

        score = similarity(
            btm_name,
            product["name"]
        )

        if score > best_score:

            best_score = score
            best = product

    if best and best_score >= MIN_MATCH_SCORE:
        return best, best_score, "Naziv"

    return None, 0, ""


# ============================================================
# EXCEL
# ============================================================

def create_excel(rows):

    print()
    print("=" * 60)
    print("PRAVIM REZULTAT")
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
        "Razlika Etail"
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

    # Podaci
    for item in rows:

        ws.append([
            item["tte_code"],
            item["tte_name"],
            item["btm_name"],
            item["btm_price"],
            item["havit_price"],
            item["etail_price"],
            item["havit_difference"],
            item["etail_difference"]
        ])

    # Havit i Etail cene BOLD
    for row in range(
        2,
        ws.max_row + 1
    ):

        ws.cell(
            row=row,
            column=5
        ).font = Font(bold=True)

        ws.cell(
            row=row,
            column=6
        ).font = Font(bold=True)

    # Format cena
    for row in range(
        2,
        ws.max_row + 1
    ):

        for col in range(4, 9):

            ws.cell(
                row=row,
                column=col
            ).number_format = '#,##0'

    # Širina kolona
    widths = {
        1: 14,
        2: 45,
        3: 45,
        4: 14,
        5: 14,
        6: 16,
        7: 16,
        8: 16
    }

    for col, width in widths.items():

        ws.column_dimensions[
            get_column_letter(col)
        ].width = width

    ws.freeze_panes = "A2"

    wb.save(
        OUTPUT_FILE
    )

    print()
    print("==============================================")
    print("GOTOVO")
    print("==============================================")
    print("Fajl:", OUTPUT_FILE)
    print("Broj redova:", len(rows))


# ============================================================
# MAIN
# ============================================================

def main():

    print()
    print("==============================================")
    print("TTE / BTM CENOVNO POREĐENJE")
    print("==============================================")

    # --------------------------------------------------------
    # 1. UVEK PREUZMI SVEŽ TTE XML
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
    # 2. OBRADI XML
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
    # 3. BTM
    # --------------------------------------------------------

    if not os.path.exists(BTM_FILE):

        raise Exception(
            "Nema BTM fajla: " + BTM_FILE
        )

    btm_products = read_btm_excel(
        BTM_FILE
    )

    # --------------------------------------------------------
    # 4. POREĐENJE
    # --------------------------------------------------------

    results = []

    for btm in btm_products:

        havit, havit_score, _ = match_product(
            btm,
            havit_products
        )

        etail, etail_score, _ = match_product(
            btm,
            etail_products
        )

        # ARTIKLE KOJIH NEMA U TTE PONUDI NE UBACUJEMO

        if havit is None and etail is None:
            continue

        # TTE podaci
        tte = havit if havit else etail

        havit_price = (
            havit["netto_price"]
            if havit
            else None
        )

        etail_price = (
            etail["netto_price"]
            if etail
            else None
        )

        havit_difference = None
        etail_difference = None

        if havit_price is not None:

            havit_difference = (
                havit_price -
                btm["price"]
            )

        if etail_price is not None:

            etail_difference = (
                etail_price -
                btm["price"]
            )

        results.append({
            "tte_code": tte["code"],
            "tte_name": tte["name"],
            "btm_name": btm["name"],
            "btm_price": btm["price"],
            "havit_price": havit_price,
            "etail_price": etail_price,
            "havit_difference": havit_difference,
            "etail_difference": etail_difference
        })

    # --------------------------------------------------------
    # 5. SORT
    # --------------------------------------------------------

    results.sort(
        key=lambda x: x["tte_code"]
    )

    print()
    print("TTE/BTM rezultata:", len(results))

    # --------------------------------------------------------
    # 6. EXCEL
    # --------------------------------------------------------

    create_excel(results)


if __name__ == "__main__":
    main()
