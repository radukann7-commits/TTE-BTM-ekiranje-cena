import os
import re
import html
import xlrd
import requests
import xml.etree.ElementTree as ET

from difflib import SequenceMatcher

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
    "?output_type=xml&api_key="
    + str(TTE_ETAIL_API_KEY)
)

HAVIT_URL = (
    "https://tte.rs/api/sr/products"
    "?output_type=xml&api_key="
    + str(TTE_HAVIT_API_KEY)
)


# ============================================================
# POMOĆNE FUNKCIJE
# ============================================================

def clean_text(value):
    if value is None:
        return ""

    value = str(value)
    value = html.unescape(value)

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

    for old, new in replacements.items():
        value = value.replace(old, new)

    value = re.sub(r"[^a-z0-9]+", " ", value)

    return " ".join(value.split())


def normalize_code(value):
    if value is None:
        return ""

    value = str(value).strip()

    if value.endswith(".0"):
        value = value[:-2]

    return re.sub(r"[^0-9A-Za-z]", "", value).upper()


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

    text = str(value).strip()

    if not text:
        return None

    text = text.replace("RSD", "")
    text = text.replace("rsd", "")
    text = text.replace("DIN", "")
    text = text.replace("din", "")
    text = text.replace(" ", "")

    # 13.900 -> 13900
    if re.fullmatch(r"\d{1,3}(\.\d{3})+", text):
        text = text.replace(".", "")

    # 13,900 -> 13900
    elif re.fullmatch(r"\d{1,3}(,\d{3})+", text):
        text = text.replace(",", "")

    else:
        text = text.replace(",", ".")

    try:
        return float(text)
    except Exception:
        return None


def get_xml_value(product, tag):
    element = product.find(tag)

    if element is None:
        return ""

    return clean_text(element.text)


# ============================================================
# PREUZIMANJE TTE XML-A
# ============================================================

def download_xml(url, naziv):

    print("")
    print("=" * 70)
    print("PREUZIMANJE:", naziv)
    print("=" * 70)

    response = requests.get(
        url,
        timeout=120,
        headers={
            "User-Agent": "Mozilla/5.0"
        }
    )

    print("HTTP status:", response.status_code)
    print("Content-Type:", response.headers.get("content-type"))
    print("Veličina:", len(response.content), "bytes")

    response.raise_for_status()

    if not response.content:
        raise Exception(
            f"{naziv}: XML je prazan."
        )

    try:
        ET.fromstring(response.content)
    except Exception as e:
        print("")
        print("PRVI DEO ODGOVORA:")
        print(response.content[:500])

        raise Exception(
            f"{naziv}: odgovor nije ispravan XML. Greška: {e}"
        )

    print(naziv, "XML JE ISPRAVAN.")

    return response.content


# ============================================================
# ČITANJE TTE XML-A
# ============================================================

def parse_tte_xml(xml_data, source_name):

    print("")
    print("=" * 70)
    print("OBRADA:", source_name)
    print("=" * 70)

    root = ET.fromstring(xml_data)

    products = []

    for product in root.findall(".//product"):

        article_number = get_xml_value(
            product,
            "article_number"
        )

        name = get_xml_value(
            product,
            "name"
        )

        brand = get_xml_value(
            product,
            "brand"
        )

        ean = get_xml_value(
            product,
            "ean"
        )

        # TTE NETO CENA
        net_price = get_xml_value(
            product,
            "neto_price"
        )

        # Rezervne opcije ako nema neto_price
        if not net_price:
            net_price = get_xml_value(
                product,
                "net_price"
            )

        if not net_price:
            net_price = get_xml_value(
                product,
                "price"
            )

        price = number_from_text(net_price)

        if not article_number and not name:
            continue

        products.append({
            "article_number": normalize_code(
                article_number
            ),
            "article_number_raw": article_number,
            "name": name,
            "name_norm": normalize_text(name),
            "brand": brand,
            "ean": normalize_ean(ean),
            "price": price,
            "source": source_name,
        })

    print(
        "Pronađeno proizvoda:",
        len(products)
    )

    return products


# ============================================================
# BTM XLS
# ============================================================

def find_btm_header(sheet):

    best_row = None
    best_score = 0

    for row_index in range(
        min(sheet.nrows, 50)
    ):

        values = [
            normalize_text(
                sheet.cell_value(
                    row_index,
                    col
                )
            )
            for col in range(sheet.ncols)
        ]

        row_text = " | ".join(values)

        score = 0

        if "sifra" in row_text:
            score += 3

        if "naziv" in row_text:
            score += 3

        if "ean" in row_text:
            score += 2

        if "cena" in row_text:
            score += 2

        if score > best_score:
            best_score = score
            best_row = row_index

    if best_row is None:
        raise Exception(
            "Ne mogu da pronađem zaglavlje BTM tabele."
        )

    print(
        "BTM header red:",
        best_row + 1
    )

    return best_row


def find_column(headers, possible_names):

    # Prvo tražimo potpuno poklapanje
    for index, header in enumerate(headers):

        normalized = normalize_text(header)

        for possible in possible_names:

            if normalized == normalize_text(
                possible
            ):
                return index

    # Zatim delimično poklapanje
    for index, header in enumerate(headers):

        normalized = normalize_text(header)

        for possible in possible_names:

            possible_norm = normalize_text(
                possible
            )

            if possible_norm in normalized:
                return index

    return None


def read_btm_xls(filename):

    print("")
    print("=" * 70)
    print("ČITANJE BTM XLS")
    print("=" * 70)

    if not os.path.exists(filename):
        raise Exception(
            f"BTM fajl ne postoji: {filename}"
        )

    workbook = xlrd.open_workbook(
        filename,
        formatting_info=False
    )

    print(
        "Sheetovi:",
        workbook.sheet_names()
    )

    sheet = workbook.sheet_by_index(0)

    print(
        "Aktivni sheet:",
        sheet.name
    )

    print(
        "Redova:",
        sheet.nrows
    )

    print(
        "Kolona:",
        sheet.ncols
    )

    header_row = find_btm_header(sheet)

    headers = [
        clean_text(
            sheet.cell_value(
                header_row,
                col
            )
        )
        for col in range(sheet.ncols)
    ]

    print("")
    print("BTM KOLONE:")

    for i, header in enumerate(headers):
        print(i, "=>", header)

    # ŠIFRA
    code_col = find_column(
        headers,
        [
            "Šifra",
            "Sifra",
            "Artikal",
            "Artikal šifra",
            "Article number",
            "Product code",
            "Code"
        ]
    )

    # NAZIV
    name_col = find_column(
        headers,
        [
            "Naziv",
            "Naziv artikla",
            "Artikal naziv",
            "Product name",
            "Name"
        ]
    )

    # EAN
    ean_col = find_column(
        headers,
        [
            "EAN",
            "Barkod",
            "Bar code",
            "Barcode",
            "EAN kod"
        ]
    )

    # CENA
    price_col = find_column(
        headers,
        [
            "Cena",
            "Cena bez PDV",
            "Cena bez PDV-a",
            "VPC",
            "Nabavna cena",
            "B2B cena",
            "Cena RSD",
            "Price"
        ]
    )

    print("")
    print("PRONAĐENE BTM KOLONE:")
    print("Šifra:", code_col)
    print("Naziv:", name_col)
    print("EAN:", ean_col)
    print("Cena:", price_col)

    if name_col is None:
        raise Exception(
            "Ne mogu da pronađem kolonu sa nazivom artikla."
        )

    if price_col is None:
        raise Exception(
            "Ne mogu da pronađem kolonu sa cenom."
        )

    btm_products = []

    for row in range(
        header_row + 1,
        sheet.nrows
    ):

        name = clean_text(
            sheet.cell_value(
                row,
                name_col
            )
        )

        if not name:
            continue

        code = ""

        if code_col is not None:
            code = clean_text(
                sheet.cell_value(
                    row,
                    code_col
                )
            )

        ean = ""

        if ean_col is not None:
            ean = clean_text(
                sheet.cell_value(
                    row,
                    ean_col
                )
            )

        price_raw = sheet.cell_value(
            row,
            price_col
        )

        price = number_from_text(
            price_raw
        )

        if price is None:
            continue

        btm_products.append({
            "code": normalize_code(code),
            "code_raw": code,
            "name": name,
            "name_norm": normalize_text(name),
            "ean": normalize_ean(ean),
            "price": price
        })

    print("")
    print(
        "BTM proizvoda:",
        len(btm_products)
    )

    return btm_products


# ============================================================
# MATCHING PO NAZIVU
# ============================================================

def similarity(a, b):

    if not a or not b:
        return 0

    return SequenceMatcher(
        None,
        normalize_text(a),
        normalize_text(b)
    ).ratio()


def find_tte_match(
    btm,
    tte_products,
    used
):

    # --------------------------------------------------------
    # 1. EAN
    # --------------------------------------------------------

    if btm["ean"]:

        for index, product in enumerate(
            tte_products
        ):

            if index in used:
                continue

            if (
                product["ean"]
                and
                product["ean"] == btm["ean"]
            ):
                return (
                    index,
                    "EAN",
                    1.0
                )

    # --------------------------------------------------------
    # 2. ŠIFRA
    # --------------------------------------------------------

    if btm["code"]:

        for index, product in enumerate(
            tte_products
        ):

            if index in used:
                continue

            if (
                product["article_number"]
                and
                product["article_number"]
                == btm["code"]
            ):
                return (
                    index,
                    "ŠIFRA",
                    1.0
                )

    # --------------------------------------------------------
    # 3. NAZIV
    # --------------------------------------------------------

    btm_name = btm["name_norm"]

    if not btm_name:
        return (
            None,
            None,
            0
        )

    best_index = None
    best_score = 0

    for index, product in enumerate(
        tte_products
    ):

        if index in used:
            continue

        tte_name = product["name_norm"]

        if not tte_name:
            continue

        score = similarity(
            btm_name,
            tte_name
        )

        # Dodatni bonus za model/broj
        btm_numbers = set(
            re.findall(
                r"\b[a-zA-Z]?\d+[a-zA-Z0-9]*\b",
                btm_name
            )
        )

        tte_numbers = set(
            re.findall(
                r"\b[a-zA-Z]?\d+[a-zA-Z0-9]*\b",
                tte_name
            )
        )

        common_numbers = (
            btm_numbers
            .intersection(tte_numbers)
        )

        if common_numbers:
            score += 0.15

        if score > best_score:
            best_score = score
            best_index = index

    # Minimalni prag
    if (
        best_index is not None
        and
        best_score >= 0.70
    ):
        return (
            best_index,
            "NAZIV",
            best_score
        )

    return (
        None,
        None,
        0
    )


# ============================================================
# FORMAT CENE
# ============================================================

def format_price(value):

    if value is None:
        return ""

    if isinstance(value, float):

        if value.is_integer():
            return int(value)

    return value


# ============================================================
# GLAVNI PROGRAM
# ============================================================

def main():

    print("")
    print("=" * 70)
    print("TTE / BTM POREĐENJE")
    print("=" * 70)

    # --------------------------------------------------------
    # PROVERA SECRET-A
    # --------------------------------------------------------

    if not TTE_ETAIL_API_KEY:
        raise Exception(
            "Nedostaje GitHub Secret: TTE_ETAIL_API_KEY"
        )

    if not TTE_HAVIT_API_KEY:
        raise Exception(
            "Nedostaje GitHub Secret: TTE_HAVIT_API_KEY"
        )

    print("TTE API ključevi: OK")

    # --------------------------------------------------------
    # BTM
    # --------------------------------------------------------

    btm_products = read_btm_xls(
        BTM_FILE
    )

    # --------------------------------------------------------
    # SVEŽI TTE XML
    # --------------------------------------------------------

    etail_xml = download_xml(
        ETAIL_URL,
        "ETAIL SPEC"
    )

    havit_xml = download_xml(
        HAVIT_URL,
        "HAVIT"
    )

    # --------------------------------------------------------
    # OBRADA XML-A
    # --------------------------------------------------------

    etail_products = parse_tte_xml(
        etail_xml,
        "ETAIL SPEC"
    )

    havit_products = parse_tte_xml(
        havit_xml,
        "HAVIT"
    )

    # --------------------------------------------------------
    # MATCHING
    # --------------------------------------------------------

    results = []

    used_havit = set()
    used_etail = set()

    stats = {
        "btm": len(btm_products),
        "havit": 0,
        "etail": 0,
        "both": 0,
        "none": 0
    }

    for btm in btm_products:

        # ====================================================
        # HAVIT
        # ====================================================

        havit_index, havit_method, havit_score = (
            find_tte_match(
                btm,
                havit_products,
                used_havit
            )
        )

        havit_match = None

        if havit_index is not None:

            havit_match = (
                havit_products[
                    havit_index
                ]
            )

            used_havit.add(
                havit_index
            )

        # ====================================================
        # ETAIL
        # ====================================================

        etail_index, etail_method, etail_score = (
            find_tte_match(
                btm,
                etail_products,
                used_etail
            )
        )

        etail_match = None

        if etail_index is not None:

            etail_match = (
                etail_products[
                    etail_index
                ]
            )

            used_etail.add(
                etail_index
            )

        # ====================================================
        # AKO ARTIKAL NE POSTOJI U TTE PONUDI
        # NE UBACUJEMO GA
        # ====================================================

        if (
            havit_match is None
            and
            etail_match is None
        ):

            stats["none"] += 1
            continue

        # ====================================================
        # TTE PODACI
        # ====================================================

        tte_product = (
            havit_match
            if havit_match
            else etail_match
        )

        tte_code = (
            tte_product[
                "article_number_raw"
            ]
        )

        tte_name = (
            tte_product["name"]
        )

        # ====================================================
        # CENE
        # ====================================================

        btm_price = btm["price"]

        havit_price = None

        if havit_match:
            havit_price = (
                havit_match["price"]
            )

        etail_price = None

        if etail_match:
            etail_price = (
                etail_match["price"]
            )

        # ====================================================
        # RAZLIKE
        # ====================================================

        havit_difference = None

        if (
            btm_price is not None
            and
            havit_price is not None
        ):

            havit_difference = (
                havit_price
                - btm_price
            )

        etail_difference = None

        if (
            btm_price is not None
            and
            etail_price is not None
        ):

            etail_difference = (
                etail_price
                - btm_price
            )

        # ====================================================
        # STATISTIKA
        # ====================================================

        if havit_match:
            stats["havit"] += 1

        if etail_match:
            stats["etail"] += 1

        if (
            havit_match
            and
            etail_match
        ):
            stats["both"] += 1

        # ====================================================
        # REZULTAT
        # ====================================================

        results.append({
            "tte_code": tte_code,
            "tte_name": tte_name,
            "btm_price": format_price(
                btm_price
            ),
            "havit_price": format_price(
                havit_price
            ),
            "havit_difference": format_price(
                havit_difference
            ),
            "etail_price": format_price(
                etail_price
            ),
            "etail_difference": format_price(
                etail_difference
            )
        })

    # ========================================================
    # SORTIRANJE PO NAZIVU
    # ========================================================

    results.sort(
        key=lambda x: normalize_text(
            x["tte_name"]
        )
    )

    # ========================================================
    # EXCEL
    # ========================================================

    print("")
    print("=" * 70)
    print("PRAVLJENJE EXCELA")
    print("=" * 70)

    workbook = Workbook()

    sheet = workbook.active
    sheet.title = "BTM vs TTE"

    headers = [
        "TTE šifra",
        "TTE naziv",
        "BTM cena",
        "HAVIT cenovnik",
        "Razlika HAVIT",
        "ETAIL spec",
        "Razlika ETAIL"
    ]

    sheet.append(headers)

    # --------------------------------------------------------
    # HEADER
    # --------------------------------------------------------

    for cell in sheet[1]:

        cell.font = Font(
            bold=True
        )

        cell.alignment = Alignment(
            horizontal="center",
            vertical="center"
        )

    # --------------------------------------------------------
    # PODACI
    # --------------------------------------------------------

    for item in results:

        sheet.append([
            item["tte_code"],
            item["tte_name"],
            item["btm_price"],
            item["havit_price"],
            item["havit_difference"],
            item["etail_price"],
            item["etail_difference"]
        ])

    # --------------------------------------------------------
    # BOLD HAVIT I ETAIL CENE
    # --------------------------------------------------------

    for row in range(
        2,
        sheet.max_row + 1
    ):

        # HAVIT
        sheet.cell(
            row=row,
            column=4
        ).font = Font(
            bold=True
        )

        # ETAIL
        sheet.cell(
            row=row,
            column=6
        ).font = Font(
            bold=True
        )

    # --------------------------------------------------------
    # FORMAT BROJEVA
    # --------------------------------------------------------

    for row in range(
        2,
        sheet.max_row + 1
    ):

        for col in [
            3,
            4,
            5,
            6,
            7
        ]:

            cell = sheet.cell(
                row=row,
                column=col
            )

            if isinstance(
                cell.value,
                (int, float)
            ):

                cell.number_format = "#,##0"

    # --------------------------------------------------------
    # PORAVNANJE
    # --------------------------------------------------------

    for row in sheet.iter_rows():

        for cell in row:

            cell.alignment = Alignment(
                vertical="center"
            )

    # --------------------------------------------------------
    # ŠIRINE KOLONA
    # --------------------------------------------------------

    widths = {
        1: 15,
        2: 55,
        3: 15,
        4: 18,
        5: 18,
        6: 18,
        7: 18
    }

    for column, width in widths.items():

        sheet.column_dimensions[
            get_column_letter(column)
        ].width = width

    # --------------------------------------------------------
    # FILTER
    # --------------------------------------------------------

    sheet.auto_filter.ref = (
        sheet.dimensions
    )

    # --------------------------------------------------------
    # FREEZE HEADER
    # --------------------------------------------------------

    sheet.freeze_panes = "A2"

    # --------------------------------------------------------
    # SAČUVAJ EXCEL
    # --------------------------------------------------------

    workbook.save(
        OUTPUT_FILE
    )

    # ========================================================
    # ISPIS
    # ========================================================

    print("")
    print("=" * 70)
    print("GOTOVO")
    print("=" * 70)

    print(
        "BTM proizvoda:",
        stats["btm"]
    )

    print(
        "HAVIT pronađeno:",
        stats["havit"]
    )

    print(
        "ETAIL pronađeno:",
        stats["etail"]
    )

    print(
        "U oba XML-a:",
        stats["both"]
    )

    print(
        "Izbačeno - nema u TTE:",
        stats["none"]
    )

    print(
        "Konačan broj redova:",
        len(results)
    )

    print("")
    print(
        "Excel:",
        OUTPUT_FILE
    )

    if not results:
        raise Exception(
            "Nije pronađen nijedan proizvod koji postoji u TTE XML-u."
        )


# ============================================================
# START
# ============================================================

if __name__ == "__main__":
    main()
