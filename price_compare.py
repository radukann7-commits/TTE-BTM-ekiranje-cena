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
PDV = 1.20

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


def normalize_ean(value):
    """EAN se koristi ISKLJUČIVO za povezivanje artikala."""
    if value is None:
        return ""

    value = str(value).strip()

    # Excel/BTM može dati EAN kao 13-cifren broj sa .0
    if value.endswith(".0"):
        value = value[:-2]

    # Ako je EAN već decimalno zapisan kao float, sačuvaj samo cifre.
    ean = re.sub(r"\D", "", value)

    return ean


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

    # 13.900 = 13900
    if re.fullmatch(r"\d{1,3}(\.\d{3})+", text):
        text = text.replace(".", "")

    # 13,900 = 13900
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

    if value.is_integer():
        return int(value)

    return value


def get_xml_value(product, tag):
    element = product.find(tag)
    if element is None:
        return ""
    return clean_text(element.text)


# ============================================================
# TTE XML
# ============================================================

def download_xml(url, naziv):
    print("")
    print("=" * 70)
    print("PREUZIMANJE:", naziv)
    print("=" * 70)

    response = requests.get(
        url,
        timeout=120,
        headers={"User-Agent": "Mozilla/5.0"}
    )

    print("HTTP status:", response.status_code)
    print("Content-Type:", response.headers.get("content-type"))
    print("Veličina:", len(response.content), "bytes")

    response.raise_for_status()

    if not response.content:
        raise Exception(f"{naziv}: XML je prazan.")

    try:
        ET.fromstring(response.content)
    except Exception as e:
        print("PRVI DEO ODGOVORA:")
        print(response.content[:500])
        raise Exception(
            f"{naziv}: odgovor nije ispravan XML. Greška: {e}"
        )

    print(naziv, "XML JE ISPRAVAN.")
    return response.content


def parse_tte_xml(xml_data, source_name):
    print("")
    print("=" * 70)
    print("OBRADA:", source_name)
    print("=" * 70)

    root = ET.fromstring(xml_data)
    products = {}

    for product in root.findall(".//product"):
        ean = normalize_ean(get_xml_value(product, "ean"))

        # Bez EAN-a nema automatskog povezivanja.
        if not ean:
            continue

        article_number = get_xml_value(product, "article_number")
        name = get_xml_value(product, "name")
        brand = get_xml_value(product, "brand")

        # TTE API daje NETO cenu.
        net_price_text = get_xml_value(product, "neto_price")

        if not net_price_text:
            net_price_text = get_xml_value(product, "net_price")

        if not net_price_text:
            net_price_text = get_xml_value(product, "price")

        net_price = number_from_text(net_price_text)

        # Ako isti EAN postoji više puta, zadržavamo prvi zapis.
        if ean not in products:
            products[ean] = {
                "ean": ean,
                "article_number": article_number,
                "name": name,
                "brand": brand,
                "net_price": net_price,
                "price_with_vat": (
                    net_price * PDV
                    if net_price is not None
                    else None
                ),
                "source": source_name,
            }

    print("Proizvoda sa EAN-om:", len(products))
    return products


# ============================================================
# BTM XLS
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

    print("BTM header red:", best_row + 1)
    return best_row


def find_column(headers, possible_names):
    normalized_headers = [clean_text(h).lower() for h in headers]

    for index, header in enumerate(normalized_headers):
        for possible in possible_names:
            if header == possible.lower():
                return index

    for index, header in enumerate(normalized_headers):
        for possible in possible_names:
            if possible.lower() in header:
                return index

    return None


def read_btm_xls(filename):
    print("")
    print("=" * 70)
    print("ČITANJE BTM XLS")
    print("=" * 70)

    if not os.path.exists(filename):
        raise Exception(f"BTM fajl ne postoji: {filename}")

    workbook = xlrd.open_workbook(filename, formatting_info=False)
    sheet = workbook.sheet_by_index(0)

    print("Sheetovi:", workbook.sheet_names())
    print("Aktivni sheet:", sheet.name)
    print("Redova:", sheet.nrows)
    print("Kolona:", sheet.ncols)

    header_row = find_btm_header(sheet)

    headers = [
        clean_text(sheet.cell_value(header_row, col))
        for col in range(sheet.ncols)
    ]

    print("BTM KOLONE:")
    for i, header in enumerate(headers):
        print(i, "=>", header)

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
        ["Cena", "Cena RSD", "Cena sa PDV", "Price"]
    )

    code_col = find_column(
        headers,
        ["Šifra", "Sifra", "Artikal", "Artikal šifra", "Product code", "Code"]
    )

    print("")
    print("PRONAĐENE BTM KOLONE:")
    print("EAN:", ean_col)
    print("Naziv:", name_col)
    print("Cena:", price_col)
    print("Šifra:", code_col)

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

        if not name:
            continue

        products.append({
            "ean": ean,
            "name": name,
            "price_with_vat": price,
            "code": (
                clean_text(sheet.cell_value(row, code_col))
                if code_col is not None
                else ""
            )
        })

    print("BTM proizvoda:", len(products))
    print("BTM proizvoda sa EAN-om:", sum(1 for p in products if p["ean"]))
    return products


# ============================================================
# GLAVNI PROGRAM
# ============================================================

def main():
    print("")
    print("=" * 70)
    print("TTE / BTM POREĐENJE - SAMO EAN")
    print("=" * 70)

    if not TTE_ETAIL_API_KEY:
        raise Exception("Nedostaje GitHub Secret: TTE_ETAIL_API_KEY")

    if not TTE_HAVIT_API_KEY:
        raise Exception("Nedostaje GitHub Secret: TTE_HAVIT_API_KEY")

    print("TTE API ključevi: OK")
    print("PDV:", int((PDV - 1) * 100), "%")
    print("POKLOPAVANJE: ISKLJUČIVO PO EAN-u")

    # --------------------------------------------------------
    # BTM
    # --------------------------------------------------------
    btm_products = read_btm_xls(BTM_FILE)

    # --------------------------------------------------------
    # TTE XML
    # --------------------------------------------------------
    etail_xml = download_xml(ETAIL_URL, "ETAIL SPEC")
    havit_xml = download_xml(HAVIT_URL, "HAVIT")

    etail_products = parse_tte_xml(etail_xml, "ETAIL SPEC")
    havit_products = parse_tte_xml(havit_xml, "HAVIT")

    # --------------------------------------------------------
    # POREĐENJE
    # --------------------------------------------------------
    results = []

    matched_havit = 0
    matched_etail = 0
    unmatched = 0

    for btm in btm_products:
        ean = btm["ean"]

        havit = havit_products.get(ean) if ean else None
        etail = etail_products.get(ean) if ean else None

        if havit:
            matched_havit += 1

        if etail:
            matched_etail += 1

        if not havit and not etail:
            unmatched += 1

        # U rezultat ulaze svi BTM artikli.
        # Ako nema EAN ili nema poklapanja, TTE podaci ostaju prazni.
        tte = havit or etail

        btm_price = btm["price_with_vat"]

        havit_net = havit["net_price"] if havit else None
        havit_vat = havit["price_with_vat"] if havit else None

        etail_net = etail["net_price"] if etail else None
        etail_vat = etail["price_with_vat"] if etail else None

        havit_difference = (
            havit_vat - btm_price
            if havit_vat is not None and btm_price is not None
            else None
        )

        etail_difference = (
            etail_vat - btm_price
            if etail_vat is not None and btm_price is not None
            else None
        )

        results.append({
            "ean": ean,
            "btm_code": btm["code"],
            "btm_name": btm["name"],
            "btm_price": format_price(btm_price),
            "tte_code": tte["article_number"] if tte else "",
            "tte_name": tte["name"] if tte else "",
            "havit_net": format_price(havit_net),
            "havit_vat": format_price(havit_vat),
            "havit_difference": format_price(havit_difference),
            "etail_net": format_price(etail_net),
            "etail_vat": format_price(etail_vat),
            "etail_difference": format_price(etail_difference),
            "match": (
                "HAVIT + ETAIL" if havit and etail
                else "HAVIT" if havit
                else "ETAIL" if etail
                else "NIJE PRONAĐEN"
            )
        })

    # --------------------------------------------------------
    # EXCEL
    # --------------------------------------------------------
    print("")
    print("=" * 70)
    print("PRAVLJENJE EXCELA")
    print("=" * 70)

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "BTM vs TTE"

    headers = [
        "EAN",
        "BTM šifra",
        "BTM naziv",
        "BTM cena sa PDV",
        "TTE šifra",
        "TTE naziv",
        "HAVIT neto",
        "HAVIT sa PDV",
        "Razlika HAVIT",
        "ETAIL neto",
        "ETAIL sa PDV",
        "Razlika ETAIL",
        "Poklapanje"
    ]

    sheet.append(headers)

    for cell in sheet[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for item in results:
        sheet.append([
            item["ean"],
            item["btm_code"],
            item["btm_name"],
            item["btm_price"],
            item["tte_code"],
            item["tte_name"],
            item["havit_net"],
            item["havit_vat"],
            item["havit_difference"],
            item["etail_net"],
            item["etail_vat"],
            item["etail_difference"],
            item["match"]
        ])

    for row in range(2, sheet.max_row + 1):
        for col in [4, 7, 8, 9, 10, 11, 12]:
            cell = sheet.cell(row=row, column=col)
            if isinstance(cell.value, (int, float)):
                cell.number_format = "#,##0.00"

        for col in [8, 11, 13]:
            sheet.cell(row=row, column=col).font = Font(bold=True)

        for cell in sheet[row]:
            cell.alignment = Alignment(vertical="center")

    widths = {
        1: 16,
        2: 15,
        3: 55,
        4: 18,
        5: 15,
        6: 55,
        7: 16,
        8: 18,
        9: 18,
        10: 16,
        11: 18,
        12: 18,
        13: 18,
    }

    for column, width in widths.items():
        sheet.column_dimensions[get_column_letter(column)].width = width

    sheet.auto_filter.ref = sheet.dimensions
    sheet.freeze_panes = "A2"

    workbook.save(OUTPUT_FILE)

    # --------------------------------------------------------
    # STATISTIKA
    # --------------------------------------------------------
    print("")
    print("=" * 70)
    print("ZAVRŠENO")
    print("=" * 70)
    print("BTM artikala:", len(btm_products))
    print("HAVIT poklapanja po EAN-u:", matched_havit)
    print("ETAIL poklapanja po EAN-u:", matched_etail)
    print("Bez poklapanja:", unmatched)
    print("Izlaz:", OUTPUT_FILE)


if __name__ == "__main__":
    main()
