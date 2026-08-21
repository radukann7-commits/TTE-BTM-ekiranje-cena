import os
import html
import re
import io
import csv
import zipfile
import xlrd
import requests
import xml.etree.ElementTree as ET

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter

BTM_FILE = "BTM_export.xls"
OUTPUT_FILE = "BTM_TTE_poredjenje.xlsx"

TTE_ETAIL_API_KEY = os.environ.get("TTE_ETAIL_API_KEY")
TTE_HAVIT_API_KEY = os.environ.get("TTE_HAVIT_API_KEY")

ETAIL_URL = "https://tte.rs/api/sr/products?output_type=xml&api_key=" + str(TTE_ETAIL_API_KEY)
HAVIT_URL = "https://tte.rs/api/sr/products?output_type=xml&api_key=" + str(TTE_HAVIT_API_KEY)

# Ovi brendovi se izbacuju iz poređenja.
# JBL, Samsung i Apple ostaju.
EXCLUDED_BRANDS = {
    "bavin", "bavitel", "havit", "powerology", "green lion", "porodo"
}


def clean_text(value):
    if value is None:
        return ""
    value = html.unescape(str(value))
    value = re.sub(r"<[^>]+>", " ", value)
    value = value.replace("\xa0", " ").replace("\n", " ").replace("\r", " ").replace("\t", " ")
    return re.sub(r"\s+", " ", value).strip()


def normalize_text(value):
    value = clean_text(value).lower()
    for a, b in {"č": "c", "ć": "c", "š": "s", "ž": "z", "đ": "d"}.items():
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
    text = text.replace("RSD", "").replace("rsd", "").replace("DIN", "").replace("din", "").replace(" ", "")
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
    return clean_text(element.text) if element is not None else ""


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


def read_tsv_or_xlsx_or_xls(data):
    text = data.decode("utf-8-sig", errors="replace")
    first_line = text.splitlines()[0] if text.splitlines() else ""

    if "\t" in first_line:
        return list(csv.reader(io.StringIO(text), delimiter="\t"))

    if zipfile.is_zipfile(io.BytesIO(data)):
        wb = load_workbook(io.BytesIO(data), data_only=True, read_only=True)
        ws = wb.active
        values = [list(row) for row in ws.iter_rows(values_only=True)]
        wb.close()
        return values

    wb = xlrd.open_workbook(file_contents=data, formatting_info=False)
    ws = wb.sheet_by_index(0)
    return [[ws.cell_value(r, c) for c in range(ws.ncols)] for r in range(ws.nrows)]


def find_btm_header(values):
    best_row = None
    best_score = 0
    for row_index in range(min(len(values), 50)):
        row_text = " | ".join(clean_text(x).lower() for x in values[row_index])
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


def read_btm():
    if not os.path.exists(BTM_FILE):
        raise Exception(f"BTM fajl ne postoji: {BTM_FILE}")

    data = open(BTM_FILE, "rb").read()
    values = read_tsv_or_xlsx_or_xls(data)
    header_row = find_btm_header(values)
    headers = [clean_text(x) for x in values[header_row]]

    ean_col = find_column(headers, ["EAN", "Barkod", "Bar code", "Barcode", "EAN kod"])
    name_col = find_column(headers, ["Naziv", "Naziv artikla", "Artikal naziv", "Product name", "Name"])
    price_col = find_column(headers, ["Cena", "Cena RSD", "Cena bez PDV", "Cena NETO", "Neto cena", "Price"])
    brand_col = find_column(headers, ["Brend", "Brand", "Proizvođač", "Proizvodjac"])
    tte_code_col = find_column(headers, ["TTE šifra", "TTE sifra", "Šifra", "Sifra", "Artikal", "Product code", "Code"])

    if ean_col is None:
        raise Exception("Ne mogu da pronađem BTM EAN/Barkod kolonu.")
    if name_col is None:
        raise Exception("Ne mogu da pronađem BTM naziv artikla.")
    if price_col is None:
        raise Exception("Ne mogu da pronađem BTM cenu.")

    products = []
    for row in values[header_row + 1:]:
        row = list(row) + [""] * max(0, len(headers) - len(row))
        ean = normalize_ean(row[ean_col])
        name = clean_text(row[name_col])
        brand = clean_text(row[brand_col]) if brand_col is not None else ""
        price = number_from_text(row[price_col])
        tte_code = clean_text(row[tte_code_col]) if tte_code_col is not None else ""

        if not ean or not name or price is None:
            continue

        # Ako BTM ima posebnu kolonu brenda, koristi nju za izbacivanje.
        # Ako nema, proveri i naziv artikla kao rezervu.
        brand_norm = normalize_text(brand)
        name_norm = normalize_text(name)
        if brand_norm in EXCLUDED_BRANDS:
            continue
        if not brand_norm and any(re.search(r"\b" + re.escape(b) + r"\b", name_norm) for b in EXCLUDED_BRANDS):
            continue

        products.append({
            "ean": ean,
            "name": name,
            "net_price": price,
            "tte_code": tte_code,
        })

    print("BTM artikala posle filtera:", len(products))
    return products


def download_xml(url, source_name):
    response = requests.get(url, timeout=120, headers={"User-Agent": "Mozilla/5.0"})
    response.raise_for_status()
    if not response.content:
        raise Exception(f"{source_name}: XML je prazan.")
    try:
        ET.fromstring(response.content)
    except Exception as e:
        raise Exception(f"{source_name}: odgovor nije ispravan XML: {e}")
    return response.content


def parse_tte_xml(xml_data, source_name):
    root = ET.fromstring(xml_data)
    products = {}

    for product in root.findall(".//product"):
        ean = normalize_ean(get_xml_value(product, "ean"))
        if not ean:
            continue

        # VAŽNO: uzima se ISKLJUČIVO netto_price.
        # Nema fallback-a na neto_price, net_price ili price.
        netto_text = get_xml_value(product, "netto_price")
        netto_price = number_from_text(netto_text)
        if netto_price is None:
            continue

        if ean not in products:
            products[ean] = {
                "netto_price": netto_price,
                "article_number": get_xml_value(product, "article_number"),
            }

    print(f"{source_name}: artikala sa EAN + netto_price = {len(products)}")
    return products


def main():
    if not TTE_ETAIL_API_KEY:
        raise Exception("Nedostaje GitHub Secret: TTE_ETAIL_API_KEY")
    if not TTE_HAVIT_API_KEY:
        raise Exception("Nedostaje GitHub Secret: TTE_HAVIT_API_KEY")

    btm_products = read_btm()
    etail_products = parse_tte_xml(download_xml(ETAIL_URL, "ETAIL SPEC"), "ETAIL SPEC")
    havit_products = parse_tte_xml(download_xml(HAVIT_URL, "HAVIT"), "HAVIT")

    results = []

    for btm in btm_products:
        ean = btm["ean"]
        havit = havit_products.get(ean)
        etail = etail_products.get(ean)

        # Poređenje se radi samo kada postoji bar jedna odgovarajuća TTE cena.
        if havit is None and etail is None:
            continue

        btm_net = btm["net_price"]
        havit_net = havit["netto_price"] if havit else None
        etail_net = etail["netto_price"] if etail else None

        havit_diff = round(havit_net - btm_net, 2) if havit_net is not None else None
        etail_diff = round(etail_net - btm_net, 2) if etail_net is not None else None

        # Prikazuju se samo artikli gde postoji razlika u ceni.
        if not (
            (havit_diff is not None and abs(havit_diff) >= 0.01)
            or (etail_diff is not None and abs(etail_diff) >= 0.01)
        ):
            continue

        # TTE šifra: prvo HAVIT XML, zatim ETAIL SPEC XML, pa BTM ako postoji.
        tte_code = ""
        if havit and havit.get("article_number"):
            tte_code = havit["article_number"]
        elif etail and etail.get("article_number"):
            tte_code = etail["article_number"]
        elif btm.get("tte_code"):
            tte_code = btm["tte_code"]

        results.append([
            ean,
            tte_code,
            btm["name"],
            format_price(btm_net),
            format_price(havit_net),
            format_price(havit_diff),
            format_price(etail_net),
            format_price(etail_diff),
        ])

    results.sort(key=lambda row: row[0])

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "BTM vs TTE"

    headers = [
        "EAN",
        "TTE šifra",
        "Artikal",
        "BTM cena NETO",
        "HAVIT cena NETO",
        "Razlika HAVIT",
        "ETAIL SPEC cena NETO",
        "Razlika ETAIL SPEC",
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
                cell.number_format = '#,##0.00'

    widths = [16, 18, 60, 18, 18, 16, 21, 20]
    for i, width in enumerate(widths, 1):
        sheet.column_dimensions[get_column_letter(i)].width = width

    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions
    workbook.save(OUTPUT_FILE)

    print("ARTIKALA SA RAZLIKOM:", len(results))
    print("IZLAZ:", OUTPUT_FILE)


if __name__ == "__main__":
    main()
