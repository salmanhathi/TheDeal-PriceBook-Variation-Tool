from flask import Flask, request, render_template, send_file
import pandas as pd
import xml.etree.ElementTree as ET
from xml.dom import minidom
from datetime import datetime, timedelta
import os

app = Flask(__name__)

# ─────────────────────────────────────────
# ACCESS CONTROL
# Set ACCESS_ON env variable in Render dashboard to "true" or "false"
# No code push needed — just change it in Render and restart
# ─────────────────────────────────────────
ACCESS_ON = os.environ.get("ACCESS_ON", "true").lower() == "true"

@app.before_request
def check_access():
    if not ACCESS_ON:
        return "🔒 Tool is currently offline. Contact the admin for access.", 403

# ─────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────
UPLOAD_FOLDER = 'uploads'
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

PRODUCT_COL = "ProductId"
PRICEBOOK_COL = "PriceBookId"
NORMAL_PRICE_COL = "Price"
DISCOUNT_PRICE_COL = "DiscountedPrice"
ONLINE_FROM_COL = "DateFrom"
ONLINE_TO_COL = "DateTo"

CURRENCY_MAP = {
    "TheDealOutlet-list-AED-prices": "AED",
    "aed-list-prices": "AED",
    "TheDealOutlet-list-SAR-prices": "SAR",
    "sar-list-prices": "SAR",
    "Staff-Price-Book-AED-promo-prices": "AED"
}

PARENT_MAP = {
    "TheDealOutlet-list-AED-prices": "aed-list-prices",
    "TheDealOutlet-list-SAR-prices": "sar-list-prices",
    "Staff-Price-Book-AED-promo-prices": "aed-list-prices"
}

# ─────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────
def format_date(date_val, is_start=True):
    """Convert Excel date to UTC format (UAE is UTC+4)."""
    if pd.isna(date_val):
        return None
    if not isinstance(date_val, datetime):
        date_val = pd.to_datetime(date_val, errors="coerce")
    if pd.isna(date_val):
        return None
    if is_start:
        dt_uae = datetime(date_val.year, date_val.month, date_val.day, 0, 0, 0)
    else:
        dt_uae = datetime(date_val.year, date_val.month, date_val.day, 23, 59, 0)
    dt_utc = dt_uae - timedelta(hours=4)
    return dt_utc.strftime("%Y-%m-%dT%H:%M:%S.000Z")

# ─────────────────────────────────────────
# ROUTES
# ─────────────────────────────────────────
@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        file = request.files.get('file')
        if not file or not file.filename.endswith(('.xls', '.xlsx')):
            return "Please upload a valid Excel file (.xls or .xlsx).", 400

        filepath = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(filepath)
        df = pd.read_excel(filepath)

        root = ET.Element("pricebooks", {
            "xmlns": "http://www.demandware.com/xml/impex/pricebook/2006-10-31",
            "xmlns:xsi": "http://www.w3.org/2001/XMLSchema-instance",
            "xmlns:xsd": "http://www.w3.org/2001/XMLSchema"
        })

        for pricebook_id, group in df.groupby(PRICEBOOK_COL):
            pricebook_el = ET.SubElement(root, "pricebook")

            # Header
            header_el = ET.SubElement(pricebook_el, "header", {
                "pricebook-id": str(pricebook_id)
            })
            currency_el = ET.SubElement(header_el, "currency")
            currency_el.text = CURRENCY_MAP.get(pricebook_id, "AED")

            display_el = ET.SubElement(header_el, "display-name", {"xml:lang": "x-default"})
            display_el.text = str(pricebook_id)

            online_flag_el = ET.SubElement(header_el, "online-flag")
            online_flag_el.text = "true"

            if pricebook_id in PARENT_MAP:
                parent_el = ET.SubElement(header_el, "parent")
                parent_el.text = PARENT_MAP[pricebook_id]

            price_tables_el = ET.SubElement(pricebook_el, "price-tables")

            for _, row in group.iterrows():
                product_id = str(row[PRODUCT_COL]).zfill(12)

                # ── Staff Price Book logic ──
                if pricebook_id == "Staff-Price-Book-AED-promo-prices":
                    if (
                        pd.notna(row[DISCOUNT_PRICE_COL]) and
                        pd.notna(row[ONLINE_FROM_COL]) and
                        pd.notna(row[ONLINE_TO_COL])
                    ):
                        pt_el = ET.SubElement(price_tables_el, "price-table", {
                            "product-id": product_id
                        })
                        online_from_el = ET.SubElement(pt_el, "online-from")
                        online_from_el.text = format_date(row[ONLINE_FROM_COL], True)
                        online_to_el = ET.SubElement(pt_el, "online-to")
                        online_to_el.text = format_date(row[ONLINE_TO_COL], False)
                        amt_el = ET.SubElement(pt_el, "amount", {"quantity": "1"})
                        amt_el.text = str(row[DISCOUNT_PRICE_COL])
                    continue  # skip normal logic for staff pricebook

                # ── Normal price ──
                pt_el = ET.SubElement(price_tables_el, "price-table", {
                    "product-id": product_id
                })
                amt_el = ET.SubElement(pt_el, "amount", {"quantity": "1"})
                amt_el.text = str(row[NORMAL_PRICE_COL])

                # ── Discount price (with dates) ──
                if (
                    pd.notna(row[DISCOUNT_PRICE_COL]) and
                    pd.notna(row[ONLINE_FROM_COL]) and
                    pd.notna(row[ONLINE_TO_COL])
                ):
                    pt_disc_el = ET.SubElement(price_tables_el, "price-table", {
                        "product-id": product_id
                    })
                    online_from_el = ET.SubElement(pt_disc_el, "online-from")
                    online_from_el.text = format_date(row[ONLINE_FROM_COL], True)
                    online_to_el = ET.SubElement(pt_disc_el, "online-to")
                    online_to_el.text = format_date(row[ONLINE_TO_COL], False)
                    amt_disc_el = ET.SubElement(pt_disc_el, "amount", {"quantity": "1"})
                    amt_disc_el.text = str(row[DISCOUNT_PRICE_COL])

        output_filename = f"UAE-PriceBook-{datetime.now().strftime('%Y%m%d-%H%M%S')}.xml"
        xml_str = minidom.parseString(
            ET.tostring(root, "utf-8")
        ).toprettyxml(indent="  ", encoding="utf-8")

        xml_path = os.path.join(UPLOAD_FOLDER, output_filename)
        with open(xml_path, "wb") as f:
            f.write(xml_str)

        return send_file(xml_path, as_attachment=True)

    return render_template('upload.html')


if __name__ == '__main__':
    app.run(host='0.0.0.0', debug=False)
