from flask import Flask, request, render_template, send_file
import pandas as pd
import xml.etree.ElementTree as ET
from xml.dom import minidom
from datetime import datetime, timedelta
from io import BytesIO
import os
import sys

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

app = Flask(
    __name__,
    template_folder=os.path.join(BASE_DIR, "templates"),
    static_folder=os.path.join(BASE_DIR, "static"),
)

# Fail loudly and clearly at startup if templates/static are missing,
# instead of letting Flask throw a TemplateNotFound stack trace on the
# first request. This matters because template_folder is resolved
# relative to this file's location, not the current working directory --
# if someone only copies app.py without the templates/static folders
# sitting right next to it, this catches that immediately.
_missing = []
if not os.path.isdir(app.template_folder):
    _missing.append(app.template_folder)
if not os.path.isdir(app.static_folder):
    _missing.append(app.static_folder)
if _missing:
    print("=" * 70)
    print("STARTUP ERROR: required folder(s) not found next to app.py:")
    for m in _missing:
        print(f"  - {m}")
    print(f"\napp.py is running from: {BASE_DIR}")
    print("Expected layout:")
    print(f"  {BASE_DIR}\\app.py")
    print(f"  {BASE_DIR}\\templates\\  (index.html, pricebook.html, variants.html, base.html)")
    print(f"  {BASE_DIR}\\static\\     (style.css, app.js)")
    print("=" * 70)
    sys.exit(1)

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
# SHARED
# ─────────────────────────────────────────
def read_upload_as_dataframe(file_storage, string_columns=None):
    """Validate + load an uploaded Excel file into a DataFrame.

    string_columns: optional list of column names to force-read as text.
    This matters because Excel silently converts numeric-looking IDs/codes
    (e.g. "02851021285", a size code like "05") into actual numbers, which
    strips leading zeros before pandas ever sees the value -- unrecoverable
    once that's happened. Forcing dtype=str at read time prevents it.
    """
    if not file_storage or not file_storage.filename:
        raise ValueError("No file was uploaded.")
    if not file_storage.filename.lower().endswith((".xls", ".xlsx")):
        raise ValueError("Please upload a valid Excel file (.xls or .xlsx).")

    if not string_columns:
        return pd.read_excel(file_storage)

    # Peek at the header row to build a dtype map without guessing column
    # order, then re-read the file (file_storage supports seek(0) for
    # both real uploads and BytesIO in tests).
    header_df = pd.read_excel(file_storage, nrows=0)
    dtype_map = {}
    for target in string_columns:
        for col in header_df.columns:
            if str(col).strip().lower() == target.strip().lower():
                dtype_map[col] = str
                break

    file_storage.seek(0)
    return pd.read_excel(file_storage, dtype=dtype_map) if dtype_map else pd.read_excel(file_storage)


def xml_bytes_from_element(root_el):
    """Pretty-print an ElementTree root into UTF-8 XML bytes (no BOM)."""
    rough = ET.tostring(root_el, "utf-8")
    return minidom.parseString(rough).toprettyxml(indent="  ", encoding="utf-8")


def send_xml(xml_bytes, filename):
    buf = BytesIO(xml_bytes)
    buf.seek(0)
    return send_file(
        buf,
        as_attachment=True,
        download_name=filename,
        mimetype="application/xml",
    )


# ─────────────────────────────────────────
# PRICEBOOK GENERATOR
# ─────────────────────────────────────────
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
    "Staff-Price-Book-AED-promo-prices": "AED",
}

PARENT_MAP = {
    "TheDealOutlet-list-AED-prices": "aed-list-prices",
    "TheDealOutlet-list-SAR-prices": "sar-list-prices",
    "Staff-Price-Book-AED-promo-prices": "aed-list-prices",
}


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


def build_pricebook_xml(df):
    required = {PRODUCT_COL, PRICEBOOK_COL, NORMAL_PRICE_COL}
    missing = required - set(df.columns)
    if missing:
        raise ValueError(f"Missing required column(s): {', '.join(sorted(missing))}")

    root = ET.Element(
        "pricebooks",
        {
            "xmlns": "http://www.demandware.com/xml/impex/pricebook/2006-10-31",
            "xmlns:xsi": "http://www.w3.org/2001/XMLSchema-instance",
            "xmlns:xsd": "http://www.w3.org/2001/XMLSchema",
        },
    )

    for pricebook_id, group in df.groupby(PRICEBOOK_COL, sort=False):
        pricebook_el = ET.SubElement(root, "pricebook")

        header_el = ET.SubElement(pricebook_el, "header", {"pricebook-id": str(pricebook_id)})
        ET.SubElement(header_el, "currency").text = CURRENCY_MAP.get(pricebook_id, "AED")

        display_el = ET.SubElement(header_el, "display-name", {"xml:lang": "x-default"})
        display_el.text = str(pricebook_id)

        ET.SubElement(header_el, "online-flag").text = "true"

        if pricebook_id in PARENT_MAP:
            ET.SubElement(header_el, "parent").text = PARENT_MAP[pricebook_id]

        price_tables_el = ET.SubElement(pricebook_el, "price-tables")

        for _, row in group.iterrows():
            product_id = str(row[PRODUCT_COL]).zfill(12)

            has_discount = (
                DISCOUNT_PRICE_COL in row
                and pd.notna(row.get(DISCOUNT_PRICE_COL))
                and pd.notna(row.get(ONLINE_FROM_COL))
                and pd.notna(row.get(ONLINE_TO_COL))
            )

            if pricebook_id == "Staff-Price-Book-AED-promo-prices":
                if has_discount:
                    pt_el = ET.SubElement(price_tables_el, "price-table", {"product-id": product_id})
                    ET.SubElement(pt_el, "online-from").text = format_date(row[ONLINE_FROM_COL], True)
                    ET.SubElement(pt_el, "online-to").text = format_date(row[ONLINE_TO_COL], False)
                    ET.SubElement(pt_el, "amount", {"quantity": "1"}).text = str(row[DISCOUNT_PRICE_COL])
                continue

            pt_el = ET.SubElement(price_tables_el, "price-table", {"product-id": product_id})
            ET.SubElement(pt_el, "amount", {"quantity": "1"}).text = str(row[NORMAL_PRICE_COL])

            if has_discount:
                pt_disc_el = ET.SubElement(price_tables_el, "price-table", {"product-id": product_id})
                ET.SubElement(pt_disc_el, "online-from").text = format_date(row[ONLINE_FROM_COL], True)
                ET.SubElement(pt_disc_el, "online-to").text = format_date(row[ONLINE_TO_COL], False)
                ET.SubElement(pt_disc_el, "amount", {"quantity": "1"}).text = str(row[DISCOUNT_PRICE_COL])

    return xml_bytes_from_element(root)


# ─────────────────────────────────────────
# VARIANTS GENERATOR
# ─────────────────────────────────────────
VARIANT_PRODUCT_COL = "product_id"
VARIANT_COLOR_COL = "color_name"
VARIANT_SIZE_CODE_COL = "size_code"
VARIANT_SIZE_NAME_COL = "size_name"


def _find_column(df, target_name):
    """Case-insensitive, whitespace-tolerant column lookup."""
    target = target_name.strip().lower()
    for col in df.columns:
        if str(col).strip().lower() == target:
            return col
    return None


def _color_value_and_display(color_name):
    """SFCC color code + display name from a raw source value.

    value:   lowercased, spaces -> underscores (e.g. "Multi Color" -> "multi_color")
    display: underscore -> space FIRST, then title-case (matches the
             original VBA order exactly -- "multi_color" -> "Multi Color",
             not "Multi_Color" or "MultiColor").
    """
    value = color_name.lower().replace(" ", "_")
    display = color_name.replace("_", " ").title()
    return value, display


VARIANT_MODES = ("COLOR", "SIZE", "BOTH", "SKU")


def build_variants_xml(df, mode, catalog_id):
    """mode: 'COLOR', 'SIZE', or 'BOTH'"""
    mode = mode.upper()
    if mode not in ("COLOR", "SIZE", "BOTH"):
        raise ValueError("Mode must be COLOR, SIZE, or BOTH.")

    pid_col = _find_column(df, VARIANT_PRODUCT_COL)
    color_col = _find_column(df, VARIANT_COLOR_COL)
    size_code_col = _find_column(df, VARIANT_SIZE_CODE_COL)
    size_name_col = _find_column(df, VARIANT_SIZE_NAME_COL)

    if pid_col is None:
        raise ValueError(f"Missing required column: {VARIANT_PRODUCT_COL}")
    if mode in ("COLOR", "BOTH") and color_col is None:
        raise ValueError(f"Missing required column for {mode} mode: {VARIANT_COLOR_COL}")
    if mode in ("SIZE", "BOTH") and (size_code_col is None or size_name_col is None):
        raise ValueError(f"Missing required column(s) for {mode} mode: {VARIANT_SIZE_CODE_COL}, {VARIANT_SIZE_NAME_COL}")

    # product_id -> ordered unique colors ; product_id -> ordered {size_code: size_name}
    colors_by_pid = {}
    sizes_by_pid = {}
    pid_order = []

    for _, row in df.iterrows():
        pid = str(row[pid_col]).strip()
        if not pid or pid.lower() == "nan":
            continue
        if pid not in pid_order:
            pid_order.append(pid)
            colors_by_pid[pid] = {}
            sizes_by_pid[pid] = {}

        if mode in ("COLOR", "BOTH"):
            color_val = row.get(color_col)
            if pd.notna(color_val):
                color_name = str(color_val).strip()
                if color_name and color_name not in colors_by_pid[pid]:
                    colors_by_pid[pid][color_name] = True

        if mode in ("SIZE", "BOTH"):
            size_code_val = row.get(size_code_col)
            if pd.notna(size_code_val):
                size_code = str(size_code_val).strip()
                size_name_val = row.get(size_name_col)
                size_name = str(size_name_val).strip() if pd.notna(size_name_val) else size_code
                if size_code and size_code not in sizes_by_pid[pid]:
                    sizes_by_pid[pid][size_code] = size_name

    root = ET.Element(
        "catalog",
        {
            "xmlns": "http://www.demandware.com/xml/impex/catalog/2006-10-31",
            "catalog-id": catalog_id,
        },
    )

    for pid in pid_order:
        colors = colors_by_pid[pid]
        sizes = sizes_by_pid[pid]
        has_color = mode != "SIZE" and len(colors) > 0
        has_size = mode != "COLOR" and len(sizes) > 0

        if not has_color and not has_size:
            continue

        product_el = ET.SubElement(root, "product", {"product-id": pid})
        variations_el = ET.SubElement(product_el, "variations")
        attributes_el = ET.SubElement(variations_el, "attributes")

        if has_color:
            va_el = ET.SubElement(
                attributes_el,
                "variation-attribute",
                {"attribute-id": "color", "variation-attribute-id": "color"},
            )
            vav_el = ET.SubElement(va_el, "variation-attribute-values", {"merge-mode": "add"})
            for color_name in colors.keys():
                value, display = _color_value_and_display(color_name)
                vv_el = ET.SubElement(vav_el, "variation-attribute-value", {"value": value})
                ET.SubElement(vv_el, "display-value", {"xml:lang": "x-default"}).text = display

        if has_size:
            va_el = ET.SubElement(
                attributes_el,
                "variation-attribute",
                {"attribute-id": "size", "variation-attribute-id": "size"},
            )
            vav_el = ET.SubElement(va_el, "variation-attribute-values", {"merge-mode": "add"})
            for size_code, size_name in sizes.items():
                vv_el = ET.SubElement(vav_el, "variation-attribute-value", {"value": size_code})
                ET.SubElement(vv_el, "display-value", {"xml:lang": "x-default"}).text = size_name

    return xml_bytes_from_element(root)


def build_sku_variants_xml(df, catalog_id):
    """
    'SKU' mode: real per-color/size SKU pairing, the thing the pooled
    attribute-palette file above structurally cannot express.

    For each master style, emits:
      - the same full color + size attribute palette as BOTH mode
        (a master must still declare every value it offers overall)
      - a <variants> list naming exactly which SKUs exist for this style

    Then, as separate flat <product> entries, one per unique (style,
    color, size) pairing actually present in the source rows:
      - product-id built as "{style}-{color_value}-{size_code}"
        (color segment uses the same slug as the attribute palette, e.g.
        "white", "multi_color" -- size segment is the raw size_code,
        unchanged, matching the confirmed format
        "013614BGH027142A-white-XXS")
      - <variation-values> pinning that one color + one size

    This combination -- palette + variants list + per-SKU variation-values
    -- is what actually restricts Black to {37 EU, 41 EU} and White to
    {41 EU, 42 EU} in SFCC. The palette file alone can only ever pool.
    """
    pid_col = _find_column(df, VARIANT_PRODUCT_COL)
    color_col = _find_column(df, VARIANT_COLOR_COL)
    size_code_col = _find_column(df, VARIANT_SIZE_CODE_COL)
    size_name_col = _find_column(df, VARIANT_SIZE_NAME_COL)

    missing = [
        name
        for name, col in (
            (VARIANT_PRODUCT_COL, pid_col),
            (VARIANT_COLOR_COL, color_col),
            (VARIANT_SIZE_CODE_COL, size_code_col),
            (VARIANT_SIZE_NAME_COL, size_name_col),
        )
        if col is None
    ]
    if missing:
        raise ValueError(f"Missing required column(s) for SKU mode: {', '.join(missing)}")

    colors_by_pid = {}       # pid -> ordered {color_name: True}          (full palette)
    sizes_by_pid = {}        # pid -> ordered {size_code: size_name}      (full palette)
    pairs_by_pid = {}        # pid -> ordered {(color_name, size_code): size_name}  (real SKUs)
    pid_order = []

    for _, row in df.iterrows():
        pid = str(row[pid_col]).strip()
        if not pid or pid.lower() == "nan":
            continue

        color_val = row.get(color_col)
        size_code_val = row.get(size_code_col)
        if pd.isna(color_val) or pd.isna(size_code_val):
            # A SKU needs both a color and a size -- a row missing either
            # can't form a real pairing, so it's skipped for this mode.
            continue

        color_name = str(color_val).strip()
        size_code = str(size_code_val).strip()
        if not color_name or not size_code:
            continue

        size_name_val = row.get(size_name_col)
        size_name = str(size_name_val).strip() if pd.notna(size_name_val) else size_code

        if pid not in pid_order:
            pid_order.append(pid)
            colors_by_pid[pid] = {}
            sizes_by_pid[pid] = {}
            pairs_by_pid[pid] = {}

        if color_name not in colors_by_pid[pid]:
            colors_by_pid[pid][color_name] = True
        if size_code not in sizes_by_pid[pid]:
            sizes_by_pid[pid][size_code] = size_name

        pair_key = (color_name, size_code)
        if pair_key not in pairs_by_pid[pid]:
            pairs_by_pid[pid][pair_key] = size_name

    root = ET.Element(
        "catalog",
        {
            "xmlns": "http://www.demandware.com/xml/impex/catalog/2006-10-31",
            "catalog-id": catalog_id,
        },
    )

    for pid in pid_order:
        colors = colors_by_pid[pid]
        sizes = sizes_by_pid[pid]
        pairs = pairs_by_pid[pid]
        if not pairs:
            continue

        product_el = ET.SubElement(root, "product", {"product-id": pid})
        variations_el = ET.SubElement(product_el, "variations")
        attributes_el = ET.SubElement(variations_el, "attributes")

        va_color_el = ET.SubElement(
            attributes_el, "variation-attribute", {"attribute-id": "color", "variation-attribute-id": "color"}
        )
        vav_color_el = ET.SubElement(va_color_el, "variation-attribute-values", {"merge-mode": "add"})
        for color_name in colors.keys():
            value, display = _color_value_and_display(color_name)
            vv_el = ET.SubElement(vav_color_el, "variation-attribute-value", {"value": value})
            ET.SubElement(vv_el, "display-value", {"xml:lang": "x-default"}).text = display

        va_size_el = ET.SubElement(
            attributes_el, "variation-attribute", {"attribute-id": "size", "variation-attribute-id": "size"}
        )
        vav_size_el = ET.SubElement(va_size_el, "variation-attribute-values", {"merge-mode": "add"})
        for size_code, size_name in sizes.items():
            vv_el = ET.SubElement(vav_size_el, "variation-attribute-value", {"value": size_code})
            ET.SubElement(vv_el, "display-value", {"xml:lang": "x-default"}).text = size_name

        # This <variants> list is what actually restricts each color to its
        # real sizes -- e.g. White only ever lists 41sEU/42sEU here, even
        # though the palette above still declares 37sEU as a valid size
        # overall (because Black offers it). No separate flat <product>
        # entries are emitted per variant -- just this list, per your spec.
        variants_el = ET.SubElement(variations_el, "variants")
        for i, (color_name, size_code) in enumerate(pairs.keys()):
            color_value, _ = _color_value_and_display(color_name)
            variant_id = f"{pid}-{color_value}-{size_code}"
            variant_attrs = {"product-id": variant_id}
            if i == 0:
                variant_attrs["default"] = "true"
            ET.SubElement(variants_el, "variant", variant_attrs)

    return xml_bytes_from_element(root)


# ─────────────────────────────────────────
# ROUTES
# ─────────────────────────────────────────
@app.route("/")
def home():
    return render_template("index.html")


@app.route("/pricebook", methods=["GET"])
def pricebook_page():
    return render_template("pricebook.html")


@app.route("/pricebook/generate", methods=["POST"])
def pricebook_generate():
    try:
        df = read_upload_as_dataframe(
            request.files.get("file"),
            string_columns=[PRODUCT_COL, PRICEBOOK_COL],
        )
        xml_bytes = build_pricebook_xml(df)
    except ValueError as e:
        return str(e), 400
    except Exception as e:
        return f"Failed to generate pricebook XML: {e}", 500

    filename = f"UAE-PriceBook-{datetime.now().strftime('%Y%m%d-%H%M%S')}.xml"
    return send_xml(xml_bytes, filename)


@app.route("/variants", methods=["GET"])
def variants_page():
    return render_template("variants.html")


@app.route("/variants/generate", methods=["POST"])
def variants_generate():
    try:
        df = read_upload_as_dataframe(
            request.files.get("file"),
            string_columns=[
                VARIANT_PRODUCT_COL,
                VARIANT_COLOR_COL,
                VARIANT_SIZE_CODE_COL,
                VARIANT_SIZE_NAME_COL,
            ],
        )
        mode = request.form.get("mode", "BOTH").strip().upper()
        catalog_id = request.form.get("catalog_id", "thedealoutlet-master-catalog").strip() or "thedealoutlet-master-catalog"

        if mode not in VARIANT_MODES:
            raise ValueError(f"Mode must be one of: {', '.join(VARIANT_MODES)}")

        if mode == "SKU":
            xml_bytes = build_sku_variants_xml(df, catalog_id)
        else:
            xml_bytes = build_variants_xml(df, mode, catalog_id)
    except ValueError as e:
        return str(e), 400
    except Exception as e:
        return f"Failed to generate variants XML: {e}", 500

    filename = f"Import-Variants-{mode.upper()}-{datetime.now().strftime('%Y%m%d-%H%M%S')}.xml"
    return send_xml(xml_bytes, filename)


if __name__ == "__main__":
    app.run(host="0.0.0.0", debug=False)
