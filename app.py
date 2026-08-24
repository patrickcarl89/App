"""
SKU Storage & Sales Tracker — Streamlit + Google Sheets edition
=================================================================
A single-file POS / inventory web app. Google Sheets is the database
(via gspread + a service account).

Run with:  streamlit run app.py
"""

import datetime as dt
import os
import uuid

import pandas as pd
import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
from PIL import Image, ImageDraw, ImageFont

# Optional camera-based barcode/QR decoding. The app works fully without
# pyzbar installed — the SKU field always accepts manual typing and
# USB/Bluetooth hardware scanners (which just "type" into the field).
try:
    from pyzbar.pyzbar import decode as zbar_decode
    SCANNER_AVAILABLE = True
except ImportError:
    SCANNER_AVAILABLE = False


# ----------------------------------------------------------------------
# PAGE CONFIG + STYLE
# ----------------------------------------------------------------------
st.set_page_config(
    page_title="SKU Tracker",
    page_icon="🧾",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown("""
<style>
    /* Hide Streamlit chrome */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header [data-testid="stToolbar"] {visibility: hidden;}

    /* Reclaim the top padding wide layout wastes */
    .block-container {padding-top: 1.2rem; padding-bottom: 3rem; max-width: 1400px;}

    /* Modern buttons */
    div.stButton > button {
        border-radius: 10px;
        height: 2.9em;
        font-weight: 600;
        border: 1px solid transparent;
        transition: all 0.15s ease;
    }
    div.stButton > button[kind="primary"] {
        background: linear-gradient(90deg,#16a34a,#059669);
        color: white;
        box-shadow: 0 2px 8px rgba(5,150,105,0.35);
    }
    div.stButton > button[kind="primary"]:hover {
        background: linear-gradient(90deg,#15803d,#047857);
        transform: translateY(-1px);
    }
    div.stButton > button[kind="secondary"] {
        background: #f1f5f9; color: #0f172a; border: 1px solid #e2e8f0;
    }
    div.stButton > button[kind="secondary"]:hover {
        background: #e2e8f0; transform: translateY(-1px);
    }
    div.stFormSubmitButton > button {
        background: linear-gradient(90deg,#2563eb,#4f46e5) !important;
        color: white !important; border: none !important;
    }

    /* Metric cards */
    [data-testid="stMetric"] {
        background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 12px;
        padding: 10px 14px;
    }
    [data-testid="stMetricValue"] {font-size: 1.4rem;}

    .app-banner {
        background: linear-gradient(90deg,#2563eb,#7c3aed);
        padding: 16px 20px; border-radius: 14px; color: white;
        margin-bottom: 14px;
    }
    .checkout-card {
        background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 14px;
        padding: 16px; margin-bottom: 10px;
    }
    div[data-testid="stExpander"] details summary p {font-weight: 600;}
</style>
""", unsafe_allow_html=True)


# ----------------------------------------------------------------------
# CONFIG
# ----------------------------------------------------------------------

# Items containing these keywords are FORCED into the "Extra" category no
# matter what category text was typed in. Case-insensitive substring match,
# so "Jasmine Rice" still hits the "rice" rule.
CATEGORY_OVERRIDES = {
    "rice": "Extra",
    "paper cups": "Extra",
    "napkins": "Extra",
    "plastic spoons": "Extra",
    "plastic forks": "Extra",
    "straws": "Extra",
    "sauce packets": "Extra",
    "condiments": "Extra",
}

SHEET_COLUMNS = {
    "Employees": ["EmployeeID", "Name", "PIN", "Role"],
    "Inventory": ["SKU", "ItemName", "Category", "UnitCost", "UnitPrice",
                  "QuantityOnHand", "ReorderLevel", "LastUpdated", "AvailableToday"],
    # OrderID groups every line item scanned/tapped into one checkout so a
    # whole receipt can be tracked / voided together.
    "Transactions": ["OrderID", "TransactionID", "DateTime", "SKU", "ItemName", "Quantity",
                      "UnitPrice", "TotalAmount", "PaymentMethod", "ReferenceNumber",
                      "StaffTabOwner", "EmployeeName", "Status", "VoidReason", "VoidedBy",
                      "VoidedDateTime"],
    "StockIntake": ["IntakeID", "DateTime", "SKU", "ItemName", "QuantityAdded",
                     "UnitCost", "EmployeeName", "Notes"],
    "AuditLog": ["LogID", "DateTime", "EmployeeName", "Action", "Details"],
    "Reconciliation": ["ReconID", "Date", "ExpectedCash", "ActualCash", "CashVariance",
                        "ExpectedOnline", "ActualOnline", "OnlineVariance",
                        "EmployeeName", "Notes"],
    # Cross-shift blind counts (Requirement 5).
    "ShiftCounts": ["ShiftCountID", "DateTime", "ShiftType", "SKU", "ItemName",
                     "ExpectedQty", "ActualQty", "Variance", "EmployeeName", "AuthorizedBy", "Notes"],
}

NUMERIC_COLUMNS = {
    "Inventory": ["UnitCost", "UnitPrice", "QuantityOnHand", "ReorderLevel"],
    "Transactions": ["Quantity", "UnitPrice", "TotalAmount"],
    "StockIntake": ["QuantityAdded", "UnitCost"],
    "Reconciliation": ["ExpectedCash", "ActualCash", "CashVariance",
                        "ExpectedOnline", "ActualOnline", "OnlineVariance"],
    "ShiftCounts": ["ExpectedQty", "ActualQty", "Variance"],
}

# Generous headroom so appends don't need to resize the grid mid-shift.
ROW_HINTS = {
    "Employees": 100, "Inventory": 1000, "Transactions": 5000,
    "StockIntake": 2000, "AuditLog": 5000, "Reconciliation": 1000,
    "ShiftCounts": 5000,
}

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

# Drop a real QR code image with this filename next to app.py to replace
# the generated placeholder used for Online payments.
QR_CODE_PATH = "payment_qr.png"


# ----------------------------------------------------------------------
# SMALL HELPERS
# ----------------------------------------------------------------------

def now_str():
    return dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def today_str():
    return dt.date.today().strftime("%Y-%m-%d")


def new_id(prefix):
    return f"{prefix}-{dt.datetime.now().strftime('%Y%m%d%H%M%S')}-{uuid.uuid4().hex[:5].upper()}"


def clean_str(x):
    if x is None:
        return ""
    s = str(x).strip()
    if s.endswith(".0"):
        try:
            float(s)
            s = s[:-2]
        except ValueError:
            pass
    return s


def fmt_money(x):
    try:
        return f"{float(x):,.2f}"
    except (ValueError, TypeError):
        return "0.00"


def resolve_category(item_name, requested_category):
    """Force-override categories for known 'Extra' items."""
    name_lower = (item_name or "").strip().lower()
    for keyword, forced_cat in CATEGORY_OVERRIDES.items():
        if keyword in name_lower:
            return forced_cat
    requested_category = (requested_category or "").strip()
    return requested_category if requested_category else "General"


def is_true(value):
    return str(value).strip().upper() == "TRUE"


@st.cache_data(show_spinner=False)
def get_qr_image():
    """Loads a real payment_qr.png if present, otherwise draws a clearly
    labelled placeholder so the UI is complete out of the box."""
    if os.path.exists(QR_CODE_PATH):
        return Image.open(QR_CODE_PATH)

    img = Image.new("RGB", (280, 280), "#ffffff")
    draw = ImageDraw.Draw(img)
    draw.rounded_rectangle([10, 10, 270, 270], radius=16, outline="#2563eb", width=6)
    font = ImageFont.load_default()
    lines = ["PLACEHOLDER QR", "GCash / Maya", "Replace payment_qr.png", "next to app.py"]
    y = 105
    for line in lines:
        bbox = draw.textbbox((0, 0), line, font=font)
        w = bbox[2] - bbox[0]
        draw.text(((280 - w) / 2, y), line, fill="#1e293b", font=font)
        y += 20
    return img


# ----------------------------------------------------------------------
# GOOGLE SHEETS DATA LAYER
# ----------------------------------------------------------------------

@st.cache_resource(show_spinner=False)
def get_client():
    try:
        creds_info = dict(st.secrets["gcp_service_account"])
        creds = Credentials.from_service_account_info(creds_info, scopes=SCOPES)
        return gspread.authorize(creds)
    except Exception as exc:
        st.error(
            "Google Sheets authentication failed. Check the "
            "`[gcp_service_account]` block in your secrets.toml. "
            f"Details: {exc}"
        )
        st.stop()


@st.cache_resource(show_spinner=False)
def get_spreadsheet():
    client = get_client()
    try:
        return client.open_by_key(st.secrets["SPREADSHEET_ID"])
    except Exception as exc:
        st.error(
            "Could not open the spreadsheet. Check `SPREADSHEET_ID` in secrets.toml "
            f"and make sure the sheet is shared with the service account email. Details: {exc}"
        )
        st.stop()


@st.cache_resource(show_spinner="Setting up database...")
def init_database():
    """Ensures every worksheet + header row exists, migrates in any newly
    added columns without touching existing data, and seeds demo data the
    first time the workbook is empty. Runs once per app process."""
    sh = get_spreadsheet()
    existing = {ws.title: ws for ws in sh.worksheets()}

    for name, cols in SHEET_COLUMNS.items():
        if name not in existing:
            ws = sh.add_worksheet(title=name, rows=ROW_HINTS.get(name, 1000), cols=len(cols))
            ws.append_row(cols)
            existing[name] = ws
        else:
            ws = existing[name]
            header = ws.row_values(1)
            if not header:
                ws.append_row(cols)
            else:
                # Non-destructive schema migration: append any brand-new
                # columns to the end of the existing header, leaving all
                # existing data and column positions untouched.
                missing = [c for c in cols if c not in header]
                if missing:
                    ws.update(values=[header + missing], range_name="A1")

    emp_ws = existing["Employees"]
    if len(emp_ws.get_all_values()) <= 1:
        emp_ws.append_rows([
            ["EMP-0001", "Manager Admin", "1234", "Manager"],
            ["EMP-0002", "Staff One", "1111", "Employee"],
        ])

    inv_ws = existing["Inventory"]
    if len(inv_ws.get_all_values()) <= 1:
        ts = now_str()
        demo = [
            ["SKU-0001", "Jasmine Rice 25kg", resolve_category("Jasmine Rice 25kg", "Grocery"),
             900, 1150, 40, 10, ts, "FALSE"],
            ["SKU-0002", "Paper Cups (50pcs)", resolve_category("Paper Cups (50pcs)", "Supplies"),
             45, 70, 60, 15, ts, "FALSE"],
            ["SKU-0003", "Bottled Water 500ml", resolve_category("Bottled Water 500ml", "Beverage"),
             8, 20, 120, 24, ts, "FALSE"],
            ["KIOSK-0001", "Chicken Adobo (Rice Meal)", "Chicken", 60, 120, 25, 5, ts, "TRUE"],
            ["KIOSK-0002", "Crispy Pork Sisig", "Pork", 55, 130, 20, 5, ts, "TRUE"],
            ["KIOSK-0003", "Jasmine Rice (Side Cup)", resolve_category("Jasmine Rice (Side Cup)", "Extra"),
             8, 20, 50, 10, ts, "TRUE"],
        ]
        inv_ws.append_rows(demo)
    return True


@st.cache_data(ttl=20, show_spinner=False)
def load_df(sheet_name):
    ws = get_spreadsheet().worksheet(sheet_name)
    records = ws.get_all_records()
    cols = SHEET_COLUMNS[sheet_name]
    df = pd.DataFrame(records) if records else pd.DataFrame(columns=cols)
    for c in cols:
        if c not in df.columns:
            df[c] = ""
    df = df[cols]
    for c in NUMERIC_COLUMNS.get(sheet_name, []):
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)
    for c in df.columns:
        if df[c].dtype == object:
            df[c] = df[c].fillna("").astype(str)
    return df.reset_index(drop=True)


def _live_header(ws, sheet_name):
    return ws.row_values(1) or SHEET_COLUMNS[sheet_name]


def append_row(sheet_name, row_dict):
    ws = get_spreadsheet().worksheet(sheet_name)
    header = _live_header(ws, sheet_name)
    ws.append_row([row_dict.get(c, "") for c in header], value_input_option="USER_ENTERED")
    load_df.clear()


def append_rows(sheet_name, list_of_dicts):
    if not list_of_dicts:
        return
    ws = get_spreadsheet().worksheet(sheet_name)
    header = _live_header(ws, sheet_name)
    rows = [[d.get(c, "") for c in header] for d in list_of_dicts]
    ws.append_rows(rows, value_input_option="USER_ENTERED")
    load_df.clear()


def overwrite_sheet(sheet_name, df):
    """Fully rewrites a sheet using the canonical SHEET_COLUMNS order.
    Used for tables we mutate in place (Inventory, Transactions on void)."""
    ws = get_spreadsheet().worksheet(sheet_name)
    cols = SHEET_COLUMNS[sheet_name]
    df = df[cols]
    values = [cols] + df.astype(str).values.tolist()
    ws.clear()
    ws.update(values=values, range_name="A1")
    load_df.clear()


def log_audit(action, details):
    emp = st.session_state.get("employee")
    append_row("AuditLog", {
        "LogID": new_id("LOG"),
        "DateTime": now_str(),
        "EmployeeName": emp["Name"] if emp else "SYSTEM",
        "Action": action,
        "Details": details,
    })


def find_employee_by_pin(pin):
    df = load_df("Employees")
    match = df[df["PIN"].astype(str).str.strip() == clean_str(pin)]
    if match.empty:
        return None
    row = match.iloc[0]
    return {"EmployeeID": row["EmployeeID"], "Name": row["Name"], "Role": row["Role"]}


def add_item_to_cart(sku, item_name, unit_price, qty_to_add, available_qty):
    """Shared cart-append logic used by both the scanner flow and the Kiosk."""
    already_in_cart = sum(i["Quantity"] for i in st.session_state.cart if i["SKU"] == sku)
    if qty_to_add + already_in_cart > available_qty:
        st.warning(
            f"Only {int(available_qty)} unit(s) of '{item_name}' in stock "
            f"({already_in_cart} already in cart)."
        )
        return False

    for cart_item in st.session_state.cart:
        if cart_item["SKU"] == sku:
            cart_item["Quantity"] += qty_to_add
            return True

    st.session_state.cart.append({
        "SKU": sku, "ItemName": item_name, "Quantity": qty_to_add, "UnitPrice": float(unit_price),
    })
    return True


# ----------------------------------------------------------------------
# SKU SCANNER WIDGET
# ----------------------------------------------------------------------

def reset_scanner(key_prefix):
    """Bumps the widget key's version so the SKU field renders blank on
    the next run (Streamlit won't allow clearing a widget's own
    session_state key after it has been instantiated in the same run)."""
    ver_key = f"{key_prefix}_version"
    st.session_state[ver_key] = st.session_state.get(ver_key, 0) + 1


def sku_scanner_input(key_prefix, label="SKU / Barcode"):
    ver = st.session_state.get(f"{key_prefix}_version", 0)
    text_key = f"{key_prefix}_sku_{ver}"

    with st.expander("📷 Scan with camera", expanded=False):
        if SCANNER_AVAILABLE:
            img_file = st.camera_input("Point at the barcode / QR code", key=f"{key_prefix}_cam_{ver}")
            if img_file is not None:
                try:
                    results = zbar_decode(Image.open(img_file))
                except Exception:
                    results = []
                if results:
                    st.session_state[text_key] = results[0].data.decode("utf-8")
                    st.success(f"Detected: {st.session_state[text_key]}")
                else:
                    st.warning("No barcode/QR detected in that photo — try again or type it below.")
        else:
            st.caption(
                "Camera decoding isn't installed on this server (needs `pyzbar` + the "
                "system `libzbar0` library). A USB/Bluetooth scanner or manual typing "
                "into the field below works either way."
            )

    return st.text_input(
        label, key=text_key,
        placeholder="Scan with a hardware scanner or type here, then press Enter"
    ).strip()


# ----------------------------------------------------------------------
# LOGIN
# ----------------------------------------------------------------------

def show_login():
    st.markdown(
        "<div class='app-banner'><h2 style='margin:0;'>🧾 SKU Tracker</h2>"
        "<p style='margin:0;opacity:.85;'>Sign in with your employee PIN</p></div>",
        unsafe_allow_html=True,
    )
    _, mid, _ = st.columns([1, 1, 1])
    with mid:
        with st.form("login_form", clear_on_submit=True):
            pin = st.text_input("PIN", type="password", placeholder="Enter your PIN")
            submitted = st.form_submit_button("Log In", use_container_width=True)

        if submitted:
            emp = find_employee_by_pin(pin)
            if emp:
                st.session_state.employee = emp
                st.session_state.cart = []
                st.session_state.intake_unlocked = False
                st.session_state.last_receipt = None
                log_audit("Login", f"{emp['Name']} ({emp['Role']}) logged in.")
                st.rerun()
            else:
                st.error("PIN not recognized. Please try again.")

        st.info("Demo PINs → Manager: **1234** · Employee: **1111**", icon="ℹ️")


# ----------------------------------------------------------------------
# TAB 1: POINT OF SALE — 70/30 split
# ----------------------------------------------------------------------

def show_pos_tab():
    st.subheader("Point of Sale")
    left, right = st.columns([7, 3], gap="large")

    with left:
        show_pos_scanner_and_cart()
    with right:
        show_pos_checkout_panel()


def show_pos_scanner_and_cart():
    inv = load_df("Inventory")
    sku = sku_scanner_input("pos")
    match = inv[inv["SKU"].str.lower() == sku.lower()] if sku else pd.DataFrame()

    if sku and match.empty:
        st.warning(f"No inventory item found for SKU '{sku}'.")
    elif not match.empty:
        row = match.iloc[0]
        c1, c2, c3 = st.columns(3)
        c1.metric("Item", row["ItemName"])
        c2.metric("Price", fmt_money(row["UnitPrice"]))
        c3.metric("In Stock", int(row["QuantityOnHand"]))

        qty = st.number_input("Quantity", min_value=1, step=1, value=1, key="pos_qty")
        if st.button("➕ Add to Cart", type="primary", use_container_width=True):
            if add_item_to_cart(row["SKU"], row["ItemName"], row["UnitPrice"], int(qty), row["QuantityOnHand"]):
                reset_scanner("pos")
                st.toast(f"Added {qty} × {row['ItemName']}", icon="✅")
                st.rerun()

    st.divider()
    st.markdown("#### 🛒 Cart")

    if not st.session_state.cart:
        st.caption("Cart is empty. Scan a SKU above or tap an item in the Kiosk tab to begin.")
        return

    for idx, item in enumerate(st.session_state.cart):
        subtotal = item["Quantity"] * item["UnitPrice"]
        c1, c2, c3 = st.columns([3, 2, 1])
        c1.write(f"**{item['ItemName']}**  \n`{item['SKU']}`")
        c2.write(f"{item['Quantity']} × {fmt_money(item['UnitPrice'])} = **{fmt_money(subtotal)}**")
        if c3.button("🗑️", key=f"rm_{idx}"):
            st.session_state.cart.pop(idx)
            st.rerun()


def show_pos_checkout_panel():
    st.markdown("<div class='checkout-card'>", unsafe_allow_html=True)
    st.markdown("#### 🧮 Checkout")

    cart = st.session_state.cart
    total = sum(i["Quantity"] * i["UnitPrice"] for i in cart)
    st.metric("Total", fmt_money(total))

    if not cart:
        st.caption("Add items to begin checkout.")
        st.markdown("</div>", unsafe_allow_html=True)
        return

    payment_method = st.radio("Payment Method", ["Cash", "Online", "Staff Charge"], key="pos_payment")

    reference = ""
    staff_tab_owner = ""

    if payment_method == "Online":
        reference = st.text_input(
            "Transaction Reference Number", key="pos_reference",
            placeholder="Required for bank verification"
        ).strip()
        st.image(get_qr_image(), caption="Scan to pay — GCash / Maya", use_container_width=True)

    elif payment_method == "Staff Charge":
        staff_tab_owner = st.text_input(
            "Employee Name (Staff Tab)", key="pos_staff_tab_owner",
            placeholder="Whose staff tab is this charged to?"
        ).strip()
        st.caption("Stock is deducted normally. The amount is billed to the employee's tab instead of the drawer.")

    if st.button("✅ Complete Sale", type="primary", use_container_width=True):
        if payment_method == "Online" and not reference:
            st.error("A reference number is required for Online payments.")
        elif payment_method == "Staff Charge" and not staff_tab_owner:
            st.error("Enter the employee name this Staff Charge belongs to.")
        else:
            complete_sale(payment_method, reference, staff_tab_owner)

    st.markdown("</div>", unsafe_allow_html=True)


def complete_sale(payment_method, reference, staff_tab_owner):
    inv = load_df("Inventory")
    order_id = new_id("ORD")  # one OrderID groups the whole receipt
    ts = now_str()
    txn_rows = []

    for item in st.session_state.cart:
        idx = inv.index[inv["SKU"] == item["SKU"]]
        if len(idx):
            inv.loc[idx, "QuantityOnHand"] = inv.loc[idx, "QuantityOnHand"] - item["Quantity"]
            inv.loc[idx, "LastUpdated"] = ts
        txn_rows.append({
            "OrderID": order_id, "TransactionID": new_id("TXN"), "DateTime": ts,
            "SKU": item["SKU"], "ItemName": item["ItemName"], "Quantity": item["Quantity"],
            "UnitPrice": item["UnitPrice"], "TotalAmount": round(item["Quantity"] * item["UnitPrice"], 2),
            "PaymentMethod": payment_method,
            "ReferenceNumber": reference if payment_method == "Online" else "",
            "StaffTabOwner": staff_tab_owner if payment_method == "Staff Charge" else "",
            "EmployeeName": st.session_state.employee["Name"], "Status": "Completed",
            "VoidReason": "", "VoidedBy": "", "VoidedDateTime": "",
        })

    overwrite_sheet("Inventory", inv)
    append_rows("Transactions", txn_rows)

    total = sum(r["TotalAmount"] for r in txn_rows)
    detail = f"OrderID={order_id}, {len(txn_rows)} line(s), total={fmt_money(total)}, payment={payment_method}"
    if payment_method == "Online":
        detail += f", ref={reference}"
    elif payment_method == "Staff Charge":
        detail += f", staff_tab_owner={staff_tab_owner}"
    log_audit("Sale Completed", detail)

    # Stage the receipt for the confirmation modal instead of clearing
    # straight to a toast — the dialog is now the checkout's confirmation.
    st.session_state["last_receipt"] = {
        "order_id": order_id, "datetime": ts, "employee": st.session_state.employee["Name"],
        "items": txn_rows, "total": total, "payment_method": payment_method,
        "reference": reference, "staff_tab_owner": staff_tab_owner,
    }
    st.session_state.cart = []
    st.rerun()


@st.dialog("🧾 Receipt")
def receipt_dialog():
    receipt = st.session_state.get("last_receipt")
    if not receipt:
        return

    st.markdown(f"**Order ID:** `{receipt['order_id']}`")
    st.caption(f"{receipt['datetime']} · Served by {receipt['employee']}")
    st.divider()

    for item in receipt["items"]:
        c1, c2 = st.columns([3, 1])
        c1.write(f"{item['ItemName']}  \n{item['Quantity']} × {fmt_money(item['UnitPrice'])}")
        c2.write(f"**{fmt_money(item['TotalAmount'])}**")

    st.divider()
    st.markdown(f"### Total: {fmt_money(receipt['total'])}")

    payment_line = f"**Payment:** {receipt['payment_method']}"
    if receipt["payment_method"] == "Online" and receipt.get("reference"):
        payment_line += f"  ·  Ref: {receipt['reference']}"
    elif receipt["payment_method"] == "Staff Charge" and receipt.get("staff_tab_owner"):
        payment_line += f"  ·  Billed to: {receipt['staff_tab_owner']}"
    st.markdown(payment_line)

    st.divider()
    if st.button("✅ Close & Start Next Sale", type="primary", use_container_width=True, key="close_receipt"):
        st.session_state["last_receipt"] = None
        st.rerun()


# ----------------------------------------------------------------------
# TAB 2: KIOSK — dynamic, database-driven, no-barcode items
# ----------------------------------------------------------------------

def show_kiosk_tab():
    st.subheader("🍽️ Kiosk — Today's Menu")
    st.caption("For viands and no-barcode items. Tap a tile to add it straight to the cart.")

    inv = load_df("Inventory")
    inv["_available"] = inv["AvailableToday"].apply(is_true)
    available = inv[inv["_available"]]

    if available.empty:
        st.info(
            "No items are marked available in the Kiosk today. A manager can enable items "
            "from the **Inventory** tab → *Manage Kiosk Availability*."
        )
        return

    categories = sorted(available["Category"].unique())
    cat_tabs = st.tabs([f"🍱 {c}" for c in categories])

    for cat_tab, category in zip(cat_tabs, categories):
        with cat_tab:
            items = available[available["Category"] == category].reset_index(drop=True)
            cols_per_row = 4
            for start in range(0, len(items), cols_per_row):
                chunk = items.iloc[start:start + cols_per_row]
                cols = st.columns(cols_per_row)
                for col, (_, item) in zip(cols, chunk.iterrows()):
                    with col:
                        out_of_stock = item["QuantityOnHand"] <= 0
                        label = f"{item['ItemName']}\n{fmt_money(item['UnitPrice'])}"
                        if out_of_stock:
                            st.button(f"❌ {item['ItemName']} (Out of stock)",
                                      key=f"kiosk_{item['SKU']}", disabled=True, use_container_width=True)
                        else:
                            if st.button(label, key=f"kiosk_{item['SKU']}", use_container_width=True):
                                if add_item_to_cart(item["SKU"], item["ItemName"], item["UnitPrice"],
                                                     1, item["QuantityOnHand"]):
                                    st.toast(f"Added {item['ItemName']}", icon="✅")
                                    st.rerun()
                            st.caption(f"In stock: {int(item['QuantityOnHand'])}")


# ----------------------------------------------------------------------
# TAB 3: STOCK INTAKE
# ----------------------------------------------------------------------

def show_intake_tab():
    st.subheader("Stock Intake")

    is_manager = st.session_state.employee["Role"] == "Manager"
    if not is_manager and not st.session_state.get("intake_unlocked"):
        st.warning(
            "🔒 Stock Intake is locked. A manager must authorize access before new stock "
            "can be added — this prevents unauthorized ('ghost') inventory being injected "
            "by an employee acting alone."
        )
        pin = st.text_input("Manager PIN", type="password", key="intake_unlock_pin")
        if st.button("🔓 Unlock Stock Intake", type="primary", key="intake_unlock_btn"):
            approver = find_employee_by_pin(pin)
            if approver and approver["Role"] == "Manager":
                st.session_state["intake_unlocked"] = True
                log_audit(
                    "Stock Intake Unlocked",
                    f"Unlocked for {st.session_state.employee['Name']} by manager {approver['Name']}"
                )
                st.success(f"Unlocked by {approver['Name']}.")
                st.rerun()
            else:
                st.error("Invalid Manager PIN.")
        return

    if not is_manager:
        c1, c2 = st.columns([4, 1])
        with c1:
            st.success("🔓 Stock Intake unlocked for this session.")
        with c2:
            if st.button("Lock", key="intake_lock_btn", use_container_width=True):
                st.session_state["intake_unlocked"] = False
                st.rerun()

    inv = load_df("Inventory")

    sku = sku_scanner_input("intake")
    existing = inv[inv["SKU"].str.lower() == sku.lower()] if sku else pd.DataFrame()

    default_name, default_cat, default_cost, default_price, default_reorder = "", "", 0.0, 0.0, 0
    default_kiosk = False
    if not existing.empty:
        r = existing.iloc[0]
        default_name, default_cat = r["ItemName"], r["Category"]
        default_cost, default_price = float(r["UnitCost"]), float(r["UnitPrice"])
        default_reorder = int(r["ReorderLevel"])
        default_kiosk = is_true(r["AvailableToday"])
        st.info(f"Existing item found: **{default_name}** (current stock: {int(r['QuantityOnHand'])})")

    with st.form("intake_form"):
        name = st.text_input("Item Name", value=default_name)
        category = st.text_input(
            "Category (auto-forced to 'Extra' for rice, paper cups, etc.)", value=default_cat
        )
        c1, c2 = st.columns(2)
        cost = c1.number_input("Unit Cost", min_value=0.0, value=default_cost, step=0.5)
        price = c2.number_input("Unit Price", min_value=0.0, value=default_price, step=0.5)
        c3, c4 = st.columns(2)
        reorder = c3.number_input("Reorder Level", min_value=0, value=default_reorder, step=1)
        qty_add = c4.number_input("Quantity to Add", min_value=1, value=1, step=1)
        kiosk_flag = st.checkbox("Show in Kiosk today (viands / no-barcode items)", value=default_kiosk)
        notes = st.text_input("Notes (optional)")
        submitted = st.form_submit_button("➕ Add / Update Stock", use_container_width=True)

    if not submitted:
        return
    if not sku or not name.strip():
        st.error("SKU and Item Name are required.")
        return

    resolved_cat = resolve_category(name, category)
    ts = now_str()
    idx = inv.index[inv["SKU"].str.lower() == sku.lower()]
    kiosk_value = "TRUE" if kiosk_flag else "FALSE"

    if len(idx):
        inv.loc[idx, "ItemName"] = name.strip()
        inv.loc[idx, "Category"] = resolved_cat
        inv.loc[idx, "UnitCost"] = cost
        inv.loc[idx, "UnitPrice"] = price
        inv.loc[idx, "ReorderLevel"] = reorder
        inv.loc[idx, "LastUpdated"] = ts
        inv.loc[idx, "AvailableToday"] = kiosk_value
        inv.loc[idx, "QuantityOnHand"] = inv.loc[idx, "QuantityOnHand"] + qty_add
    else:
        inv = pd.concat([inv, pd.DataFrame([{
            "SKU": sku, "ItemName": name.strip(), "Category": resolved_cat,
            "UnitCost": cost, "UnitPrice": price, "QuantityOnHand": qty_add,
            "ReorderLevel": reorder, "LastUpdated": ts, "AvailableToday": kiosk_value,
        }])], ignore_index=True)

    overwrite_sheet("Inventory", inv)
    append_row("StockIntake", {
        "IntakeID": new_id("INT"), "DateTime": ts, "SKU": sku, "ItemName": name.strip(),
        "QuantityAdded": qty_add, "UnitCost": cost,
        "EmployeeName": st.session_state.employee["Name"], "Notes": notes.strip(),
    })
    log_audit("Stock Intake", f"SKU={sku}, +{qty_add} units, category={resolved_cat}, kiosk={kiosk_value}")

    if resolved_cat == "Extra" and category.strip().lower() != "extra":
        st.info("Category auto-set to **Extra** based on the item-name rules.")

    reset_scanner("intake")
    st.success(f"Stock updated for '{name}' (+{qty_add}).")
    st.rerun()


# ----------------------------------------------------------------------
# TAB 4: INVENTORY
# ----------------------------------------------------------------------

def show_inventory_tab():
    st.subheader("Inventory")

    if st.button("🔄 Refresh", key="inv_refresh"):
        load_df.clear()
        st.rerun()

    inv = load_df("Inventory")
    if inv.empty:
        st.caption("No inventory yet.")
        return

    low_stock = inv[inv["QuantityOnHand"] <= inv["ReorderLevel"]]
    kiosk_count = inv["AvailableToday"].apply(is_true).sum()
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total SKUs", len(inv))
    c2.metric("Low Stock Items", len(low_stock))
    c3.metric("Inventory Value", fmt_money((inv["QuantityOnHand"] * inv["UnitCost"]).sum()))
    c4.metric("Live in Kiosk", int(kiosk_count))

    if len(low_stock):
        with st.expander(f"⚠️ {len(low_stock)} item(s) at or below reorder level", expanded=True):
            st.dataframe(
                low_stock[["SKU", "ItemName", "QuantityOnHand", "ReorderLevel"]],
                use_container_width=True, hide_index=True,
            )

    display_cols = ["SKU", "ItemName", "Category", "UnitCost", "UnitPrice",
                     "QuantityOnHand", "ReorderLevel", "LastUpdated"]

    def highlight_low(row):
        flag = row["QuantityOnHand"] <= row["ReorderLevel"]
        return ["background-color:#ffe3e3" if flag else "" for _ in row]

    st.dataframe(inv[display_cols].style.apply(highlight_low, axis=1),
                 use_container_width=True, hide_index=True)

    with st.expander("🍽️ Manage Kiosk Availability"):
        st.caption("Toggle which items show up as tap-to-add tiles on the Kiosk tab today.")
        kiosk_edit = inv[["SKU", "ItemName", "Category", "AvailableToday"]].copy()
        kiosk_edit["AvailableToday"] = kiosk_edit["AvailableToday"].apply(is_true)
        edited_kiosk = st.data_editor(
            kiosk_edit, hide_index=True, use_container_width=True, key="kiosk_avail_editor",
            column_config={
                "SKU": st.column_config.TextColumn(disabled=True),
                "ItemName": st.column_config.TextColumn(disabled=True),
                "Category": st.column_config.TextColumn(disabled=True),
                "AvailableToday": st.column_config.CheckboxColumn("Show in Kiosk"),
            },
        )
        if st.button("💾 Save Kiosk Availability", type="primary"):
            full_inv = load_df("Inventory").set_index("SKU")
            edit_idx = edited_kiosk.set_index("SKU")
            full_inv.loc[edit_idx.index, "AvailableToday"] = edit_idx["AvailableToday"].map(
                {True: "TRUE", False: "FALSE"}
            )
            overwrite_sheet("Inventory", full_inv.reset_index())
            log_audit("Kiosk Availability Updated", f"{len(edit_idx)} SKU(s) reviewed")
            st.success("Kiosk availability updated.")
            st.rerun()


# ----------------------------------------------------------------------
# TAB 5: SALES HISTORY / VOID
# ----------------------------------------------------------------------

def show_history_tab():
    st.subheader("Sales History")
    txn = load_df("Transactions")

    filter_choice = st.radio("Show", ["Today", "All", "Voided Only"], horizontal=True, key="hist_filter")
    view = txn.copy()
    if filter_choice == "Today":
        view = view[view["DateTime"].astype(str).str.startswith(today_str())]
    elif filter_choice == "Voided Only":
        view = view[view["Status"] == "Voided"]

    if view.empty:
        st.caption("No transactions to show.")
        return

    orders = view.groupby("OrderID").agg(
        DateTime=("DateTime", "first"), Employee=("EmployeeName", "first"),
        Payment=("PaymentMethod", "first"), Reference=("ReferenceNumber", "first"),
        StaffTab=("StaffTabOwner", "first"),
        Items=("SKU", "count"), Total=("TotalAmount", "sum"),
        Status=("Status", lambda s: "Voided" if (s == "Voided").any() else "Completed"),
    ).reset_index().sort_values("DateTime", ascending=False)

    for _, order in orders.iterrows():
        icon = "🟠" if order["Status"] == "Voided" else "🟢"
        header = f"{icon} {order['OrderID']} · {order['DateTime']} · {fmt_money(order['Total'])} · {order['Payment']}"
        with st.expander(header):
            lines = view[view["OrderID"] == order["OrderID"]]
            st.dataframe(
                lines[["SKU", "ItemName", "Quantity", "UnitPrice", "TotalAmount", "Status"]],
                use_container_width=True, hide_index=True,
            )
            extra = ""
            if order["Reference"]:
                extra = f" · Ref: {order['Reference']}"
            elif order["StaffTab"]:
                extra = f" · Staff Tab: {order['StaffTab']}"
            st.caption(f"Sold by {order['Employee']} · Payment: {order['Payment']}{extra}")

            if order["Status"] == "Completed":
                if st.button("🚫 Void This Order", key=f"void_{order['OrderID']}"):
                    st.session_state["void_target"] = order["OrderID"]
                    st.rerun()
            else:
                voided_row = lines.iloc[0]
                st.warning(f"Voided by {voided_row['VoidedBy']} — reason: {voided_row['VoidReason']}")

    if st.session_state.get("void_target"):
        void_order_dialog(st.session_state["void_target"])


@st.dialog("Void Order")
def void_order_dialog(order_id):
    st.write(f"You are voiding order **{order_id}**. This cannot be undone and will restock every item on it.")
    reason = st.text_area("Reason for voiding")
    manager_pin = st.text_input("Manager PIN to authorize", type="password")

    c1, c2 = st.columns(2)
    if c1.button("Cancel", use_container_width=True):
        st.session_state["void_target"] = None
        st.rerun()

    if c2.button("Confirm Void", type="primary", use_container_width=True):
        if not reason.strip():
            st.error("A reason is required.")
            return
        approver = find_employee_by_pin(manager_pin)
        if not approver or approver["Role"] != "Manager":
            st.error("Invalid Manager PIN. Voiding requires manager authorization.")
            return
        void_order(order_id, reason.strip(), approver)
        st.session_state["void_target"] = None
        st.rerun()


def void_order(order_id, reason, approver):
    """Employees can never delete a record, only void it — and it stays
    permanently flagged for the manager's reporting."""
    txn = load_df("Transactions")
    idx = txn.index[txn["OrderID"] == order_id]
    ts = now_str()
    txn.loc[idx, "Status"] = "Voided"
    txn.loc[idx, "VoidReason"] = reason
    txn.loc[idx, "VoidedBy"] = approver["Name"]
    txn.loc[idx, "VoidedDateTime"] = ts
    overwrite_sheet("Transactions", txn)

    inv = load_df("Inventory")
    for _, row in txn.loc[idx].iterrows():
        inv_idx = inv.index[inv["SKU"] == row["SKU"]]
        if len(inv_idx):
            inv.loc[inv_idx, "QuantityOnHand"] = inv.loc[inv_idx, "QuantityOnHand"] + row["Quantity"]
    overwrite_sheet("Inventory", inv)

    log_audit("Order Voided", f"OrderID={order_id}, approved by {approver['Name']}, reason='{reason}'")
    st.success(f"Order {order_id} voided and restocked.")


# ----------------------------------------------------------------------
# TAB 6: RECONCILIATION — Revenue + Opening/Closing blind counts
# ----------------------------------------------------------------------

def show_reconciliation_tab():
    st.subheader("End-of-Day Reconciliation & Shift Counts")
    sub_tabs = st.tabs(["💰 Revenue", "🌅 Opening Count", "🌙 Closing Count (Blind)"])
    with sub_tabs[0]:
        show_revenue_reconciliation()
    with sub_tabs[1]:
        show_opening_count()
    with sub_tabs[2]:
        show_closing_count()


def show_revenue_reconciliation():
    txn = load_df("Transactions")
    today_txn = txn[(txn["DateTime"].astype(str).str.startswith(today_str())) & (txn["Status"] == "Completed")]
    expected_cash = today_txn[today_txn["PaymentMethod"] == "Cash"]["TotalAmount"].sum()
    expected_online = today_txn[today_txn["PaymentMethod"] == "Online"]["TotalAmount"].sum()
    staff_tab_total = today_txn[today_txn["PaymentMethod"] == "Staff Charge"]["TotalAmount"].sum()

    st.markdown("##### 💰 Drawer & Online Revenue")
    c1, c2 = st.columns(2)
    with c1:
        st.metric("Expected Cash", fmt_money(expected_cash))
        actual_cash = st.number_input("Actual Cash Counted", min_value=0.0, step=1.0, key="recon_cash")
    with c2:
        st.metric("Expected Online", fmt_money(expected_online))
        actual_online = st.number_input("Actual Online (bank-verified)", min_value=0.0, step=1.0, key="recon_online")

    cash_var = round(actual_cash - expected_cash, 2)
    online_var = round(actual_online - expected_online, 2)
    c3, c4 = st.columns(2)
    c3.metric("Cash Variance", fmt_money(cash_var), delta=f"{cash_var:+.2f}",
              delta_color="off" if cash_var == 0 else "inverse")
    c4.metric("Online Variance", fmt_money(online_var), delta=f"{online_var:+.2f}",
              delta_color="off" if online_var == 0 else "inverse")
    if cash_var != 0 or online_var != 0:
        st.warning("⚠️ Revenue variance detected — double-check the drawer / bank statement.")
    else:
        st.success("Drawer and online totals match expected revenue.")

    if staff_tab_total > 0:
        st.divider()
        st.markdown("##### 🧾 Staff Tab (not part of drawer cash)")
        st.metric("Staff Tab Total Today", fmt_money(staff_tab_total))
        by_owner = (
            today_txn[today_txn["PaymentMethod"] == "Staff Charge"]
            .groupby("StaffTabOwner")["TotalAmount"].sum().reset_index()
            .rename(columns={"StaffTabOwner": "Employee", "TotalAmount": "Amount Owed"})
        )
        st.dataframe(by_owner, hide_index=True, use_container_width=True)

    st.divider()
    notes = st.text_area("Notes", key="recon_notes")
    is_manager = st.session_state.employee["Role"] == "Manager"
    if st.button("💾 Save Revenue Reconciliation", use_container_width=True, disabled=not is_manager,
                 type="primary"):
        save_reconciliation(expected_cash, actual_cash, cash_var, expected_online, actual_online, online_var, notes)
    if not is_manager:
        st.caption("🔒 Saving requires a Manager login.")


def save_reconciliation(expected_cash, actual_cash, cash_var, expected_online, actual_online, online_var, notes):
    append_row("Reconciliation", {
        "ReconID": new_id("REC"), "Date": today_str(),
        "ExpectedCash": expected_cash, "ActualCash": actual_cash, "CashVariance": cash_var,
        "ExpectedOnline": expected_online, "ActualOnline": actual_online, "OnlineVariance": online_var,
        "EmployeeName": st.session_state.employee["Name"], "Notes": notes.strip(),
    })
    log_audit("End-of-Day Revenue Reconciliation Saved", f"CashVar={cash_var:+.2f}, OnlineVar={online_var:+.2f}")
    st.success("Revenue reconciliation saved.")


def get_latest_closing_counts():
    """Most recent Closing ShiftCounts entry per SKU, used to cross-check
    the next Opening Count."""
    sc = load_df("ShiftCounts")
    closing = sc[sc["ShiftType"] == "Closing"]
    if closing.empty:
        return {}
    closing = closing.sort_values("DateTime")
    latest = closing.groupby("SKU").last()
    return latest["ActualQty"].to_dict()


def show_closing_count():
    st.markdown(
        "Count the physical stock **without looking at system numbers**, then submit. "
        "This is a blind count — variances are recorded for management review and are "
        "not shown here so counts can't be adjusted to match what's 'expected'."
    )
    inv = load_df("Inventory")
    blind_df = inv[["SKU", "ItemName"]].copy()
    blind_df["ActualQty"] = 0

    edited = st.data_editor(
        blind_df, hide_index=True, use_container_width=True, key="closing_editor",
        column_config={
            "SKU": st.column_config.TextColumn(disabled=True),
            "ItemName": st.column_config.TextColumn(disabled=True),
            "ActualQty": st.column_config.NumberColumn("Physical Count", min_value=0, step=1),
        },
    )
    notes = st.text_area("Shift notes (optional)", key="closing_notes")

    st.caption("🔒 Applying this count to live inventory requires manager authorization.")
    if st.button("Submit Closing Count for Authorization", type="primary",
                 use_container_width=True, key="submit_closing"):
        st.session_state["pending_closing_count"] = {"df": edited.copy(), "notes": notes}
        st.rerun()

    if st.session_state.get("pending_closing_count"):
        closing_count_auth_dialog()


@st.dialog("Manager Authorization — Closing Count")
def closing_count_auth_dialog():
    st.write(
        "This blind count is staged but **not yet applied**. A manager must enter their PIN "
        "to authorize writing these physical counts to live inventory."
    )
    manager_pin = st.text_input("Manager PIN", type="password", key="closing_auth_pin")

    c1, c2 = st.columns(2)
    if c1.button("Cancel", use_container_width=True, key="closing_auth_cancel"):
        st.session_state["pending_closing_count"] = None
        st.rerun()

    if c2.button("✅ Confirm & Apply", type="primary", use_container_width=True, key="closing_auth_confirm"):
        approver = find_employee_by_pin(manager_pin)
        if not approver or approver["Role"] != "Manager":
            st.error("Invalid Manager PIN. Applying a closing count requires manager authorization.")
            return
        pending = st.session_state["pending_closing_count"]
        inv = load_df("Inventory")
        submit_closing_count(inv, pending["df"], pending["notes"], approver)
        st.session_state["pending_closing_count"] = None
        st.rerun()


def submit_closing_count(inv, counted_df, notes, approver):
    ts = now_str()
    rows, variance_rows = [], []

    for _, r in counted_df.iterrows():
        sku = r["SKU"]
        expected = float(inv.loc[inv["SKU"] == sku, "QuantityOnHand"].iloc[0])
        actual = float(r["ActualQty"])
        variance = actual - expected
        rows.append({
            "ShiftCountID": new_id("SC"), "DateTime": ts, "ShiftType": "Closing",
            "SKU": sku, "ItemName": r["ItemName"], "ExpectedQty": expected,
            "ActualQty": actual, "Variance": variance,
            "EmployeeName": st.session_state.employee["Name"], "AuthorizedBy": approver["Name"],
            "Notes": notes.strip(),
        })
        if variance != 0:
            variance_rows.append((sku, r["ItemName"], expected, actual, variance))

    append_rows("ShiftCounts", rows)

    inv2 = load_df("Inventory")
    for r in rows:
        idx = inv2.index[inv2["SKU"] == r["SKU"]]
        if len(idx):
            inv2.loc[idx, "QuantityOnHand"] = r["ActualQty"]
            inv2.loc[idx, "LastUpdated"] = ts
    overwrite_sheet("Inventory", inv2)

    summary = "; ".join(f"{s}:{v:+.0f}" for s, _, _, _, v in variance_rows) if variance_rows else "no variances"
    log_audit(
        "Closing Count Applied (Blind)",
        f"Counted by {st.session_state.employee['Name']}, authorized by {approver['Name']}, "
        f"{len(rows)} SKU(s), {len(variance_rows)} variance(s): {summary}"
    )

    is_manager = st.session_state.employee["Role"] == "Manager"
    if is_manager:
        if variance_rows:
            st.error(f"🚨 {len(variance_rows)} SKU(s) had a variance (manager view):")
            st.dataframe(
                pd.DataFrame(variance_rows, columns=["SKU", "ItemName", "Expected", "Actual", "Variance"]),
                hide_index=True, use_container_width=True,
            )
        else:
            st.success("Closing count matches system records exactly.")
    else:
        st.success(
            f"✅ Closing count applied (authorized by {approver['Name']}). "
            "Any differences are recorded for management review."
        )


def show_opening_count():
    st.markdown(
        "Count the physical stock **before opening**, without looking at last night's numbers. "
        "The system cross-references your count against last night's closing count and "
        "**instantly flags any mismatch** — a possible sign of cross-shift loss or theft."
    )
    inv = load_df("Inventory")
    open_df = inv[["SKU", "ItemName"]].copy()
    open_df["ActualQty"] = 0

    edited = st.data_editor(
        open_df, hide_index=True, use_container_width=True, key="opening_editor",
        column_config={
            "SKU": st.column_config.TextColumn(disabled=True),
            "ItemName": st.column_config.TextColumn(disabled=True),
            "ActualQty": st.column_config.NumberColumn("Physical Count", min_value=0, step=1),
        },
    )
    notes = st.text_area("Shift notes (optional)", key="opening_notes")

    st.caption("🔒 Applying this count to live inventory requires manager authorization.")
    if st.button("Submit Opening Count for Authorization", type="primary",
                 use_container_width=True, key="submit_opening"):
        st.session_state["pending_opening_count"] = {"df": edited.copy(), "notes": notes}
        st.rerun()

    if st.session_state.get("pending_opening_count"):
        opening_count_auth_dialog()


@st.dialog("Manager Authorization — Opening Count")
def opening_count_auth_dialog():
    st.write(
        "This opening count is staged but **not yet applied**. A manager must enter their PIN "
        "to authorize writing these physical counts to live inventory."
    )
    manager_pin = st.text_input("Manager PIN", type="password", key="opening_auth_pin")

    c1, c2 = st.columns(2)
    if c1.button("Cancel", use_container_width=True, key="opening_auth_cancel"):
        st.session_state["pending_opening_count"] = None
        st.rerun()

    if c2.button("✅ Confirm & Apply", type="primary", use_container_width=True, key="opening_auth_confirm"):
        approver = find_employee_by_pin(manager_pin)
        if not approver or approver["Role"] != "Manager":
            st.error("Invalid Manager PIN. Applying an opening count requires manager authorization.")
            return
        pending = st.session_state["pending_opening_count"]
        submit_opening_count(pending["df"], pending["notes"], approver)
        st.session_state["pending_opening_count"] = None
        st.rerun()


def submit_opening_count(counted_df, notes, approver):
    ts = now_str()
    prev_closing = get_latest_closing_counts()
    rows, flags = [], []

    for _, r in counted_df.iterrows():
        sku = r["SKU"]
        actual = float(r["ActualQty"])
        prev = prev_closing.get(sku)
        has_baseline = prev is not None
        variance = (actual - prev) if has_baseline else 0.0
        rows.append({
            "ShiftCountID": new_id("SC"), "DateTime": ts, "ShiftType": "Opening",
            "SKU": sku, "ItemName": r["ItemName"],
            "ExpectedQty": prev if has_baseline else "",
            "ActualQty": actual, "Variance": variance if has_baseline else "",
            "EmployeeName": st.session_state.employee["Name"], "AuthorizedBy": approver["Name"],
            "Notes": notes.strip(),
        })
        if has_baseline and variance != 0:
            flags.append((sku, r["ItemName"], prev, actual, variance))

    append_rows("ShiftCounts", rows)

    inv2 = load_df("Inventory")
    for r in rows:
        idx = inv2.index[inv2["SKU"] == r["SKU"]]
        if len(idx):
            inv2.loc[idx, "QuantityOnHand"] = r["ActualQty"]
            inv2.loc[idx, "LastUpdated"] = ts
    overwrite_sheet("Inventory", inv2)

    if flags:
        summary = "; ".join(f"{s}: closed {p:.0f} -> opened {a:.0f} (Δ{v:+.0f})" for s, _, p, a, v in flags)
        log_audit(
            "🚨 Cross-Shift Discrepancy Flagged",
            f"Counted by {st.session_state.employee['Name']}, authorized by {approver['Name']}: {summary}"
        )
        st.error(f"🚨 {len(flags)} SKU(s) don't match last night's closing count:")
        st.dataframe(
            pd.DataFrame(flags, columns=["SKU", "ItemName", "LastClosing", "OpenedAt", "Variance"]),
            hide_index=True, use_container_width=True,
        )
    else:
        log_audit(
            "Opening Count Applied",
            f"Counted by {st.session_state.employee['Name']}, authorized by {approver['Name']}, "
            f"{len(rows)} SKU(s), matches last closing count."
        )
        st.success(f"✅ Opening count applied (authorized by {approver['Name']}). No discrepancies.")


# ----------------------------------------------------------------------
# TAB 7: AUDIT LOG (Manager only) + high-risk dashboard
# ----------------------------------------------------------------------

def show_manager_dashboard():
    """Quick-glance summary of high-risk events: voided orders, stock-count
    variances, and cross-shift discrepancies from the last 7 days."""
    st.markdown("#### 🚨 High-Risk Events — Last 7 Days")

    cutoff = (dt.datetime.now() - dt.timedelta(days=7)).strftime("%Y-%m-%d %H:%M:%S")

    txn = load_df("Transactions")
    voided = txn[(txn["Status"] == "Voided") & (txn["VoidedDateTime"] >= cutoff)]
    voided_orders = voided.drop_duplicates(subset="OrderID").sort_values("VoidedDateTime", ascending=False)

    sc = load_df("ShiftCounts")
    sc_recent = sc[(sc["DateTime"] >= cutoff) & (sc["Variance"] != 0)]
    closing_variances = sc_recent[sc_recent["ShiftType"] == "Closing"].sort_values("DateTime", ascending=False)
    opening_flags = sc_recent[sc_recent["ShiftType"] == "Opening"].sort_values("DateTime", ascending=False)

    c1, c2, c3 = st.columns(3)
    c1.metric("Voided Orders", len(voided_orders),
              delta="review" if len(voided_orders) else "none", delta_color="inverse")
    c2.metric("Closing Count Variances", len(closing_variances),
              delta="review" if len(closing_variances) else "none", delta_color="inverse")
    c3.metric("Cross-Shift Discrepancies", len(opening_flags),
              delta="review" if len(opening_flags) else "none", delta_color="inverse")

    if voided_orders.empty and closing_variances.empty and opening_flags.empty:
        st.success("No high-risk events in the last 7 days. ✅")
        st.divider()
        return

    if len(voided_orders):
        with st.expander(f"🟠 {len(voided_orders)} Voided Order(s)", expanded=False):
            st.dataframe(
                voided_orders[["OrderID", "DateTime", "EmployeeName", "TotalAmount",
                                "VoidReason", "VoidedBy", "VoidedDateTime"]],
                hide_index=True, use_container_width=True,
            )

    if len(closing_variances):
        with st.expander(f"📦 {len(closing_variances)} Closing Count Variance(s)", expanded=False):
            st.dataframe(
                closing_variances[["DateTime", "SKU", "ItemName", "ExpectedQty", "ActualQty",
                                    "Variance", "EmployeeName", "AuthorizedBy"]],
                hide_index=True, use_container_width=True,
            )

    if len(opening_flags):
        with st.expander(f"🌅 {len(opening_flags)} Cross-Shift Discrepanc(y/ies)", expanded=True):
            st.dataframe(
                opening_flags[["DateTime", "SKU", "ItemName", "ExpectedQty", "ActualQty",
                                "Variance", "EmployeeName", "AuthorizedBy"]],
                hide_index=True, use_container_width=True,
            )

    st.divider()


def show_audit_tab():
    show_manager_dashboard()

    st.subheader("Full Audit Log")
    df = load_df("AuditLog").sort_values("DateTime", ascending=False)

    search = st.text_input("Filter by employee / action / details", key="audit_search")
    if search:
        mask = df.apply(lambda r: search.lower() in " ".join(r.astype(str)).lower(), axis=1)
        df = df[mask]

    st.dataframe(df, use_container_width=True, hide_index=True)


# ----------------------------------------------------------------------
# MAIN
# ----------------------------------------------------------------------

def main():
    init_database()
    st.session_state.setdefault("employee", None)
    st.session_state.setdefault("cart", [])
    st.session_state.setdefault("void_target", None)
    st.session_state.setdefault("pending_closing_count", None)
    st.session_state.setdefault("pending_opening_count", None)
    st.session_state.setdefault("intake_unlocked", False)
    st.session_state.setdefault("last_receipt", None)

    if not st.session_state.employee:
        show_login()
        return

    emp = st.session_state.employee
    col1, col2 = st.columns([5, 1])
    with col1:
        st.markdown(f"**{emp['Name']}** · {emp['Role']}")
    with col2:
        if st.button("Logout", use_container_width=True):
            log_audit("Logout", f"{emp['Name']} logged out.")
            st.session_state.employee = None
            st.session_state.cart = []
            st.session_state.intake_unlocked = False
            st.session_state.last_receipt = None
            st.rerun()

    # Pops up right after a checkout, on top of whichever tab is active.
    if st.session_state.get("last_receipt"):
        receipt_dialog()

    tab_names = ["🧾 Sell", "🍽️ Kiosk", "📦 Stock In", "📊 Inventory", "🧮 History", "📋 Reconcile"]
    if emp["Role"] == "Manager":
        tab_names.append("🔒 Audit Log")
    tabs = st.tabs(tab_names)

    with tabs[0]:
        show_pos_tab()
    with tabs[1]:
        show_kiosk_tab()
    with tabs[2]:
        show_intake_tab()
    with tabs[3]:
        show_inventory_tab()
    with tabs[4]:
        show_history_tab()
    with tabs[5]:
        show_reconciliation_tab()
    if emp["Role"] == "Manager":
        with tabs[6]:
            show_audit_tab()


if __name__ == "__main__":
    main()