"""
SKU Storage & Sales Tracker
============================
A single-file desktop application (Tkinter + pandas/openpyxl) that uses a
local Excel workbook (sku_tracker.xlsx) as its database.

Run with:  python main.py
"""

import os
import uuid
import datetime as dt
import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog

# ----------------------------------------------------------------------
# CONFIG
# ----------------------------------------------------------------------

DB_FILE = "sku_tracker.xlsx"

# Requirement 2: items in this list are FORCED into the "Extra" category
# regardless of what category text is typed in. Matching is done as a
# case-insensitive substring match against the item name, so "Jasmine Rice"
# still matches "rice". Add / edit entries here as needed.
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

# Requirement 1: exact column headers for each sheet of the Excel workbook.
SHEET_COLUMNS = {
    "Employees": ["EmployeeID", "Name", "PIN", "Role"],
    "Inventory": ["SKU", "ItemName", "Category", "UnitCost", "UnitPrice",
                  "QuantityOnHand", "ReorderLevel", "LastUpdated"],
    "Transactions": ["TransactionID", "DateTime", "SKU", "ItemName", "Quantity",
                      "UnitPrice", "TotalAmount", "PaymentMethod", "ReferenceNumber",
                      "EmployeeName", "Status", "VoidReason", "VoidedBy", "VoidedDateTime"],
    "StockIntake": ["IntakeID", "DateTime", "SKU", "ItemName", "QuantityAdded",
                     "UnitCost", "EmployeeName", "Notes"],
    "AuditLog": ["LogID", "DateTime", "EmployeeName", "Action", "Details"],
    "Reconciliation": ["ReconID", "Date", "ExpectedCash", "ActualCash", "CashVariance",
                        "ExpectedOnline", "ActualOnline", "OnlineVariance",
                        "EmployeeName", "Notes"],
}

NUMERIC_COLUMNS = {
    "Inventory": ["UnitCost", "UnitPrice", "QuantityOnHand", "ReorderLevel"],
    "Transactions": ["Quantity", "UnitPrice", "TotalAmount"],
    "StockIntake": ["QuantityAdded", "UnitCost"],
    "Reconciliation": ["ExpectedCash", "ActualCash", "CashVariance",
                        "ExpectedOnline", "ActualOnline", "OnlineVariance"],
}

STRING_ID_COLUMNS = {
    "Employees": ["EmployeeID", "PIN"],
    "Inventory": ["SKU"],
    "Transactions": ["TransactionID", "SKU", "ReferenceNumber"],
    "StockIntake": ["IntakeID", "SKU"],
}


def now_str():
    return dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def today_str():
    return dt.date.today().strftime("%Y-%m-%d")


def new_id(prefix):
    return f"{prefix}-{dt.datetime.now().strftime('%Y%m%d%H%M%S')}-{uuid.uuid4().hex[:5].upper()}"


def clean_str(x):
    """Normalize a cell value to a clean string (fixes '1234.0' from Excel)."""
    if pd.isna(x):
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
    """Requirement 2: force override categories for known 'Extra' items."""
    name_lower = (item_name or "").strip().lower()
    for keyword, forced_cat in CATEGORY_OVERRIDES.items():
        if keyword in name_lower:
            return forced_cat
    requested_category = (requested_category or "").strip()
    return requested_category if requested_category else "General"


# ----------------------------------------------------------------------
# DATA LAYER
# ----------------------------------------------------------------------

class ExcelDB:
    """Loads and persists every sheet of the Excel workbook."""

    def __init__(self, path):
        self.path = path
        self.sheets = {}
        self._load_or_create()

    def _load_or_create(self):
        if not os.path.exists(self.path):
            self._create_default_workbook()
        else:
            try:
                raw = pd.read_excel(self.path, sheet_name=None, engine="openpyxl")
            except Exception as exc:
                raise RuntimeError(
                    f"Could not read {self.path}. Is it open in Excel? Details: {exc}"
                )
            for name, columns in SHEET_COLUMNS.items():
                df = raw.get(name, pd.DataFrame(columns=columns)).copy()
                for c in columns:
                    if c not in df.columns:
                        df[c] = ""
                df = df[columns]
                self.sheets[name] = self._clean_df(name, df)

    def _clean_df(self, name, df):
        df = df.copy()
        for c in STRING_ID_COLUMNS.get(name, []):
            df[c] = df[c].apply(clean_str)
        for c in NUMERIC_COLUMNS.get(name, []):
            df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)
        for c in df.columns:
            if df[c].dtype == object:
                df[c] = df[c].fillna("")
        return df.reset_index(drop=True)

    def _create_default_workbook(self):
        self.sheets = {name: pd.DataFrame(columns=cols) for name, cols in SHEET_COLUMNS.items()}

        # Seed one manager account so the app is usable on first launch.
        self.sheets["Employees"] = pd.DataFrame([
            {"EmployeeID": "EMP-0001", "Name": "Manager Admin", "PIN": "1234", "Role": "Manager"},
            {"EmployeeID": "EMP-0002", "Name": "Staff One", "PIN": "1111", "Role": "Employee"},
        ])

        # Seed a few demo inventory rows, showing the category auto-override.
        demo_rows = [
            {"SKU": "SKU-0001", "ItemName": "Jasmine Rice 25kg", "Category": "Grocery",
             "UnitCost": 900, "UnitPrice": 1150, "QuantityOnHand": 40, "ReorderLevel": 10,
             "LastUpdated": now_str()},
            {"SKU": "SKU-0002", "ItemName": "Paper Cups (50pcs)", "Category": "Supplies",
             "UnitCost": 45, "UnitPrice": 70, "QuantityOnHand": 60, "ReorderLevel": 15,
             "LastUpdated": now_str()},
            {"SKU": "SKU-0003", "ItemName": "Bottled Water 500ml", "Category": "Beverage",
             "UnitCost": 8, "UnitPrice": 20, "QuantityOnHand": 120, "ReorderLevel": 24,
             "LastUpdated": now_str()},
        ]
        for row in demo_rows:
            row["Category"] = resolve_category(row["ItemName"], row["Category"])
        self.sheets["Inventory"] = pd.DataFrame(demo_rows)

        for name in self.sheets:
            self.sheets[name] = self._clean_df(name, self.sheets[name])

        self.save()

    def save(self):
        try:
            with pd.ExcelWriter(self.path, engine="openpyxl") as writer:
                for name, columns in SHEET_COLUMNS.items():
                    df = self.sheets[name][columns]
                    df.to_excel(writer, sheet_name=name, index=False)
        except PermissionError:
            raise RuntimeError(
                f"Cannot save to {self.path}. Please close the file if it is open "
                "in Excel, then try again."
            )

    # convenience accessors -------------------------------------------------
    def df(self, name):
        return self.sheets[name]

    def set_df(self, name, df):
        self.sheets[name] = self._clean_df(name, df)


# ----------------------------------------------------------------------
# MAIN APPLICATION
# ----------------------------------------------------------------------

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("SKU Storage & Sales Tracker")
        self.geometry("1150x720")
        self.minsize(1000, 650)

        try:
            self.db = ExcelDB(DB_FILE)
        except RuntimeError as exc:
            messagebox.showerror("Database Error", str(exc))
            self.destroy()
            return

        self.current_employee = None  # dict: EmployeeID, Name, Role
        self.cart = []  # list of dicts for the current sale

        self.container = ttk.Frame(self)
        self.container.pack(fill="both", expand=True)

        self.show_login()

    # ------------------------------------------------------------------
    # persistence / audit helpers
    # ------------------------------------------------------------------
    def save_db(self):
        try:
            self.db.save()
        except RuntimeError as exc:
            messagebox.showerror("Save Error", str(exc))

    def log_action(self, action, details):
        df = self.db.df("AuditLog")
        row = {
            "LogID": new_id("LOG"),
            "DateTime": now_str(),
            "EmployeeName": self.current_employee["Name"] if self.current_employee else "SYSTEM",
            "Action": action,
            "Details": details,
        }
        df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
        self.db.set_df("AuditLog", df)

    def find_employee_by_pin(self, pin):
        df = self.db.df("Employees")
        match = df[df["PIN"] == clean_str(pin)]
        if match.empty:
            return None
        row = match.iloc[0]
        return {"EmployeeID": row["EmployeeID"], "Name": row["Name"], "Role": row["Role"]}

    # ------------------------------------------------------------------
    # LOGIN SCREEN
    # ------------------------------------------------------------------
    def clear_container(self):
        for widget in self.container.winfo_children():
            widget.destroy()

    def show_login(self):
        self.current_employee = None
        self.cart = []
        self.clear_container()

        wrap = ttk.Frame(self.container)
        wrap.place(relx=0.5, rely=0.5, anchor="center")

        ttk.Label(wrap, text="SKU Storage & Sales Tracker", font=("Segoe UI", 18, "bold")).pack(pady=(0, 20))
        ttk.Label(wrap, text="Enter your Employee PIN to log in", font=("Segoe UI", 11)).pack(pady=(0, 10))

        pin_var = tk.StringVar()
        entry = ttk.Entry(wrap, textvariable=pin_var, show="*", font=("Segoe UI", 14), justify="center", width=15)
        entry.pack(pady=5)
        entry.focus_set()

        def attempt_login(event=None):
            emp = self.find_employee_by_pin(pin_var.get())
            if not emp:
                messagebox.showerror("Login Failed", "PIN not recognized. Please try again.")
                pin_var.set("")
                return
            self.current_employee = emp
            self.log_action("Login", f"{emp['Name']} ({emp['Role']}) logged in.")
            self.save_db()
            self.build_main_ui()

        entry.bind("<Return>", attempt_login)
        ttk.Button(wrap, text="Login", command=attempt_login).pack(pady=15)
        ttk.Label(wrap, text="Default demo PINs -> Manager: 1234   Employee: 1111",
                  foreground="gray").pack(pady=(10, 0))

    # ------------------------------------------------------------------
    # MAIN UI (after login)
    # ------------------------------------------------------------------
    def build_main_ui(self):
        self.clear_container()

        topbar = ttk.Frame(self.container, padding=8)
        topbar.pack(fill="x")
        ttk.Label(
            topbar,
            text=f"Logged in as: {self.current_employee['Name']} ({self.current_employee['Role']})",
            font=("Segoe UI", 11, "bold"),
        ).pack(side="left")
        ttk.Button(topbar, text="Logout", command=self.show_login).pack(side="right")

        notebook = ttk.Notebook(self.container)
        notebook.pack(fill="both", expand=True, padx=8, pady=8)

        pos_tab = ttk.Frame(notebook)
        intake_tab = ttk.Frame(notebook)
        inventory_tab = ttk.Frame(notebook)
        history_tab = ttk.Frame(notebook)
        recon_tab = ttk.Frame(notebook)

        notebook.add(pos_tab, text="Point of Sale")
        notebook.add(intake_tab, text="Stock Intake")
        notebook.add(inventory_tab, text="Inventory")
        notebook.add(history_tab, text="Sales History / Void")
        notebook.add(recon_tab, text="End-of-Day Reconciliation")

        if self.current_employee["Role"] == "Manager":
            audit_tab = ttk.Frame(notebook)
            notebook.add(audit_tab, text="Audit Log")
            self.build_audit_tab(audit_tab)

        self.build_pos_tab(pos_tab)
        self.build_intake_tab(intake_tab)
        self.build_inventory_tab(inventory_tab)
        self.build_history_tab(history_tab)
        self.build_reconciliation_tab(recon_tab)

    # ------------------------------------------------------------------
    # TAB 1: POINT OF SALE  (Requirements 5, 6)
    # ------------------------------------------------------------------
    def build_pos_tab(self, parent):
        left = ttk.Frame(parent, padding=10)
        left.pack(side="left", fill="y")
        right = ttk.Frame(parent, padding=10)
        right.pack(side="left", fill="both", expand=True)

        ttk.Label(left, text="Scan / Enter SKU:", font=("Segoe UI", 10, "bold")).grid(row=0, column=0, sticky="w")
        self.pos_sku_var = tk.StringVar()
        sku_entry = ttk.Entry(left, textvariable=self.pos_sku_var, width=25)
        sku_entry.grid(row=1, column=0, pady=(0, 10))
        sku_entry.bind("<Return>", self.pos_lookup_sku)
        sku_entry.focus_set()
        self.pos_sku_entry = sku_entry

        self.pos_item_var = tk.StringVar(value="-")
        self.pos_price_var = tk.StringVar(value="-")
        self.pos_stock_var = tk.StringVar(value="-")

        ttk.Label(left, text="Item:").grid(row=2, column=0, sticky="w")
        ttk.Label(left, textvariable=self.pos_item_var, font=("Segoe UI", 10, "bold")).grid(row=3, column=0, sticky="w")
        ttk.Label(left, text="Unit Price:").grid(row=4, column=0, sticky="w", pady=(8, 0))
        ttk.Label(left, textvariable=self.pos_price_var).grid(row=5, column=0, sticky="w")
        ttk.Label(left, text="In Stock:").grid(row=6, column=0, sticky="w", pady=(8, 0))
        ttk.Label(left, textvariable=self.pos_stock_var).grid(row=7, column=0, sticky="w")

        ttk.Label(left, text="Quantity:").grid(row=8, column=0, sticky="w", pady=(10, 0))
        self.pos_qty_var = tk.IntVar(value=1)
        ttk.Spinbox(left, from_=1, to=9999, textvariable=self.pos_qty_var, width=10).grid(row=9, column=0, sticky="w")

        ttk.Button(left, text="Add to Cart", command=self.pos_add_to_cart).grid(row=10, column=0, pady=15, sticky="ew")

        # cart -------------------------------------------------------------
        columns = ("SKU", "Item", "Qty", "Price", "Subtotal")
        self.pos_cart_tree = ttk.Treeview(right, columns=columns, show="headings", height=12)
        for c in columns:
            self.pos_cart_tree.heading(c, text=c)
            self.pos_cart_tree.column(c, width=110)
        self.pos_cart_tree.pack(fill="both", expand=True)

        ttk.Button(right, text="Remove Selected Line", command=self.pos_remove_line).pack(anchor="w", pady=5)

        total_frame = ttk.Frame(right)
        total_frame.pack(fill="x", pady=10)
        self.pos_total_var = tk.StringVar(value="0.00")
        ttk.Label(total_frame, text="TOTAL: ", font=("Segoe UI", 14, "bold")).pack(side="left")
        ttk.Label(total_frame, textvariable=self.pos_total_var, font=("Segoe UI", 14, "bold")).pack(side="left")

        pay_frame = ttk.Frame(right)
        pay_frame.pack(fill="x", pady=5)
        ttk.Label(pay_frame, text="Payment Method:").grid(row=0, column=0, sticky="w")
        self.pos_payment_var = tk.StringVar(value="Cash")
        payment_combo = ttk.Combobox(pay_frame, textvariable=self.pos_payment_var,
                                      values=["Cash", "Online"], state="readonly", width=15)
        payment_combo.grid(row=0, column=1, padx=10)
        payment_combo.bind("<<ComboboxSelected>>", self.pos_toggle_reference)

        ttk.Label(pay_frame, text="Reference # (Online only):").grid(row=1, column=0, sticky="w", pady=(8, 0))
        self.pos_reference_var = tk.StringVar()
        self.pos_reference_entry = ttk.Entry(pay_frame, textvariable=self.pos_reference_var, width=25, state="disabled")
        self.pos_reference_entry.grid(row=1, column=1, pady=(8, 0), padx=10)

        ttk.Button(right, text="Complete Sale", command=self.pos_complete_sale).pack(pady=15, ipadx=10, ipady=5)

    def pos_toggle_reference(self, event=None):
        if self.pos_payment_var.get() == "Online":
            self.pos_reference_entry.configure(state="normal")
        else:
            self.pos_reference_entry.configure(state="disabled")
            self.pos_reference_var.set("")

    def pos_lookup_sku(self, event=None):
        sku = clean_str(self.pos_sku_var.get())
        inv = self.db.df("Inventory")
        match = inv[inv["SKU"].str.lower() == sku.lower()]
        if match.empty:
            messagebox.showwarning("Not Found", f"No inventory item with SKU '{sku}'.")
            self.pos_item_var.set("-")
            self.pos_price_var.set("-")
            self.pos_stock_var.set("-")
            return
        row = match.iloc[0]
        self.pos_item_var.set(row["ItemName"])
        self.pos_price_var.set(fmt_money(row["UnitPrice"]))
        self.pos_stock_var.set(str(int(row["QuantityOnHand"])))
        self.pos_qty_var.set(1)

    def pos_add_to_cart(self):
        sku = clean_str(self.pos_sku_var.get())
        inv = self.db.df("Inventory")
        match = inv[inv["SKU"].str.lower() == sku.lower()]
        if match.empty:
            messagebox.showwarning("No Item", "Scan or enter a valid SKU first.")
            return
        row = match.iloc[0]
        qty = self.pos_qty_var.get()
        if qty <= 0:
            messagebox.showwarning("Invalid Quantity", "Quantity must be at least 1.")
            return

        already_in_cart = sum(item["Quantity"] for item in self.cart if item["SKU"] == row["SKU"])
        if qty + already_in_cart > row["QuantityOnHand"]:
            messagebox.showwarning(
                "Insufficient Stock",
                f"Only {int(row['QuantityOnHand'])} unit(s) of '{row['ItemName']}' in stock "
                f"({already_in_cart} already in cart)."
            )
            return

        self.cart.append({
            "SKU": row["SKU"], "ItemName": row["ItemName"],
            "Quantity": qty, "UnitPrice": float(row["UnitPrice"]),
        })
        self.pos_refresh_cart()
        self.pos_sku_var.set("")
        self.pos_item_var.set("-")
        self.pos_price_var.set("-")
        self.pos_stock_var.set("-")
        self.pos_sku_entry.focus_set()

    def pos_remove_line(self):
        sel = self.pos_cart_tree.selection()
        if not sel:
            return
        idx = self.pos_cart_tree.index(sel[0])
        del self.cart[idx]
        self.pos_refresh_cart()

    def pos_refresh_cart(self):
        self.pos_cart_tree.delete(*self.pos_cart_tree.get_children())
        total = 0.0
        for item in self.cart:
            subtotal = item["Quantity"] * item["UnitPrice"]
            total += subtotal
            self.pos_cart_tree.insert("", "end", values=(
                item["SKU"], item["ItemName"], item["Quantity"],
                fmt_money(item["UnitPrice"]), fmt_money(subtotal)
            ))
        self.pos_total_var.set(fmt_money(total))

    def pos_complete_sale(self):
        if not self.cart:
            messagebox.showwarning("Empty Cart", "Add at least one item before completing the sale.")
            return
        payment_method = self.pos_payment_var.get()
        reference = clean_str(self.pos_reference_var.get())
        if payment_method == "Online" and not reference:
            messagebox.showwarning(
                "Reference Required",
                "Online payments require a Transaction Reference Number for bank verification."
            )
            return

        inv = self.db.df("Inventory")
        txn = self.db.df("Transactions")
        timestamp = now_str()
        new_rows = []
        for item in self.cart:
            idx = inv.index[inv["SKU"] == item["SKU"]]
            if len(idx) == 0:
                continue
            inv.loc[idx, "QuantityOnHand"] = inv.loc[idx, "QuantityOnHand"] - item["Quantity"]
            inv.loc[idx, "LastUpdated"] = timestamp

            new_rows.append({
                "TransactionID": new_id("TXN"),
                "DateTime": timestamp,
                "SKU": item["SKU"],
                "ItemName": item["ItemName"],
                "Quantity": item["Quantity"],
                "UnitPrice": item["UnitPrice"],
                "TotalAmount": round(item["Quantity"] * item["UnitPrice"], 2),
                "PaymentMethod": payment_method,
                "ReferenceNumber": reference if payment_method == "Online" else "",
                "EmployeeName": self.current_employee["Name"],
                "Status": "Completed",
                "VoidReason": "",
                "VoidedBy": "",
                "VoidedDateTime": "",
            })

        self.db.set_df("Inventory", inv)
        txn = pd.concat([txn, pd.DataFrame(new_rows)], ignore_index=True)
        self.db.set_df("Transactions", txn)

        total = sum(r["TotalAmount"] for r in new_rows)
        self.log_action(
            "Sale Completed",
            f"{len(new_rows)} line(s), total {fmt_money(total)}, payment={payment_method}, ref={reference}"
        )
        self.save_db()

        messagebox.showinfo("Sale Complete", f"Sale recorded. Total: {fmt_money(total)}")
        self.cart = []
        self.pos_refresh_cart()
        self.pos_reference_var.set("")
        self.pos_payment_var.set("Cash")
        self.pos_toggle_reference()
        self.refresh_inventory_tree()

    # ------------------------------------------------------------------
    # TAB 2: STOCK INTAKE  (Requirements 2, 3, 6)
    # ------------------------------------------------------------------
    def build_intake_tab(self, parent):
        frame = ttk.Frame(parent, padding=15)
        frame.pack(fill="both", expand=True)

        ttk.Label(frame, text="Scan / Enter SKU:", font=("Segoe UI", 10, "bold")).grid(row=0, column=0, sticky="w")
        self.intake_sku_var = tk.StringVar()
        sku_entry = ttk.Entry(frame, textvariable=self.intake_sku_var, width=30)
        sku_entry.grid(row=0, column=1, sticky="w", pady=5)
        sku_entry.bind("<Return>", self.intake_lookup_sku)

        fields = [
            ("Item Name:", "intake_name_var"),
            ("Category (auto-forced for Extra items):", "intake_category_var"),
            ("Unit Cost:", "intake_cost_var"),
            ("Unit Price:", "intake_price_var"),
            ("Reorder Level:", "intake_reorder_var"),
        ]
        for i, (label, attr) in enumerate(fields, start=1):
            ttk.Label(frame, text=label).grid(row=i, column=0, sticky="w", pady=3)
            var = tk.StringVar()
            setattr(self, attr, var)
            ttk.Entry(frame, textvariable=var, width=30).grid(row=i, column=1, sticky="w", pady=3)

        ttk.Label(frame, text="Quantity to Add:").grid(row=len(fields) + 1, column=0, sticky="w", pady=3)
        self.intake_qty_var = tk.IntVar(value=1)
        ttk.Spinbox(frame, from_=1, to=99999, textvariable=self.intake_qty_var, width=10).grid(
            row=len(fields) + 1, column=1, sticky="w", pady=3
        )

        ttk.Label(frame, text="Notes:").grid(row=len(fields) + 2, column=0, sticky="w", pady=3)
        self.intake_notes_var = tk.StringVar()
        ttk.Entry(frame, textvariable=self.intake_notes_var, width=40).grid(
            row=len(fields) + 2, column=1, sticky="w", pady=3
        )

        ttk.Button(frame, text="Add / Update Stock", command=self.intake_submit).grid(
            row=len(fields) + 3, column=0, columnspan=2, pady=15
        )

        ttk.Label(
            frame,
            text="Tip: if the SKU already exists, scanning it will pre-fill the current details.\n"
                 "Items whose name contains a configured keyword (e.g. 'rice', 'paper cups') are\n"
                 "automatically placed in the 'Extra' category no matter what you type above.",
            foreground="gray", justify="left"
        ).grid(row=len(fields) + 4, column=0, columnspan=2, sticky="w", pady=10)

    def intake_lookup_sku(self, event=None):
        sku = clean_str(self.intake_sku_var.get())
        inv = self.db.df("Inventory")
        match = inv[inv["SKU"].str.lower() == sku.lower()]
        if match.empty:
            self.intake_name_var.set("")
            self.intake_category_var.set("")
            self.intake_cost_var.set("")
            self.intake_price_var.set("")
            self.intake_reorder_var.set("")
            return
        row = match.iloc[0]
        self.intake_name_var.set(row["ItemName"])
        self.intake_category_var.set(row["Category"])
        self.intake_cost_var.set(str(row["UnitCost"]))
        self.intake_price_var.set(str(row["UnitPrice"]))
        self.intake_reorder_var.set(str(int(row["ReorderLevel"])))

    def intake_submit(self):
        sku = clean_str(self.intake_sku_var.get())
        name = self.intake_name_var.get().strip()
        if not sku or not name:
            messagebox.showwarning("Missing Data", "SKU and Item Name are required.")
            return
        try:
            cost = float(self.intake_cost_var.get() or 0)
            price = float(self.intake_price_var.get() or 0)
            reorder = int(float(self.intake_reorder_var.get() or 0))
            qty_add = int(self.intake_qty_var.get())
        except ValueError:
            messagebox.showwarning("Invalid Number", "Cost, price, reorder level and quantity must be numeric.")
            return

        category = resolve_category(name, self.intake_category_var.get())
        timestamp = now_str()
        inv = self.db.df("Inventory")
        idx = inv.index[inv["SKU"].str.lower() == sku.lower()]

        if len(idx) > 0:
            inv.loc[idx, "ItemName"] = name
            inv.loc[idx, "Category"] = category
            inv.loc[idx, "UnitCost"] = cost
            inv.loc[idx, "UnitPrice"] = price
            inv.loc[idx, "ReorderLevel"] = reorder
            inv.loc[idx, "QuantityOnHand"] = inv.loc[idx, "QuantityOnHand"] + qty_add
            inv.loc[idx, "LastUpdated"] = timestamp
        else:
            new_row = {
                "SKU": sku, "ItemName": name, "Category": category,
                "UnitCost": cost, "UnitPrice": price, "QuantityOnHand": qty_add,
                "ReorderLevel": reorder, "LastUpdated": timestamp,
            }
            inv = pd.concat([inv, pd.DataFrame([new_row])], ignore_index=True)

        self.db.set_df("Inventory", inv)

        intake_df = self.db.df("StockIntake")
        intake_row = {
            "IntakeID": new_id("INT"), "DateTime": timestamp, "SKU": sku, "ItemName": name,
            "QuantityAdded": qty_add, "UnitCost": cost,
            "EmployeeName": self.current_employee["Name"], "Notes": self.intake_notes_var.get().strip(),
        }
        intake_df = pd.concat([intake_df, pd.DataFrame([intake_row])], ignore_index=True)
        self.db.set_df("StockIntake", intake_df)

        self.log_action("Stock Intake", f"SKU={sku}, +{qty_add} units, category={category}")
        self.save_db()

        messagebox.showinfo("Stock Updated", f"'{name}' stock updated (+{qty_add}).")
        for var in [self.intake_sku_var, self.intake_name_var, self.intake_category_var,
                    self.intake_cost_var, self.intake_price_var, self.intake_reorder_var,
                    self.intake_notes_var]:
            var.set("")
        self.intake_qty_var.set(1)
        self.refresh_inventory_tree()

    # ------------------------------------------------------------------
    # TAB 3: INVENTORY VIEW
    # ------------------------------------------------------------------
    def build_inventory_tab(self, parent):
        frame = ttk.Frame(parent, padding=10)
        frame.pack(fill="both", expand=True)

        ttk.Button(frame, text="Refresh", command=self.refresh_inventory_tree).pack(anchor="w", pady=5)

        columns = ("SKU", "ItemName", "Category", "UnitCost", "UnitPrice",
                   "QuantityOnHand", "ReorderLevel", "LastUpdated")
        self.inventory_tree = ttk.Treeview(frame, columns=columns, show="headings")
        for c in columns:
            self.inventory_tree.heading(c, text=c)
            self.inventory_tree.column(c, width=110)
        self.inventory_tree.tag_configure("low_stock", background="#ffcccc")
        self.inventory_tree.pack(fill="both", expand=True)

        ttk.Label(frame, text="Rows highlighted red are at or below their reorder level.",
                  foreground="gray").pack(anchor="w", pady=5)

        self.refresh_inventory_tree()

    def refresh_inventory_tree(self):
        if not hasattr(self, "inventory_tree"):
            return
        self.inventory_tree.delete(*self.inventory_tree.get_children())
        inv = self.db.df("Inventory")
        for _, row in inv.iterrows():
            low = row["QuantityOnHand"] <= row["ReorderLevel"]
            self.inventory_tree.insert("", "end", values=(
                row["SKU"], row["ItemName"], row["Category"], fmt_money(row["UnitCost"]),
                fmt_money(row["UnitPrice"]), int(row["QuantityOnHand"]), int(row["ReorderLevel"]),
                row["LastUpdated"]
            ), tags=("low_stock",) if low else ())

    # ------------------------------------------------------------------
    # TAB 4: SALES HISTORY / VOID  (Requirements 3, 4)
    # ------------------------------------------------------------------
    def build_history_tab(self, parent):
        frame = ttk.Frame(parent, padding=10)
        frame.pack(fill="both", expand=True)

        controls = ttk.Frame(frame)
        controls.pack(fill="x", pady=5)
        ttk.Label(controls, text="Filter:").pack(side="left")
        self.history_filter_var = tk.StringVar(value="Today")
        ttk.Combobox(controls, textvariable=self.history_filter_var,
                     values=["Today", "All", "Voided Only"], state="readonly", width=15).pack(side="left", padx=5)
        ttk.Button(controls, text="Refresh", command=self.refresh_history_tree).pack(side="left", padx=5)
        ttk.Button(controls, text="Void Selected Sale", command=self.void_selected_sale).pack(side="left", padx=15)

        columns = ("TransactionID", "DateTime", "SKU", "ItemName", "Quantity", "TotalAmount",
                   "PaymentMethod", "ReferenceNumber", "EmployeeName", "Status", "VoidReason", "VoidedBy")
        self.history_tree = ttk.Treeview(frame, columns=columns, show="headings")
        for c in columns:
            self.history_tree.heading(c, text=c)
            self.history_tree.column(c, width=105)
        self.history_tree.tag_configure("voided", background="#ffe0b3")
        self.history_tree.pack(fill="both", expand=True)

        self.refresh_history_tree()

    def refresh_history_tree(self):
        if not hasattr(self, "history_tree"):
            return
        self.history_tree.delete(*self.history_tree.get_children())
        txn = self.db.df("Transactions").copy()
        mode = self.history_filter_var.get()
        if mode == "Today":
            txn = txn[txn["DateTime"].astype(str).str.startswith(today_str())]
        elif mode == "Voided Only":
            txn = txn[txn["Status"] == "Voided"]
        txn = txn.sort_values("DateTime", ascending=False)
        for _, row in txn.iterrows():
            self.history_tree.insert("", "end", values=(
                row["TransactionID"], row["DateTime"], row["SKU"], row["ItemName"], int(row["Quantity"]),
                fmt_money(row["TotalAmount"]), row["PaymentMethod"], row["ReferenceNumber"],
                row["EmployeeName"], row["Status"], row["VoidReason"], row["VoidedBy"]
            ), tags=("voided",) if row["Status"] == "Voided" else ())

    def void_selected_sale(self):
        sel = self.history_tree.selection()
        if not sel:
            messagebox.showwarning("No Selection", "Select a transaction to void.")
            return
        values = self.history_tree.item(sel[0], "values")
        txn_id = values[0]

        txn = self.db.df("Transactions")
        idx = txn.index[txn["TransactionID"] == txn_id]
        if len(idx) == 0:
            return
        row = txn.loc[idx[0]]
        if row["Status"] == "Voided":
            messagebox.showinfo("Already Voided", "This transaction has already been voided.")
            return

        # Requirement 4: employees can never delete a record, only void it,
        # and voiding must be explicitly flagged for the manager's report.
        reason = simpledialog.askstring("Void Reason", "Enter the reason for voiding this sale:")
        if not reason:
            messagebox.showwarning("Void Cancelled", "A reason is required to void a sale.")
            return

        manager_pin = simpledialog.askstring("Manager Authorization", "Enter a Manager PIN to authorize this void:", show="*")
        if manager_pin is None:
            return
        approver = self.find_employee_by_pin(manager_pin)
        if not approver or approver["Role"] != "Manager":
            messagebox.showerror("Not Authorized", "A valid Manager PIN is required to void a sale.")
            return

        txn.loc[idx, "Status"] = "Voided"
        txn.loc[idx, "VoidReason"] = reason
        txn.loc[idx, "VoidedBy"] = approver["Name"]
        txn.loc[idx, "VoidedDateTime"] = now_str()
        self.db.set_df("Transactions", txn)

        # restock the voided item
        inv = self.db.df("Inventory")
        inv_idx = inv.index[inv["SKU"] == row["SKU"]]
        if len(inv_idx) > 0:
            inv.loc[inv_idx, "QuantityOnHand"] = inv.loc[inv_idx, "QuantityOnHand"] + row["Quantity"]
            self.db.set_df("Inventory", inv)

        self.log_action(
            "Sale Voided",
            f"TransactionID={txn_id}, originally by {row['EmployeeName']}, "
            f"approved by {approver['Name']}, reason='{reason}'"
        )
        self.save_db()
        messagebox.showinfo("Voided", "Transaction voided and flagged for the manager's report.")
        self.refresh_history_tree()
        self.refresh_inventory_tree()

    # ------------------------------------------------------------------
    # TAB 5: END-OF-DAY RECONCILIATION  (Requirement 7)
    # ------------------------------------------------------------------
    def build_reconciliation_tab(self, parent):
        frame = ttk.Frame(parent, padding=10)
        frame.pack(fill="both", expand=True)

        # --- revenue reconciliation -----------------------------------
        rev_box = ttk.Labelframe(frame, text="Revenue Reconciliation", padding=10)
        rev_box.pack(fill="x", pady=8)

        ttk.Button(rev_box, text="Compute Today's Expected Totals",
                   command=self.recon_compute_expected).grid(row=0, column=0, columnspan=2, pady=5)

        self.recon_expected_cash_var = tk.StringVar(value="0.00")
        self.recon_expected_online_var = tk.StringVar(value="0.00")
        ttk.Label(rev_box, text="Expected Cash:").grid(row=1, column=0, sticky="w")
        ttk.Label(rev_box, textvariable=self.recon_expected_cash_var, font=("Segoe UI", 10, "bold")).grid(row=1, column=1, sticky="w")
        ttk.Label(rev_box, text="Expected Online:").grid(row=2, column=0, sticky="w")
        ttk.Label(rev_box, textvariable=self.recon_expected_online_var, font=("Segoe UI", 10, "bold")).grid(row=2, column=1, sticky="w")

        ttk.Label(rev_box, text="Actual Cash in Drawer:").grid(row=3, column=0, sticky="w", pady=(10, 0))
        self.recon_actual_cash_var = tk.StringVar()
        ttk.Entry(rev_box, textvariable=self.recon_actual_cash_var, width=15).grid(row=3, column=1, sticky="w", pady=(10, 0))

        ttk.Label(rev_box, text="Actual Online (verified w/ bank):").grid(row=4, column=0, sticky="w")
        self.recon_actual_online_var = tk.StringVar()
        ttk.Entry(rev_box, textvariable=self.recon_actual_online_var, width=15).grid(row=4, column=1, sticky="w")

        ttk.Button(rev_box, text="Calculate Variance", command=self.recon_calc_variance).grid(row=5, column=0, columnspan=2, pady=8)

        self.recon_variance_label = ttk.Label(rev_box, text="", font=("Segoe UI", 10, "bold"))
        self.recon_variance_label.grid(row=6, column=0, columnspan=2, sticky="w")

        # --- stock reconciliation ---------------------------------------
        stock_box = ttk.Labelframe(frame, text="Stock Reconciliation (double-click 'ActualQty' to enter a count)", padding=10)
        stock_box.pack(fill="both", expand=True, pady=8)

        btn_row = ttk.Frame(stock_box)
        btn_row.pack(fill="x")
        ttk.Button(btn_row, text="Load Inventory for Count", command=self.recon_load_stock).pack(side="left")
        ttk.Button(btn_row, text="Apply Adjustments to Inventory (Manager)", command=self.recon_apply_adjustments).pack(side="left", padx=10)

        columns = ("SKU", "ItemName", "ExpectedQty", "ActualQty", "Variance")
        self.recon_tree = ttk.Treeview(stock_box, columns=columns, show="headings", height=10)
        for c in columns:
            self.recon_tree.heading(c, text=c)
            self.recon_tree.column(c, width=130)
        self.recon_tree.tag_configure("variance", background="#ffcccc")
        self.recon_tree.pack(fill="both", expand=True, pady=5)
        self.recon_tree.bind("<Double-1>", self.recon_edit_actual_qty)

        notes_row = ttk.Frame(frame)
        notes_row.pack(fill="x", pady=5)
        ttk.Label(notes_row, text="Notes:").pack(side="left")
        self.recon_notes_var = tk.StringVar()
        ttk.Entry(notes_row, textvariable=self.recon_notes_var, width=60).pack(side="left", padx=5)

        ttk.Button(frame, text="Save End-of-Day Reconciliation Report",
                   command=self.recon_save_report).pack(pady=10)

    def recon_compute_expected(self):
        txn = self.db.df("Transactions")
        today_txn = txn[(txn["DateTime"].astype(str).str.startswith(today_str())) & (txn["Status"] == "Completed")]
        cash_total = today_txn[today_txn["PaymentMethod"] == "Cash"]["TotalAmount"].sum()
        online_total = today_txn[today_txn["PaymentMethod"] == "Online"]["TotalAmount"].sum()
        self.recon_expected_cash_var.set(fmt_money(cash_total))
        self.recon_expected_online_var.set(fmt_money(online_total))

    def recon_calc_variance(self):
        try:
            expected_cash = float(self.recon_expected_cash_var.get())
            expected_online = float(self.recon_expected_online_var.get())
            actual_cash = float(self.recon_actual_cash_var.get() or 0)
            actual_online = float(self.recon_actual_online_var.get() or 0)
        except ValueError:
            messagebox.showwarning("Invalid Input", "Enter numeric values for actual cash / online totals.")
            return

        cash_var = round(actual_cash - expected_cash, 2)
        online_var = round(actual_online - expected_online, 2)

        msg = f"Cash Variance: {cash_var:+.2f}   |   Online Variance: {online_var:+.2f}"
        self.recon_variance_label.configure(text=msg)

        if cash_var != 0 or online_var != 0:
            self.recon_variance_label.configure(foreground="red")
            messagebox.showwarning(
                "Discrepancy Detected",
                f"Revenue variance detected!\nCash: {cash_var:+.2f}\nOnline: {online_var:+.2f}"
            )
        else:
            self.recon_variance_label.configure(foreground="green")

    def recon_load_stock(self):
        self.recon_tree.delete(*self.recon_tree.get_children())
        inv = self.db.df("Inventory")
        for _, row in inv.iterrows():
            self.recon_tree.insert("", "end", values=(
                row["SKU"], row["ItemName"], int(row["QuantityOnHand"]), "", ""
            ))

    def recon_edit_actual_qty(self, event):
        item_id = self.recon_tree.identify_row(event.y)
        col = self.recon_tree.identify_column(event.x)
        if not item_id or col != "#4":  # ActualQty column
            return
        values = list(self.recon_tree.item(item_id, "values"))
        expected = int(values[2])
        actual = simpledialog.askinteger("Physical Count", f"Enter the counted quantity for {values[1]} ({values[0]}):")
        if actual is None:
            return
        variance = actual - expected
        values[3] = actual
        values[4] = variance
        self.recon_tree.item(item_id, values=values, tags=("variance",) if variance != 0 else ())

    def recon_apply_adjustments(self):
        if self.current_employee["Role"] != "Manager":
            messagebox.showerror("Not Authorized", "Only a Manager may apply stock adjustments.")
            return
        inv = self.db.df("Inventory")
        applied = 0
        flagged = []
        for item_id in self.recon_tree.get_children():
            sku, item_name, expected, actual, variance = self.recon_tree.item(item_id, "values")
            if actual == "":
                continue
            variance = int(variance)
            if variance == 0:
                continue
            idx = inv.index[inv["SKU"] == sku]
            if len(idx) == 0:
                continue
            inv.loc[idx, "QuantityOnHand"] = int(actual)
            inv.loc[idx, "LastUpdated"] = now_str()
            applied += 1
            flagged.append(f"{sku} ({item_name}): expected {expected}, counted {actual}, variance {variance:+d}")

        if applied == 0:
            messagebox.showinfo("No Adjustments", "No counted items with a variance were found.")
            return

        self.db.set_df("Inventory", inv)
        details = "; ".join(flagged)
        self.log_action("Stock Adjustment Applied", details)
        self.save_db()
        messagebox.showinfo("Adjustments Applied", f"{applied} SKU(s) adjusted to match the physical count.")
        self.refresh_inventory_tree()

    def recon_save_report(self):
        if self.current_employee["Role"] != "Manager":
            messagebox.showerror("Not Authorized", "Only a Manager may save the end-of-day reconciliation report.")
            return
        try:
            expected_cash = float(self.recon_expected_cash_var.get())
            expected_online = float(self.recon_expected_online_var.get())
            actual_cash = float(self.recon_actual_cash_var.get() or 0)
            actual_online = float(self.recon_actual_online_var.get() or 0)
        except ValueError:
            messagebox.showwarning("Invalid Input", "Enter numeric values for actual cash / online totals first.")
            return

        cash_var = round(actual_cash - expected_cash, 2)
        online_var = round(actual_online - expected_online, 2)

        stock_variances = []
        for item_id in self.recon_tree.get_children():
            sku, item_name, expected, actual, variance = self.recon_tree.item(item_id, "values")
            if actual != "" and int(variance) != 0:
                stock_variances.append(f"{sku}:{variance}")

        recon_df = self.db.df("Reconciliation")
        row = {
            "ReconID": new_id("REC"), "Date": today_str(),
            "ExpectedCash": expected_cash, "ActualCash": actual_cash, "CashVariance": cash_var,
            "ExpectedOnline": expected_online, "ActualOnline": actual_online, "OnlineVariance": online_var,
            "EmployeeName": self.current_employee["Name"],
            "Notes": self.recon_notes_var.get().strip() + (
                f" | Stock variances: {', '.join(stock_variances)}" if stock_variances else ""
            ),
        }
        recon_df = pd.concat([recon_df, pd.DataFrame([row])], ignore_index=True)
        self.db.set_df("Reconciliation", recon_df)

        self.log_action(
            "End-of-Day Reconciliation Saved",
            f"CashVar={cash_var:+.2f}, OnlineVar={online_var:+.2f}, StockVariances={len(stock_variances)}"
        )
        self.save_db()

        alert = ""
        if cash_var != 0 or online_var != 0 or stock_variances:
            alert = "\n\nALERT: One or more discrepancies were recorded on this report."
        messagebox.showinfo("Report Saved", "End-of-day reconciliation report saved." + alert)

    # ------------------------------------------------------------------
    # TAB 6: AUDIT LOG  (Manager only, Requirement 3)
    # ------------------------------------------------------------------
    def build_audit_tab(self, parent):
        frame = ttk.Frame(parent, padding=10)
        frame.pack(fill="both", expand=True)
        ttk.Button(frame, text="Refresh", command=self.refresh_audit_tree).pack(anchor="w", pady=5)

        columns = ("LogID", "DateTime", "EmployeeName", "Action", "Details")
        self.audit_tree = ttk.Treeview(frame, columns=columns, show="headings")
        for c in columns:
            self.audit_tree.heading(c, text=c)
            self.audit_tree.column(c, width=150 if c != "Details" else 350)
        self.audit_tree.pack(fill="both", expand=True)

        self.refresh_audit_tree()

    def refresh_audit_tree(self):
        if not hasattr(self, "audit_tree"):
            return
        self.audit_tree.delete(*self.audit_tree.get_children())
        df = self.db.df("AuditLog").sort_values("DateTime", ascending=False)
        for _, row in df.iterrows():
            self.audit_tree.insert("", "end", values=(
                row["LogID"], row["DateTime"], row["EmployeeName"], row["Action"], row["Details"]
            ))


if __name__ == "__main__":
    app = App()
    app.mainloop()
