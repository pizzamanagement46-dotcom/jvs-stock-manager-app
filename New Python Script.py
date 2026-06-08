"""
Stock Manager Pro — Streamlit Web Edition
==========================================
Converted from Tkinter desktop app to Streamlit web app.
Run with: streamlit run stock_app_streamlit.py
"""

import streamlit as st
import json, os, copy
from datetime import datetime, date
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter
import io

# ── Page Config ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Stock Manager Pro",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ── Constants ────────────────────────────────────────────────────────────────
DATA_FILE = "stock_data.json"
LOW_STOCK  = 10
SHIFTS     = ["Morning", "Evening", "Night"]

# ── Custom CSS (dark theme matching original) ────────────────────────────────
st.markdown("""
<style>
    /* Dark background */
    .stApp { background-color: #1e272e; }
    section[data-testid="stSidebar"] { background-color: #2f3640; }
    section[data-testid="stSidebar"] * { color: #ffffff !important; }

    /* Metric cards */
    [data-testid="metric-container"] {
        background-color: #2f3640;
        border: 1px solid #353b48;
        border-radius: 10px;
        padding: 12px;
    }
    [data-testid="metric-container"] label { color: #aaaaaa !important; }
    [data-testid="metric-container"] [data-testid="stMetricValue"] {
        color: #44bd32 !important; font-size: 1.6rem !important;
    }

    /* Headers */
    h1, h2, h3, h4 { color: #00a8ff !important; }
    p, label, span { color: #ecf0f1; }

    /* Tables */
    [data-testid="stDataFrame"] { border-radius: 8px; overflow: hidden; }

    /* Buttons */
    .stButton > button {
        background-color: #273c75;
        color: white;
        border: none;
        border-radius: 6px;
        font-weight: bold;
    }
    .stButton > button:hover { background-color: #40739e; }

    /* Input fields */
    .stTextInput > div > div > input,
    .stNumberInput > div > div > input,
    .stSelectbox > div > div,
    .stDateInput > div > div > input {
        background-color: #353b48 !important;
        color: #ffffff !important;
        border: 1px solid #40739e !important;
    }
    
    /* Tabs */
    .stTabs [data-baseweb="tab-list"] { background-color: #2f3640; border-radius: 8px; }
    .stTabs [data-baseweb="tab"] { color: #aaaaaa; }
    .stTabs [aria-selected="true"] { color: #00a8ff !important; font-weight: bold; }
    
    /* Success / warning / error */
    .stSuccess { background-color: #1e8449 !important; }
    .stWarning { background-color: #7d6608 !important; }
    
    /* Section divider */
    hr { border-color: #353b48; }
</style>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# DATA LAYER  (same logic as original — just reads/writes stock_data.json)
# ══════════════════════════════════════════════════════════════════════════════

def load_data():
    if os.path.exists(DATA_FILE):
        with open(DATA_FILE, "r") as f:
            d = json.load(f)
    else:
        d = {}
    d.setdefault("items", {})
    d.setdefault("history", [])
    d.setdefault("storage_stores", [])
    d.setdefault("storage_products", {})
    d.setdefault("storage_qty", {})
    d.setdefault("storage_history", [])
    d.setdefault("stores_list", [])
    d.setdefault("store_products", {})
    d.setdefault("daily_shift_data", {})
    d.setdefault("drop_log", [])
    d.setdefault("ledger_products", {})
    d.setdefault("stock_transactions", [])
    d.setdefault("ledger_entries", [])
    d.setdefault("ledger_payments", [])
    d.setdefault("audit_log", [])
    d.setdefault("reconcile_log", [])
    return d


def save_data(d):
    tmp = DATA_FILE + ".tmp"
    try:
        with open(tmp, "w") as f:
            json.dump(d, f, indent=2)
        os.replace(tmp, DATA_FILE)
        return True
    except Exception as e:
        st.error(f"💾 Save Error: {e}")
        return False


# Cache data in session state so it persists across reruns
if "data" not in st.session_state:
    st.session_state.data = load_data()

data = st.session_state.data


def persist():
    """Save and refresh data reference."""
    save_data(st.session_state.data)
    st.session_state.data = load_data()


# ── Ledger helpers ────────────────────────────────────────────────────────────

def ledger_products():
    return data.get("ledger_products", {})

def stock_transactions():
    return data.get("stock_transactions", [])

def ledger_entries():
    return data.get("ledger_entries", [])

def ledger_payments():
    return data.get("ledger_payments", [])

def ledger_stock_balance(product):
    bal = 0
    for tx in stock_transactions():
        if tx["product"] == product:
            if tx["type"] == "receive":
                bal += int(tx.get("qty", 0) or 0)
            else:
                bal -= int(tx.get("qty", 0) or 0)
    return bal

def ledger_price_balance(product):
    bal = 0.0
    for e in ledger_entries():
        if e["product"] == product:
            bal += float(e.get("debit", 0) or 0)
            bal -= float(e.get("credit", 0) or 0)
    return bal


# ══════════════════════════════════════════════════════════════════════════════
# SIDEBAR NAVIGATION
# ══════════════════════════════════════════════════════════════════════════════

with st.sidebar:
    st.markdown("## 📦 Stock Manager Pro")
    st.markdown("---")
    page = st.radio(
        "Navigation",
        ["🏠 Dashboard", "🏪 Storage Count", "📤 Drop Stock",
         "🕐 Shift Sale", "📒 Ledger"],
        label_visibility="collapsed"
    )
    st.markdown("---")
    st.caption(f"Last saved: {datetime.now().strftime('%H:%M:%S')}")
    if st.button("🔄 Refresh Data"):
        st.session_state.data = load_data()
        st.rerun()


# ══════════════════════════════════════════════════════════════════════════════
# PAGE 1 — DASHBOARD
# ══════════════════════════════════════════════════════════════════════════════

if page == "🏠 Dashboard":
    st.title("📊 Dashboard")
    st.caption("Live summary of your stock across all stores")

    # ── Summary cards ─────────────────────────────────────────────────────────
    total_products  = len(ledger_products())
    total_stores    = len(data.get("storage_stores", []))
    drop_count      = len(data.get("drop_log", []))
    low_stock_items = [p for p in ledger_products()
                       if ledger_stock_balance(p) < LOW_STOCK]

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("📦 Total Products", total_products)
    c2.metric("🏪 Stores", total_stores)
    c3.metric("📤 Total Drops", drop_count)
    c4.metric("⚠️ Low Stock Items", len(low_stock_items),
              delta=f"-{len(low_stock_items)} below {LOW_STOCK}",
              delta_color="inverse")

    st.divider()

    # ── Product stock balances ─────────────────────────────────────────────────
    st.subheader("📦 Product Stock Balances")
    if ledger_products():
        import pandas as pd
        rows = []
        for prod, meta in sorted(ledger_products().items()):
            bal   = ledger_stock_balance(prod)
            price = float(meta.get("price", 0))
            rows.append({
                "Product":       prod,
                "Unit":          meta.get("unit", "pcs"),
                "Stock Balance": bal,
                "Price (₹)":     price,
                "Total Value (₹)": round(bal * price, 2),
                "Status":        "⚠️ Low" if bal < LOW_STOCK else "✅ OK"
            })
        df = pd.DataFrame(rows)
        st.dataframe(df, use_container_width=True, hide_index=True)
    else:
        st.info("No products added yet. Go to Ledger → Manage Products to add products.")

    st.divider()

    # ── Recent drop log ────────────────────────────────────────────────────────
    st.subheader("📤 Recent Drop Log (Last 10)")
    drop_log = data.get("drop_log", [])
    if drop_log:
        import pandas as pd
        recent = sorted(drop_log, key=lambda x: x.get("ts",""), reverse=True)[:10]
        df2 = pd.DataFrame(recent)
        # Show only relevant columns if they exist
        cols = [c for c in ["ts","store","product","qty","note"] if c in df2.columns]
        st.dataframe(df2[cols], use_container_width=True, hide_index=True)
    else:
        st.info("No drops recorded yet.")


# ══════════════════════════════════════════════════════════════════════════════
# PAGE 2 — STORAGE COUNT
# ══════════════════════════════════════════════════════════════════════════════

elif page == "🏪 Storage Count":
    st.title("🏪 Storage Count")

    tab1, tab2 = st.tabs(["📋 View Storage", "⚙️ Manage Stores & Products"])

    # ── TAB 1: View / Edit daily storage ─────────────────────────────────────
    with tab1:
        storage_stores   = data.get("storage_stores", [])
        storage_products = data.get("storage_products", {})
        storage_qty      = data.get("storage_qty", {})

        if not storage_stores:
            st.warning("No stores added yet. Go to 'Manage Stores & Products' tab.")
        else:
            col1, col2 = st.columns([2, 2])
            with col1:
                sel_store = st.selectbox("Select Store", storage_stores)
            with col2:
                sel_date  = st.date_input("Date", value=date.today())

            date_key = sel_date.strftime("%Y-%m-%d")
            store_products = storage_products.get(sel_store, [])

            if not store_products:
                st.info(f"No products assigned to **{sel_store}**. Add them in the Manage tab.")
            else:
                st.markdown(f"### 📅 {sel_store} — {date_key}")

                # Build edit form
                with st.form(key=f"storage_form_{sel_store}_{date_key}"):
                    st.markdown("Enter values for each product:")
                    header = st.columns([3, 2, 2, 2, 2, 2, 3])
                    for h, col in zip(
                        ["Product", "MIH", "RCV", "DRP", "EXC", "NOTE", "Balance"],
                        header
                    ):
                        col.markdown(f"**{h}**")

                    new_vals = {}
                    for prod in store_products:
                        existing = (storage_qty
                                    .get(sel_store, {})
                                    .get(prod, {})
                                    .get("daily", {})
                                    .get(date_key, {}))
                        cols = st.columns([3, 2, 2, 2, 2, 2, 3])
                        cols[0].markdown(f"**{prod}**")
                        mih  = cols[1].text_input("", value=existing.get("mih",""),  key=f"mih_{prod}_{date_key}",  label_visibility="collapsed")
                        rcv  = cols[2].text_input("", value=existing.get("rcv",""),  key=f"rcv_{prod}_{date_key}",  label_visibility="collapsed")
                        drp  = cols[3].text_input("", value=existing.get("drp",""),  key=f"drp_{prod}_{date_key}",  label_visibility="collapsed")
                        exc  = cols[4].text_input("", value=existing.get("exc",""),  key=f"exc_{prod}_{date_key}",  label_visibility="collapsed")
                        note = cols[5].text_input("", value=existing.get("note",""), key=f"note_{prod}_{date_key}", label_visibility="collapsed")
                        # Calculate balance
                        try:
                            bal = int(mih or 0) + int(rcv or 0) - int(drp or 0) + int(exc or 0)
                            bal_str = str(bal)
                        except:
                            bal_str = "?"
                        cols[6].markdown(f"**{bal_str}**")
                        new_vals[prod] = {"mih": mih, "rcv": rcv, "drp": drp,
                                          "exc": exc, "note": note, "balance": bal_str}

                    if st.form_submit_button("💾 Save Storage Data", use_container_width=True):
                        sq = data.setdefault("storage_qty", {})
                        sq.setdefault(sel_store, {})
                        for prod, vals in new_vals.items():
                            sq[sel_store].setdefault(prod, {}).setdefault("daily", {})[date_key] = vals
                        persist()
                        st.success(f"✅ Saved storage data for {sel_store} on {date_key}")
                        st.rerun()

    # ── TAB 2: Manage Stores & Products ─────────────────────────────────────
    with tab2:
        st.subheader("🏪 Add / Remove Stores")
        c1, c2 = st.columns([3, 1])
        new_store = c1.text_input("New Store Name", placeholder="e.g. Main Warehouse")
        if c2.button("➕ Add Store", use_container_width=True):
            if new_store.strip():
                stores = data.setdefault("storage_stores", [])
                if new_store.strip() not in stores:
                    stores.append(new_store.strip())
                    persist()
                    st.success(f"Store '{new_store}' added!")
                    st.rerun()
                else:
                    st.warning("Store already exists.")

        stores = data.get("storage_stores", [])
        if stores:
            del_store = st.selectbox("Remove Store", ["-- select --"] + stores, key="del_store")
            if st.button("🗑 Delete Store", type="secondary"):
                if del_store != "-- select --":
                    data["storage_stores"].remove(del_store)
                    data.get("storage_products", {}).pop(del_store, None)
                    data.get("storage_qty", {}).pop(del_store, None)
                    persist()
                    st.success(f"Store '{del_store}' deleted.")
                    st.rerun()

        st.divider()
        st.subheader("📦 Assign Products to Store")
        if stores:
            c1, c2 = st.columns(2)
            with c1:
                assign_store = st.selectbox("Store", stores, key="assign_store")
            with c2:
                assign_prod = st.text_input("Product Name", placeholder="e.g. Milk 1L")
            if st.button("➕ Assign Product to Store"):
                if assign_prod.strip():
                    sp = data.setdefault("storage_products", {})
                    sp.setdefault(assign_store, [])
                    if assign_prod.strip() not in sp[assign_store]:
                        sp[assign_store].append(assign_prod.strip())
                        persist()
                        st.success(f"'{assign_prod}' added to {assign_store}")
                        st.rerun()
                    else:
                        st.warning("Product already in this store.")

            # Show current assignment
            st.markdown("**Current Products per Store:**")
            import pandas as pd
            sp = data.get("storage_products", {})
            if sp:
                for s, prods in sp.items():
                    if prods:
                        st.markdown(f"🏪 **{s}**: {', '.join(prods)}")
            else:
                st.info("No products assigned yet.")


# ══════════════════════════════════════════════════════════════════════════════
# PAGE 3 — DROP STOCK
# ══════════════════════════════════════════════════════════════════════════════

elif page == "📤 Drop Stock":
    st.title("📤 Drop Stock")

    tab1, tab2 = st.tabs(["➕ New Drop", "📋 Drop Log"])

    with tab1:
        stores = data.get("storage_stores", [])
        if not stores:
            st.warning("No stores found. Add stores in Storage Count → Manage Stores.")
        else:
            c1, c2 = st.columns(2)
            with c1:
                drop_store = st.selectbox("From Store", stores)
            with c2:
                drop_date = st.date_input("Date", value=date.today(), key="drop_date")

            store_prods = data.get("storage_products", {}).get(drop_store, [])
            if not store_prods:
                st.warning(f"No products assigned to **{drop_store}**.")
            else:
                drop_prod = st.selectbox("Product", store_prods)
                c1, c2 = st.columns(2)
                drop_qty  = c1.number_input("Quantity to Drop", min_value=1, value=1)
                drop_note = c2.text_input("Note (optional)")

                if st.button("📤 Record Drop", type="primary", use_container_width=True):
                    entry = {
                        "ts":      datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "date":    drop_date.strftime("%Y-%m-%d"),
                        "store":   drop_store,
                        "product": drop_prod,
                        "qty":     drop_qty,
                        "note":    drop_note,
                    }
                    data.setdefault("drop_log", []).append(entry)
                    # Also deduct from storage_qty on that day
                    date_key = drop_date.strftime("%Y-%m-%d")
                    sq = data.setdefault("storage_qty", {})
                    sq.setdefault(drop_store, {}).setdefault(drop_prod, {}).setdefault("daily", {}).setdefault(date_key, {})
                    existing_drp = int(sq[drop_store][drop_prod]["daily"][date_key].get("drp", 0) or 0)
                    sq[drop_store][drop_prod]["daily"][date_key]["drp"] = str(existing_drp + drop_qty)
                    persist()
                    st.success(f"✅ Dropped {drop_qty} × {drop_prod} from {drop_store}")
                    st.rerun()

    with tab2:
        st.subheader("📋 All Drop Records")
        drop_log = sorted(data.get("drop_log", []),
                          key=lambda x: x.get("ts",""), reverse=True)
        if drop_log:
            import pandas as pd
            df = pd.DataFrame(drop_log)
            cols = [c for c in ["date","store","product","qty","note","ts"] if c in df.columns]
            st.dataframe(df[cols], use_container_width=True, hide_index=True)

            # Export
            if st.button("📥 Export Drop Log to Excel"):
                wb = Workbook()
                ws = wb.active
                ws.title = "Drop Log"
                headers = ["Date","Store","Product","Qty","Note","Timestamp"]
                hdr_fill = PatternFill("solid", fgColor="273c75")
                hdr_font = Font(bold=True, color="FFFFFF")
                for ci, h in enumerate(headers, 1):
                    c = ws.cell(row=1, column=ci, value=h)
                    c.fill = hdr_fill; c.font = hdr_font
                for i, row in enumerate(drop_log, 2):
                    ws.append([row.get("date",""), row.get("store",""),
                                row.get("product",""), row.get("qty",0),
                                row.get("note",""), row.get("ts","")])
                buf = io.BytesIO()
                wb.save(buf)
                buf.seek(0)
                st.download_button("⬇️ Download Excel", buf,
                                   file_name=f"drop_log_{date.today()}.xlsx",
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        else:
            st.info("No drops recorded yet.")


# ══════════════════════════════════════════════════════════════════════════════
# PAGE 4 — SHIFT WISE SALE
# ══════════════════════════════════════════════════════════════════════════════

elif page == "🕐 Shift Sale":
    st.title("🕐 Shift Wise Sale")

    stores = data.get("stores_list", []) or data.get("storage_stores", [])
    if not stores:
        st.warning("No stores configured. Add stores in Storage Count first.")
    else:
        c1, c2 = st.columns([2, 2])
        sel_store = c1.selectbox("Store", stores)
        sel_date  = c2.date_input("Date", value=date.today(), key="shift_date")
        date_key  = sel_date.strftime("%Y-%m-%d")

        st.markdown(f"### 🏪 {sel_store} — {date_key}")
        tab_m, tab_e, tab_n = st.tabs(["🌅 Morning", "🌆 Evening", "🌙 Night"])

        for shift, tab in zip(SHIFTS, [tab_m, tab_e, tab_n]):
            with tab:
                shift_key = f"{sel_store}_{date_key}_{shift}"
                existing  = data.get("daily_shift_data", {}).get(shift_key, {})

                with st.form(key=f"shift_form_{shift_key}"):
                    c1, c2, c3 = st.columns(3)
                    opening = c1.number_input("Opening Balance", value=int(existing.get("opening", 0) or 0), min_value=0, key=f"op_{shift_key}")
                    sales   = c2.number_input("Sales",           value=int(existing.get("sales",   0) or 0), min_value=0, key=f"sl_{shift_key}")
                    closing = c3.number_input("Closing Balance", value=int(existing.get("closing", 0) or 0), min_value=0, key=f"cl_{shift_key}")
                    notes   = st.text_input("Notes", value=existing.get("notes",""), key=f"nt_{shift_key}")

                    # Auto-calculated
                    expected = opening - sales
                    diff     = closing - expected
                    diff_str = f"{'+'if diff>=0 else ''}{diff}"
                    st.markdown(f"**Expected Closing:** {expected}  |  **Difference:** {diff_str}")

                    if st.form_submit_button(f"💾 Save {shift} Shift", use_container_width=True):
                        data.setdefault("daily_shift_data", {})[shift_key] = {
                            "opening": opening,
                            "sales":   sales,
                            "closing": closing,
                            "notes":   notes,
                            "ts":      datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        }
                        persist()
                        st.success(f"✅ {shift} shift saved for {sel_store} on {date_key}")
                        st.rerun()

        # Daily summary
        st.divider()
        st.subheader("📊 Daily Summary")
        total_sales = 0
        summary_rows = []
        for shift in SHIFTS:
            shift_key = f"{sel_store}_{date_key}_{shift}"
            sd = data.get("daily_shift_data", {}).get(shift_key, {})
            sales = int(sd.get("sales", 0) or 0)
            total_sales += sales
            summary_rows.append({
                "Shift":   shift,
                "Opening": sd.get("opening", "-"),
                "Sales":   sales,
                "Closing": sd.get("closing", "-"),
                "Notes":   sd.get("notes", ""),
            })
        import pandas as pd
        st.dataframe(pd.DataFrame(summary_rows), use_container_width=True, hide_index=True)
        st.metric("Total Sales Today", total_sales)


# ══════════════════════════════════════════════════════════════════════════════
# PAGE 5 — LEDGER
# ══════════════════════════════════════════════════════════════════════════════

elif page == "📒 Ledger":
    st.title("📒 Stock Ledger")

    tab_recv, tab_issue, tab_pay, tab_prods, tab_export = st.tabs(
        ["📥 Receive Stock", "📤 Issue Stock", "💰 Payments", "🛒 Products", "📊 Export"]
    )

    # ── Receive Stock ──────────────────────────────────────────────────────────
    with tab_recv:
        st.subheader("📥 Record Stock Received")
        prods = list(ledger_products().keys())
        if not prods:
            st.warning("No products yet. Go to 'Products' tab to add some.")
        else:
            c1, c2 = st.columns(2)
            recv_prod   = c1.selectbox("Product", prods, key="recv_prod")
            recv_date   = c2.date_input("Date", value=date.today(), key="recv_date")
            c1, c2, c3 = st.columns(3)
            recv_qty    = c1.number_input("Quantity", min_value=1, value=1, key="recv_qty")
            recv_source = c2.text_input("Received From", key="recv_src")
            recv_price  = c3.number_input("Price per unit (₹)", min_value=0.0, value=0.0, key="recv_price")
            recv_notes  = st.text_input("Notes", key="recv_notes")

            if st.button("✅ Record Receive", type="primary", use_container_width=True):
                ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                tx = {
                    "ts": ts, "date": recv_date.strftime("%Y-%m-%d"),
                    "product": recv_prod, "type": "receive",
                    "qty": recv_qty, "source": recv_source, "notes": recv_notes,
                }
                data.setdefault("stock_transactions", []).append(tx)
                # Ledger entry
                total_val = round(recv_qty * recv_price, 2)
                prev_bal  = ledger_price_balance(recv_prod)
                entry = {
                    "ts": ts, "date": recv_date.strftime("%Y-%m-%d"),
                    "product": recv_prod, "type": "receive",
                    "qty": recv_qty, "price": recv_price,
                    "debit": total_val, "credit": 0.0,
                    "balance": round(prev_bal + total_val, 2),
                    "source": recv_source, "notes": recv_notes,
                }
                data.setdefault("ledger_entries", []).append(entry)
                persist()
                st.success(f"✅ Received {recv_qty} × {recv_prod} from {recv_source}")
                st.rerun()

        # Recent receives
        st.divider()
        st.subheader("Recent Receives")
        recv_txs = [t for t in stock_transactions() if t.get("type") == "receive"]
        if recv_txs:
            import pandas as pd
            df = pd.DataFrame(sorted(recv_txs, key=lambda x: x.get("ts",""), reverse=True)[:20])
            cols = [c for c in ["date","product","qty","source","notes"] if c in df.columns]
            st.dataframe(df[cols], use_container_width=True, hide_index=True)
        else:
            st.info("No receives recorded yet.")

    # ── Issue Stock ────────────────────────────────────────────────────────────
    with tab_issue:
        st.subheader("📤 Record Stock Issued")
        prods = list(ledger_products().keys())
        if not prods:
            st.warning("No products yet.")
        else:
            c1, c2 = st.columns(2)
            issue_prod = c1.selectbox("Product", prods, key="issue_prod")
            issue_date = c2.date_input("Date", value=date.today(), key="issue_date")
            c1, c2    = st.columns(2)
            issue_qty  = c1.number_input("Quantity", min_value=1, value=1, key="issue_qty")
            issue_to   = c2.text_input("Issued To", key="issue_to")
            issue_notes = st.text_input("Notes", key="issue_notes")

            cur_bal = ledger_stock_balance(issue_prod)
            st.info(f"Current stock balance of **{issue_prod}**: **{cur_bal}** units")

            if st.button("✅ Record Issue", type="primary", use_container_width=True):
                if issue_qty > cur_bal:
                    st.error(f"❌ Not enough stock! Available: {cur_bal}, Requested: {issue_qty}")
                else:
                    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    tx = {
                        "ts": ts, "date": issue_date.strftime("%Y-%m-%d"),
                        "product": issue_prod, "type": "issue",
                        "qty": issue_qty, "source": issue_to, "notes": issue_notes,
                    }
                    data.setdefault("stock_transactions", []).append(tx)
                    prod_price = float(ledger_products().get(issue_prod, {}).get("price", 0))
                    total_val  = round(issue_qty * prod_price, 2)
                    prev_bal   = ledger_price_balance(issue_prod)
                    entry = {
                        "ts": ts, "date": issue_date.strftime("%Y-%m-%d"),
                        "product": issue_prod, "type": "issue",
                        "qty": issue_qty, "price": prod_price,
                        "debit": 0.0, "credit": total_val,
                        "balance": round(prev_bal - total_val, 2),
                        "source": issue_to, "notes": issue_notes,
                    }
                    data.setdefault("ledger_entries", []).append(entry)
                    persist()
                    st.success(f"✅ Issued {issue_qty} × {issue_prod} to {issue_to}")
                    st.rerun()

        # Recent issues
        st.divider()
        st.subheader("Recent Issues")
        issue_txs = [t for t in stock_transactions() if t.get("type") == "issue"]
        if issue_txs:
            import pandas as pd
            df = pd.DataFrame(sorted(issue_txs, key=lambda x: x.get("ts",""), reverse=True)[:20])
            cols = [c for c in ["date","product","qty","source","notes"] if c in df.columns]
            st.dataframe(df[cols], use_container_width=True, hide_index=True)
        else:
            st.info("No issues recorded yet.")

    # ── Payments ───────────────────────────────────────────────────────────────
    with tab_pay:
        st.subheader("💰 Record Payment")
        prods = ["(General)"] + list(ledger_products().keys())
        c1, c2 = st.columns(2)
        pay_prod   = c1.selectbox("Product (or General)", prods, key="pay_prod")
        pay_date   = c2.date_input("Date", value=date.today(), key="pay_date")
        c1, c2     = st.columns(2)
        pay_amount = c1.number_input("Amount (₹)", min_value=0.0, value=0.0, key="pay_amount")
        pay_desc   = c2.text_input("Description / Paid To", key="pay_desc")
        pay_notes  = st.text_input("Notes", key="pay_notes")

        if st.button("✅ Record Payment", type="primary", use_container_width=True):
            if pay_amount > 0:
                ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                prod = pay_prod if pay_prod != "(General)" else ""
                pmt = {
                    "ts": ts, "date": pay_date.strftime("%Y-%m-%d"),
                    "product": prod, "amount": pay_amount,
                    "description": pay_desc, "notes": pay_notes,
                }
                data.setdefault("ledger_payments", []).append(pmt)
                # Ledger entry as credit
                if prod:
                    prev_bal = ledger_price_balance(prod)
                    entry = {
                        "ts": ts, "date": pay_date.strftime("%Y-%m-%d"),
                        "product": prod, "type": "payment",
                        "qty": 0, "price": 0.0,
                        "debit": 0.0, "credit": pay_amount,
                        "balance": round(prev_bal - pay_amount, 2),
                        "source": pay_desc, "notes": pay_notes,
                    }
                    data.setdefault("ledger_entries", []).append(entry)
                persist()
                st.success(f"✅ Payment of ₹{pay_amount} recorded")
                st.rerun()
            else:
                st.error("Amount must be greater than 0.")

        st.divider()
        st.subheader("Payment History")
        pmts = sorted(data.get("ledger_payments", []), key=lambda x: x.get("ts",""), reverse=True)
        if pmts:
            import pandas as pd
            df = pd.DataFrame(pmts)
            cols = [c for c in ["date","product","description","amount","notes"] if c in df.columns]
            st.dataframe(df[cols], use_container_width=True, hide_index=True)
            total_paid = sum(p.get("amount", 0) for p in pmts)
            st.metric("Total Payments", f"₹{total_paid:,.2f}")
        else:
            st.info("No payments recorded yet.")

    # ── Products ───────────────────────────────────────────────────────────────
    with tab_prods:
        st.subheader("🛒 Manage Products")

        with st.form("add_product_form"):
            c1, c2, c3 = st.columns(3)
            prod_name  = c1.text_input("Product Name")
            prod_price = c2.number_input("Price per unit (₹)", min_value=0.0, value=0.0)
            prod_unit  = c3.selectbox("Unit", ["pcs", "kg", "g", "L", "mL", "box", "dozen", "pack"])
            submitted  = st.form_submit_button("💾 Add / Update Product", use_container_width=True)
            if submitted:
                if prod_name.strip():
                    data.setdefault("ledger_products", {})[prod_name.strip()] = {
                        "price": prod_price,
                        "unit":  prod_unit,
                    }
                    persist()
                    st.success(f"✅ Product '{prod_name}' saved!")
                    st.rerun()
                else:
                    st.error("Product name cannot be empty.")

        st.divider()
        st.subheader("Current Products")
        prods = ledger_products()
        if prods:
            import pandas as pd
            rows = [{"Product": k, "Price (₹)": v.get("price", 0),
                     "Unit": v.get("unit", "pcs"),
                     "Current Stock": ledger_stock_balance(k)}
                    for k, v in sorted(prods.items())]
            st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

            # Delete product
            del_prod = st.selectbox("Delete Product", ["-- select --"] + list(prods.keys()), key="del_prod")
            if st.button("🗑 Delete Product", type="secondary"):
                if del_prod != "-- select --":
                    data["ledger_products"].pop(del_prod, None)
                    data["stock_transactions"] = [t for t in data.get("stock_transactions", [])
                                                  if t.get("product") != del_prod]
                    data["ledger_entries"] = [e for e in data.get("ledger_entries", [])
                                              if e.get("product") != del_prod]
                    persist()
                    st.success(f"🗑 Deleted: {del_prod}")
                    st.rerun()
        else:
            st.info("No products yet. Add one above.")

    # ── Export ─────────────────────────────────────────────────────────────────
    with tab_export:
        st.subheader("📊 Export Ledger to Excel")
        st.markdown("Export everything to a single Excel file with multiple sheets.")

        if st.button("📥 Generate Excel Export", type="primary", use_container_width=True):
            wb   = Workbook()
            hfil = PatternFill("solid", fgColor="273c75")
            hfnt = Font(bold=True, color="FFFFFF", size=11)
            ctr  = Alignment(horizontal="center", vertical="center")
            grn  = PatternFill("solid", fgColor="EAFAF1")
            red  = PatternFill("solid", fgColor="FDEDEC")
            even = PatternFill("solid", fgColor="EBF5FB")
            odd  = PatternFill("solid", fgColor="D6EAF8")

            # Sheet 1 – Stock Transactions
            ws1 = wb.active; ws1.title = "Stock Ledger"
            for ci, h in enumerate(["Date","Product","Type","Qty","Source","Notes"], 1):
                c = ws1.cell(row=1, column=ci, value=h)
                c.fill = hfil; c.font = hfnt; c.alignment = ctr
            for i, tx in enumerate(sorted(stock_transactions(),
                                          key=lambda x: (x.get("date",""), x.get("ts",""))), 2):
                fill = grn if tx.get("type") == "receive" else red
                for ci, val in enumerate([tx.get("date",""), tx.get("product",""),
                                          tx.get("type","").capitalize(), tx.get("qty",""),
                                          tx.get("source",""), tx.get("notes","")], 1):
                    c = ws1.cell(row=i, column=ci, value=val)
                    c.fill = fill; c.alignment = ctr
            for ci, w in enumerate([14,20,10,8,20,28], 1):
                ws1.column_dimensions[get_column_letter(ci)].width = w

            # Sheet 2 – Product Summary
            ws2 = wb.create_sheet("Product Summary")
            for ci, h in enumerate(["Product","Price (₹)","Unit","Stock Balance","Total Value (₹)"], 1):
                c = ws2.cell(row=1, column=ci, value=h)
                c.fill = hfil; c.font = hfnt; c.alignment = ctr
            for i, (prod, meta) in enumerate(sorted(ledger_products().items()), 2):
                bal   = ledger_stock_balance(prod)
                price = float(meta.get("price", 0))
                fill  = even if i % 2 == 0 else odd
                for ci, val in enumerate([prod, price, meta.get("unit","pcs"), bal, round(bal*price,2)], 1):
                    c = ws2.cell(row=i, column=ci, value=val)
                    c.fill = fill; c.alignment = ctr

            # Sheet 3 – Payments
            ws3 = wb.create_sheet("Payments")
            for ci, h in enumerate(["Date","Product","Description","Amount (₹)","Notes"], 1):
                c = ws3.cell(row=1, column=ci, value=h)
                c.fill = hfil; c.font = hfnt; c.alignment = ctr
            for i, pmt in enumerate(sorted(ledger_payments(),
                                           key=lambda x: (x.get("date",""), x.get("ts",""))), 2):
                fill = even if i % 2 == 0 else odd
                for ci, val in enumerate([pmt.get("date",""), pmt.get("product","(General)"),
                                          pmt.get("description",""), pmt.get("amount",0),
                                          pmt.get("notes","")], 1):
                    c = ws3.cell(row=i, column=ci, value=val)
                    c.fill = fill; c.alignment = ctr

            # Sheet 4 – Drop Log
            ws4 = wb.create_sheet("Drop Log")
            for ci, h in enumerate(["Date","Store","Product","Qty","Note","Timestamp"], 1):
                c = ws4.cell(row=1, column=ci, value=h)
                c.fill = hfil; c.font = hfnt; c.alignment = ctr
            for i, row in enumerate(sorted(data.get("drop_log",[]),
                                           key=lambda x: x.get("ts","")), 2):
                fill = even if i % 2 == 0 else odd
                for ci, val in enumerate([row.get("date",""), row.get("store",""),
                                          row.get("product",""), row.get("qty",""),
                                          row.get("note",""), row.get("ts","")], 1):
                    c = ws4.cell(row=i, column=ci, value=val)
                    c.fill = fill; c.alignment = ctr

            buf = io.BytesIO()
            wb.save(buf)
            buf.seek(0)
            st.download_button(
                "⬇️ Download Excel File",
                buf,
                file_name=f"stock_export_{date.today()}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
            st.success("✅ Excel ready! Click Download above.")
