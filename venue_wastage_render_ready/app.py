import streamlit as st
import sqlite3, hashlib, secrets, json, os
from pathlib import Path
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

# Use Render's persistent disk when WASTAGE_DATA_DIR=/var/data is configured.
# On your computer, the database remains beside app.py by default.
DATA_DIR = Path(os.environ.get("WASTAGE_DATA_DIR", str(Path(__file__).resolve().parent)))
DATA_DIR.mkdir(parents=True, exist_ok=True)
DB = DATA_DIR / "venue_wastage.db"
UK_TZ = ZoneInfo("Europe/London")


def current_work_week():
    """Current UK Friday 00:00 to next Friday 00:00 (exclusive)."""
    today = datetime.now(UK_TZ).date()
    return today - timedelta(days=(today.weekday() - 4) % 7)


def work_week_start(value):
    """Friday associated with a stored local timestamp."""
    date_value = pd.Timestamp(value).date()
    return date_value - timedelta(days=(date_value.weekday() - 4) % 7)


def week_label(start):
    return f"Fri {start:%d %b %Y} – Thu {(start + timedelta(days=6)):%d %b %Y}"


def filter_work_week(frame, start):
    if frame.empty:
        return frame
    timestamps = pd.to_datetime(frame["at"], errors="coerce")
    end = start + timedelta(days=7)
    return frame.loc[(timestamps.dt.date >= start) & (timestamps.dt.date < end)].copy()


def available_work_weeks(*frames):
    weeks = {current_work_week()}
    for frame in frames:
        if not frame.empty:
            for timestamp in pd.to_datetime(frame["at"], errors="coerce").dropna():
                weeks.add(work_week_start(timestamp))
    return sorted(weeks, reverse=True)


DEFAULT_VENUES = ["Hotshots", "Rosso", "Centre", "Reds"]
LINE_DEFAULTS = ["Coors", "Madri", "Stowford", "Carling", "Carling Darkfruits", "Guinness", "Alpaca"]
GENERAL_CATEGORIES = ["Draught", "Bottle beers", "Soft bottle", "Postmix", "Spirits", "Wine"]
ROLES = ["admin", "manager", "team_till"]
PERMISSION_LABELS = {
    "line_clean": "Record line-clean wastage",
    "general_waste": "Record general wastage",
    "history": "View wastage history",
    "stock": "Manage venue stock and values",
    "users": "Manage users and permissions",
    "venues": "Add, edit and deactivate venues",
}
ROLE_DEFAULTS = {
    "admin": {k: True for k in PERMISSION_LABELS},
    "manager": {"line_clean": True, "general_waste": True, "history": True, "stock": True, "users": False, "venues": False},
    "team_till": {"line_clean": True, "general_waste": True, "history": False, "stock": False, "users": False, "venues": False},
}

st.set_page_config(page_title="Venue Wastage", page_icon="🍺", layout="wide")
st.markdown("<style>.block-container{max-width:1200px;padding-top:1rem}div[data-testid='stMetric']{background:rgba(128,128,128,.08);border:1px solid rgba(128,128,128,.18);padding:12px;border-radius:12px}div.stButton>button{border-radius:10px;min-height:2.5rem}</style>", unsafe_allow_html=True)


def db():
    c = sqlite3.connect(DB)
    c.row_factory = sqlite3.Row
    c.execute("PRAGMA foreign_keys=ON")
    c.executescript("""
    CREATE TABLE IF NOT EXISTS venues(
      id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT UNIQUE NOT NULL, active INTEGER NOT NULL DEFAULT 1
    );
    CREATE TABLE IF NOT EXISTS users(
      id INTEGER PRIMARY KEY, username TEXT UNIQUE NOT NULL, salt TEXT NOT NULL,
      pwhash TEXT NOT NULL, role TEXT NOT NULL, active INTEGER NOT NULL DEFAULT 1
    );
    CREATE TABLE IF NOT EXISTS access(
      user_id INTEGER NOT NULL, venue TEXT NOT NULL, PRIMARY KEY(user_id,venue)
    );
    CREATE TABLE IF NOT EXISTS user_permissions(
      user_id INTEGER NOT NULL, permission TEXT NOT NULL, allowed INTEGER NOT NULL,
      PRIMARY KEY(user_id,permission)
    );
    CREATE TABLE IF NOT EXISTS stock(
      id INTEGER PRIMARY KEY, venue TEXT NOT NULL, section TEXT NOT NULL,
      category TEXT NOT NULL, name TEXT NOT NULL, unit TEXT NOT NULL,
      unit_value REAL NOT NULL DEFAULT 0, active INTEGER NOT NULL DEFAULT 1,
      UNIQUE(venue,section,category,name)
    );
    CREATE TABLE IF NOT EXISTS line_logs(
      id INTEGER PRIMARY KEY, venue TEXT NOT NULL, at TEXT NOT NULL, username TEXT NOT NULL,
      notes TEXT, first_taps INTEGER NOT NULL, second_taps INTEGER NOT NULL,
      total_pints INTEGER NOT NULL, detail TEXT NOT NULL
    );
    CREATE TABLE IF NOT EXISTS general_logs(
      id INTEGER PRIMARY KEY, venue TEXT NOT NULL, at TEXT NOT NULL, username TEXT NOT NULL,
      category TEXT NOT NULL, item TEXT NOT NULL, qty REAL NOT NULL, unit TEXT NOT NULL,
      unit_value REAL NOT NULL, total_value REAL NOT NULL, notes TEXT
    );
    """)
    # Upgrade older versions of the app in place.
    existing = {r["name"] for r in c.execute("SELECT name FROM venues").fetchall()}
    for name in DEFAULT_VENUES:
        if name not in existing:
            c.execute("INSERT OR IGNORE INTO venues(name) VALUES(?)", (name,))
    c.commit()
    return c


def hash_password(password, salt):
    return hashlib.pbkdf2_hmac("sha256", password.encode(), salt.encode(), 180000).hex()


def all_venues(active_only=True):
    c = db()
    q = "SELECT name FROM venues" + (" WHERE active=1" if active_only else "") + " ORDER BY name"
    result = [r["name"] for r in c.execute(q).fetchall()]
    c.close()
    return result


def seed_stock():
    c = db()
    for v in all_venues():
        for p in LINE_DEFAULTS:
            c.execute("INSERT OR IGNORE INTO stock(venue,section,category,name,unit,unit_value) VALUES(?,?,?,?,?,0)",
                      (v, "Line clean", "Line clean", p, "pint"))
    c.commit()
    c.close()


def get_access(uid):
    c = db()
    vals = [r["venue"] for r in c.execute("SELECT venue FROM access WHERE user_id=?", (uid,)).fetchall()]
    c.close()
    return [v for v in vals if v in all_venues()]


def get_permissions(uid, role):
    defaults = ROLE_DEFAULTS.get(role, ROLE_DEFAULTS["team_till"]).copy()
    c = db()
    rows = c.execute("SELECT permission,allowed FROM user_permissions WHERE user_id=?", (uid,)).fetchall()
    c.close()
    for r in rows:
        if r["permission"] in defaults:
            defaults[r["permission"]] = bool(r["allowed"])
    return defaults


def has(permission):
    u = st.session_state.user
    if u["role"] == "admin":
        return True
    return get_permissions(u["id"], u["role"]).get(permission, False)


def get_stock(venue, section, category=None):
    c = db()
    if category is None:
        rows = c.execute("SELECT * FROM stock WHERE venue=? AND section=? AND active=1 ORDER BY category,name",
                         (venue, section)).fetchall()
    else:
        rows = c.execute("SELECT * FROM stock WHERE venue=? AND section=? AND category=? AND active=1 ORDER BY name",
                         (venue, section, category)).fetchall()
    c.close()
    return [dict(r) for r in rows]


def export_frame(frame):
    out = frame.copy()
    if "at" in out:
        dt = pd.to_datetime(out.pop("at"), errors="coerce")
        out.insert(0, "Date", dt.dt.strftime("%d/%m/%Y"))
        out.insert(1, "Time", dt.dt.strftime("%H:%M:%S"))
    if "qty" in out:
        out.rename(columns={"qty": "Quantity (number of measures)", "unit": "Measure / unit", "unit_value": "Value per measure (£)", "total_value": "Total value (£)"}, inplace=True)
        # Quantity is always the number of measures recorded, not a volume conversion.
        # Volume is provided separately for products with a known ml measure.
        import re
        def ml_total(row):
            unit = str(row.get("Measure / unit", ""))
            match = re.search(r"(\d+(?:\.\d+)?)\s*ml\b", unit, re.I)
            return round(float(row["Quantity (number of measures)"]) * float(match.group(1)), 2) if match else None
        out.insert(out.columns.get_loc("Measure / unit") + 1, "Total volume (ml)", out.apply(ml_total, axis=1))
    return out


def excel_download(sheets):
    wb = Workbook()
    wb.remove(wb.active)
    for title, data in sheets.items():
        ws = wb.create_sheet(title[:31])
        clean = data.astype(object).where(pd.notna(data), None)
        ws.append(list(clean.columns))
        for row in clean.itertuples(index=False, name=None):
            ws.append([v.item() if hasattr(v, "item") else v for v in row])
        for cell in ws[1]:
            cell.fill = PatternFill("solid", fgColor="17365D")
            cell.font = Font(bold=True, color="FFFFFF")
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        for col in ws.columns:
            width = max((len(str(cell.value)) if cell.value is not None else 0) for cell in col)
            ws.column_dimensions[get_column_letter(col[0].column)].width = min(40, max(12, width + 2))
    output = BytesIO()
    wb.save(output)
    return output.getvalue()


def money(value):
    return f"£{float(value):,.2f}"


def logout():
    st.session_state.pop("user", None)
    st.rerun()


def save_line_clean(venue, username, first, second, notes):
    detail = []
    for product in first:
        f, s = int(first[product]) * 5, int(second.get(product, 0)) * 5
        detail.append({
            "Product": product, "First pull (pints)": f,
            "Second pull (pints)": s, "Total wasted (pints)": f+s
        })
    ft, staps = sum(first.values()), sum(second.values())
    c = db()
    c.execute("""INSERT INTO line_logs(venue,at,username,notes,first_taps,second_taps,total_pints,detail)
                 VALUES(?,?,?,?,?,?,?,?)""",
              (venue, datetime.now().isoformat(timespec="seconds"), username, notes,
               ft, staps, (ft+staps)*5, json.dumps(detail)))
    c.commit()
    c.close()


def set_user_permissions(c, uid, role, custom):
    c.execute("DELETE FROM user_permissions WHERE user_id=?", (uid,))
    defaults = ROLE_DEFAULTS[role]
    for key in PERMISSION_LABELS:
        val = bool(custom.get(key, defaults[key]))
        c.execute("INSERT INTO user_permissions(user_id,permission,allowed) VALUES(?,?,?)",
                  (uid, key, int(val)))


c = db()
user_count = c.execute("SELECT COUNT(*) FROM users").fetchone()[0]
c.close()

if user_count == 0:
    st.title("🍺 Venue Wastage — Initial Admin Setup")
    st.info("Create the first Admin account. Admins can add venues, create accounts and control permissions.")
    with st.form("first_admin"):
        username = st.text_input("Admin username")
        password = st.text_input("Password (at least 8 characters)", type="password")
        confirm = st.text_input("Confirm password", type="password")
        submit = st.form_submit_button("Create Admin account", type="primary")
    if submit:
        if not username.strip() or len(password) < 8:
            st.error("Enter a username and a password of at least 8 characters.")
        elif password != confirm:
            st.error("Passwords do not match.")
        else:
            salt = secrets.token_hex(16)
            c = db()
            cur = c.execute("INSERT INTO users(username,salt,pwhash,role) VALUES(?,?,?,'admin')",
                            (username.strip().lower(), salt, hash_password(password, salt)))
            uid = cur.lastrowid
            for v in all_venues():
                c.execute("INSERT OR IGNORE INTO access(user_id,venue) VALUES(?,?)", (uid, v))
            set_user_permissions(c, uid, "admin", ROLE_DEFAULTS["admin"])
            c.commit()
            c.close()
            seed_stock()
            st.success("Admin account created. Log in to continue.")
            st.rerun()
    st.stop()

seed_stock()

if "user" not in st.session_state:
    st.title("🍺 Venue Wastage")
    with st.form("login"):
        username = st.text_input("Username")
        password = st.text_input("Password", type="password")
        submit = st.form_submit_button("Log in", type="primary")
    if submit:
        c = db()
        u = c.execute("SELECT * FROM users WHERE username=? AND active=1", (username.strip().lower(),)).fetchone()
        c.close()
        if u and secrets.compare_digest(hash_password(password, u["salt"]), u["pwhash"]):
            st.session_state.user = dict(u)
            st.rerun()
        else:
            st.error("Login failed. Check your username and password.")
    st.stop()

user = st.session_state.user
is_admin = user["role"] == "admin"
is_manager = user["role"] == "manager"
assigned = all_venues() if is_admin else get_access(user["id"])
if not assigned:
    st.error("Your account has no active venue access. Contact an Admin.")
    if st.button("Log out"):
        logout()
    st.stop()

with st.sidebar:
    st.write(f"**{user['username']}**")
    st.caption(user["role"].replace("_", " ").title())
    venue = st.selectbox("Venue", assigned)
    pages = []
    if has("line_clean"): pages.append("Line clean wastage")
    if has("general_waste"): pages.append("General wastage")
    if has("history"):
        pages.extend(["Dashboard", "History"])
    if has("stock"): pages.append("Stock settings")
    if is_admin: pages.append("Admin control")
    if not pages:
        st.warning("No permissions have been assigned to this account.")
    page = st.radio("Section", pages) if pages else None
    if st.button("Log out"):
        logout()

st.title(f"{venue} — Venue Wastage")

if page == "Line clean wastage":
    st.subheader("Line clean wastage")
    st.caption("Each tap counts as 5 pints.")
    products = get_stock(venue, "Line clean")
    if not products:
        st.info("No active line-clean products. A manager or admin can add them in Stock settings.")
    else:
        left, right = st.columns(2)
        first, second = {}, {}
        with left:
            st.markdown("#### First pull-through")
            for item in products:
                first[item["name"]] = st.number_input(f"{item['name']} taps", 0, 100, 0, key=f"f_{venue}_{item['id']}")
        with right:
            st.markdown("#### Second pull-through")
            for item in products:
                second[item["name"]] = st.number_input(f"{item['name']} taps", 0, 100, 0, key=f"s_{venue}_{item['id']}")
        ft, staps = sum(first.values()), sum(second.values())
        a, b, c = st.columns(3)
        a.metric("First pull-through", f"{ft*5} pints", f"{ft} taps")
        b.metric("Second pull-through", f"{staps*5} pints", f"{staps} taps")
        c.metric("Total line-clean waste", f"{(ft+staps)*5} pints")
        breakdown = []
        for item in products:
            p = item["name"]
            f, s = first[p]*5, second[p]*5
            breakdown.append({"Product": p, "First pull (pints)": f,
                              "Second pull (pints)": s, "Total wasted (pints)": f+s})
        st.markdown("#### Pints wasted by product")
        st.dataframe(pd.DataFrame(breakdown), hide_index=True, use_container_width=True)
        notes = st.text_input("Notes (optional)", key="line_notes")
        if st.button("Save line clean", type="primary", disabled=(ft+staps == 0)):
            save_line_clean(venue, user["username"], first, second, notes)
            for item in products:
                st.session_state[f"f_{venue}_{item['id']}"] = 0
                st.session_state[f"s_{venue}_{item['id']}"] = 0
            st.success("Line clean saved.")
            st.rerun()

elif page == "General wastage":
    st.subheader("General wastage")
    category = st.selectbox("Category", GENERAL_CATEGORIES)
    items = get_stock(venue, "General", category)
    if not items:
        st.info("No active stock in this category. Add items in Stock settings.")
    else:
        options = {f"{x['name']} — {money(x['unit_value'])} per {x['unit']}": x for x in items}
        chosen = options[st.selectbox("Stock item", list(options))]
        qty = st.number_input(f"Quantity wasted ({chosen['unit']})", min_value=0.0, step=0.25)
        st.metric("Calculated wastage value", money(qty * chosen["unit_value"]))
        notes = st.text_input("Notes (optional)", key="general_notes")
        if st.button("Save general wastage", type="primary", disabled=qty <= 0):
            c = db()
            c.execute("""INSERT INTO general_logs(venue,at,username,category,item,qty,unit,unit_value,total_value,notes)
                         VALUES(?,?,?,?,?,?,?,?,?,?)""",
                      (venue, datetime.now().isoformat(timespec="seconds"), user["username"], category,
                       chosen["name"], qty, chosen["unit"], chosen["unit_value"],
                       qty*chosen["unit_value"], notes))
            c.commit()
            c.close()
            st.success("General wastage saved.")
            st.rerun()

elif page == "Dashboard":
    st.subheader(f"{venue} — Wastage dashboard")
    period = st.selectbox("Reporting period", ["Current working week (Fri–Thu)", "Previous working week (Fri–Thu)", "Last 7 days", "Last 30 days", "This month", "All time"])
    now = pd.Timestamp(datetime.now(UK_TZ).replace(tzinfo=None))
    since = (now - pd.Timedelta(days=7) if period == "Last 7 days" else
             now - pd.Timedelta(days=30) if period == "Last 30 days" else
             now.normalize().replace(day=1) if period == "This month" else None)
    c = db()
    lines = pd.read_sql_query("SELECT * FROM line_logs WHERE venue=? ORDER BY at DESC", c, params=(venue,))
    general = pd.read_sql_query("SELECT * FROM general_logs WHERE venue=? ORDER BY at DESC", c, params=(venue,))
    c.close()
    if period in ("Current working week (Fri–Thu)", "Previous working week (Fri–Thu)"):
        week_start = current_work_week()
        if period.startswith("Previous"):
            week_start -= timedelta(days=7)
        st.caption(f"Reporting week: {week_label(week_start)} (UK time)")
        lines = filter_work_week(lines, week_start)
        general = filter_work_week(general, week_start)
    for frame in (lines, general):
        if not frame.empty:
            frame["at"] = pd.to_datetime(frame["at"], errors="coerce")
            frame.dropna(subset=["at"], inplace=True)
            if since is not None:
                frame.drop(frame.index[frame["at"] < since], inplace=True)
    a,b,c,d = st.columns(4)
    a.metric("Line-clean waste", f"{lines['total_pints'].sum() if not lines.empty else 0:,.0f} pints")
    b.metric("General waste value", money(general['total_value'].sum() if not general.empty else 0))
    c.metric("Line-clean entries", len(lines))
    d.metric("General waste entries", len(general))
    st.markdown("### Daily trends")
    if not lines.empty:
        day_line = lines.assign(Day=lines["at"].dt.strftime("%Y-%m-%d")).groupby("Day")["total_pints"].sum()
        st.caption("Line-clean pints by day")
        st.bar_chart(day_line)
    if not general.empty:
        day_general = general.assign(Day=general["at"].dt.strftime("%Y-%m-%d")).groupby("Day")["total_value"].sum()
        st.caption("General wastage value (£) by day")
        st.bar_chart(day_general)
    if lines.empty and general.empty:
        st.info("No wastage recorded for this reporting period.")
    left,right = st.columns(2)
    with left:
        st.markdown("### General waste by category")
        if not general.empty:
            totals = general.groupby("category")["total_value"].sum().sort_values(ascending=False)
            st.bar_chart(totals)
            st.dataframe(totals.rename("Value (£)"), use_container_width=True)
        else:
            st.caption("No general wastage entries")
    with right:
        st.markdown("### Line-clean pints by product")
        totals = {}
        for detail in lines["detail"] if not lines.empty else []:
            try:
                for item in json.loads(detail):
                    product = item["Product"]
                    totals[product] = totals.get(product, 0) + float(item["Total wasted (pints)"])
            except (KeyError, ValueError, TypeError):
                continue
        if totals:
            product_totals = pd.Series(totals, name="Pints wasted").sort_values(ascending=False)
            st.bar_chart(product_totals)
            st.dataframe(product_totals, use_container_width=True)
        else:
            st.caption("No line-clean wastage entries")
    st.markdown("### Team activity")
    staff = pd.concat([lines[["username"]], general[["username"]]], ignore_index=True)
    if not staff.empty:
        st.dataframe(staff.groupby("username").size().rename("Records").sort_values(ascending=False), use_container_width=True)
    else:
        st.caption("No activity for this period")
    st.markdown("### Latest entries")
    recent = []
    for _, row in lines.head(10).iterrows():
        recent.append({"at": row["at"], "Type": "Line clean", "Team member": row["username"], "Detail": f"{row['total_pints']} pints"})
    for _, row in general.head(10).iterrows():
        recent.append({"at": row["at"], "Type": "General", "Team member": row["username"], "Detail": f"{row['qty']:g} × {row['unit']} {row['item']} — {money(row['total_value'])}"})
    if recent:
        st.dataframe(export_frame(pd.DataFrame(recent).sort_values("at", ascending=False).head(10)), hide_index=True, use_container_width=True)
    if not lines.empty or not general.empty:
        st.download_button("Download dashboard data (Excel)", excel_download({"Line clean": export_frame(lines), "General waste": export_frame(general)}), file_name=f"{venue}_dashboard.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

elif page == "History":
    st.subheader(f"{venue} — Weekly wastage history")
    st.caption("Working weeks run Friday 00:00 through Thursday 23:59:59, UK time. Older records are retained.")
    c = db()
    all_lines = pd.read_sql_query("SELECT * FROM line_logs WHERE venue=? ORDER BY at DESC", c, params=(venue,))
    all_general = pd.read_sql_query("SELECT * FROM general_logs WHERE venue=? ORDER BY at DESC", c, params=(venue,))
    c.close()
    weeks = available_work_weeks(all_lines, all_general)
    selected_week = st.selectbox("Select working week", weeks, format_func=week_label)
    lines = filter_work_week(all_lines, selected_week)
    general = filter_work_week(all_general, selected_week)
    st.caption(f"Selected: {week_label(selected_week)}")
    col1, col2, col3 = st.columns(3)
    col1.metric("Line-clean waste", f"{lines['total_pints'].sum() if not lines.empty else 0:,.0f} pints")
    col2.metric("General wastage value", money(general['total_value'].sum() if not general.empty else 0))
    col3.metric("Records", len(lines) + len(general))
    t1, t2 = st.tabs(["Line cleans", "General wastage"])
    with t1:
        if lines.empty:
            st.info("No line cleans saved for this week.")
        else:
            for _, row in lines.iterrows():
                with st.expander(f"{row['at'][:16].replace('T',' ')} — {row['total_pints']} pints — {row['username']}"):
                    st.dataframe(pd.DataFrame(json.loads(row["detail"])), hide_index=True, use_container_width=True)
                    if row["notes"]: st.caption(row["notes"])
    with t2:
        if general.empty:
            st.info("No general wastage saved for this week.")
        else:
            st.dataframe(export_frame(general[["at","username","category","item","qty","unit","unit_value","total_value","notes"]]), hide_index=True, use_container_width=True)
    if not lines.empty or not general.empty:
        st.download_button("Download selected week (Excel)",
            excel_download({"Line clean": export_frame(lines), "General waste": export_frame(general)}),
            file_name=f"{venue}_wastage_{selected_week.isoformat()}_to_{(selected_week + timedelta(days=6)).isoformat()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    else:
        st.info("No entries for this week yet. Select an older week to review its history.")

elif page == "Stock settings":
    st.subheader("Stock settings")
    st.caption("These changes apply only to the currently selected venue.")
    kind = st.radio("Stock list", ["Line-clean products", "General wastage products"], horizontal=True)
    if kind == "General wastage products":
        category = st.selectbox("Category", GENERAL_CATEGORIES)
        section, cat = "General", category
    else:
        section, cat = "Line clean", "Line clean"
    items = get_stock(venue, section, cat)
    for item in items:
        with st.expander(f"{item['name']} — {money(item['unit_value'])} per {item['unit']}"):
            with st.form(f"stock_{item['id']}"):
                name = st.text_input("Stock name", item["name"])
                unit = st.text_input("Unit", item["unit"])
                value = st.number_input("Value per unit (£)", min_value=0.0, value=float(item["unit_value"]), step=0.05)
                active = st.checkbox("Available to staff", value=bool(item["active"]))
                save = st.form_submit_button("Save stock item")
            if save:
                c = db()
                c.execute("UPDATE stock SET name=?,unit=?,unit_value=?,active=? WHERE id=?",
                          (name.strip(), unit.strip(), value, int(active), item["id"]))
                c.commit()
                c.close()
                st.success("Stock item updated.")
                st.rerun()
    st.markdown("#### Add a stock item")
    with st.form("add_stock"):
        new_name = st.text_input("New stock name")
        new_unit = st.text_input("Unit (e.g. pint, bottle, measure, glass)", "pint" if section == "Line clean" else "bottle")
        new_value = st.number_input("Value per unit (£)", min_value=0.0, value=0.0, step=0.05)
        add = st.form_submit_button("Add stock item")
    if add:
        if not new_name.strip() or not new_unit.strip():
            st.error("Enter a stock name and unit.")
        else:
            c = db()
            try:
                c.execute("INSERT INTO stock(venue,section,category,name,unit,unit_value) VALUES(?,?,?,?,?,?)",
                          (venue, section, cat, new_name.strip(), new_unit.strip(), new_value))
                c.commit()
                st.success("Stock item added.")
                st.rerun()
            except sqlite3.IntegrityError:
                st.error("That stock item already exists at this venue.")
            c.close()

elif page == "Admin control" and is_admin:
    st.subheader("Admin control")
    tabs = st.tabs(["Venues", "Users & permissions"])
    with tabs[0]:
        st.markdown("#### Add a venue")
        with st.form("add_venue"):
            new_venue = st.text_input("Venue name")
            add_venue = st.form_submit_button("Add venue", type="primary")
        if add_venue:
            if not new_venue.strip():
                st.error("Enter a venue name.")
            else:
                c = db()
                try:
                    c.execute("INSERT INTO venues(name) VALUES(?)", (new_venue.strip(),))
                    c.commit()
                    c.close()
                    seed_stock()
                    st.success(f"{new_venue.strip()} added.")
                    st.rerun()
                except sqlite3.IntegrityError:
                    c.close()
                    st.error("A venue with that name already exists.")
        st.markdown("#### Existing venues")
        c = db()
        venue_rows = c.execute("SELECT * FROM venues ORDER BY name").fetchall()
        c.close()
        for vr in venue_rows:
            with st.expander(f"{vr['name']} — {'Active' if vr['active'] else 'Inactive'}"):
                with st.form(f"venue_{vr['id']}"):
                    venue_name = st.text_input("Venue name", vr["name"])
                    venue_active = st.checkbox("Venue active / available for selection", bool(vr["active"]))
                    save_venue = st.form_submit_button("Save venue")
                if save_venue:
                    if not venue_name.strip():
                        st.error("Venue name cannot be blank.")
                    else:
                        c = db()
                        try:
                            c.execute("UPDATE venues SET name=?,active=? WHERE id=?",
                                      (venue_name.strip(), int(venue_active), vr["id"]))
                            if venue_name.strip() != vr["name"]:
                                c.execute("UPDATE access SET venue=? WHERE venue=?", (venue_name.strip(), vr["name"]))
                                c.execute("UPDATE stock SET venue=? WHERE venue=?", (venue_name.strip(), vr["name"]))
                                c.execute("UPDATE line_logs SET venue=? WHERE venue=?", (venue_name.strip(), vr["name"]))
                                c.execute("UPDATE general_logs SET venue=? WHERE venue=?", (venue_name.strip(), vr["name"]))
                            c.commit()
                            c.close()
                            st.success("Venue updated.")
                            st.rerun()
                        except sqlite3.IntegrityError:
                            c.close()
                            st.error("That venue name is already in use.")
    with tabs[1]:
        st.markdown("#### Create an account")
        with st.form("create_user"):
            new_username = st.text_input("Username")
            new_password = st.text_input("Temporary password (minimum 8 characters)", type="password")
            new_role = st.selectbox("Account role", ROLES, format_func=lambda x: {"admin":"Admin","manager":"Manager","team_till":"Team / Till"}[x])
            venue_access = st.multiselect("Venue access", all_venues(), default=[venue])
            st.caption("You can customise the individual permissions after creating the account.")
            create_user = st.form_submit_button("Create account")
        if create_user:
            if not new_username.strip() or len(new_password) < 8 or not venue_access:
                st.error("Enter a username, password of at least 8 characters and at least one venue.")
            else:
                salt = secrets.token_hex(16)
                c = db()
                try:
                    cur = c.execute("INSERT INTO users(username,salt,pwhash,role) VALUES(?,?,?,?)",
                                    (new_username.strip().lower(), salt, hash_password(new_password, salt), new_role))
                    uid = cur.lastrowid
                    for v in venue_access:
                        c.execute("INSERT INTO access(user_id,venue) VALUES(?,?)", (uid, v))
                    set_user_permissions(c, uid, new_role, ROLE_DEFAULTS[new_role])
                    c.commit()
                    c.close()
                    st.success("Account created.")
                    st.rerun()
                except sqlite3.IntegrityError:
                    c.close()
                    st.error("That username already exists.")
        st.markdown("#### Existing accounts")
        c = db()
        users = c.execute("SELECT * FROM users ORDER BY username").fetchall()
        c.close()
        for person in users:
            with st.expander(f"{person['username']} — {person['role'].replace('_',' ').title()}"):
                with st.form(f"edit_user_{person['id']}"):
                    role = st.selectbox("Role", ROLES, index=ROLES.index(person["role"]),
                                        format_func=lambda x: {"admin":"Admin","manager":"Manager","team_till":"Team / Till"}[x])
                    access = st.multiselect("Venue access", all_venues(), default=get_access(person["id"]))
                    active = st.checkbox("Account enabled", bool(person["active"]))
                    current_perms = get_permissions(person["id"], person["role"])
                    st.markdown("**Individual permissions**")
                    perms = {}
                    for key, label in PERMISSION_LABELS.items():
                        default_val = ROLE_DEFAULTS[role][key] if role != person["role"] else current_perms[key]
                        # Admin-only controls remain protected even if a non-admin role is edited.
                        if key in ("users", "venues") and role != "admin":
                            default_val = False
                        perms[key] = st.checkbox(label, value=default_val, key=f"perm_{person['id']}_{key}")
                    reset_password = st.text_input("New password (leave blank to keep current)", type="password")
                    save_user = st.form_submit_button("Save account and permissions")
                if save_user:
                    if not access:
                        st.error("Assign at least one active venue.")
                    elif reset_password and len(reset_password) < 8:
                        st.error("New password must be at least 8 characters.")
                    elif person["id"] == user["id"] and (role != "admin" or not active):
                        st.error("You cannot remove or disable your own Admin access.")
                    else:
                        if role != "admin":
                            perms["users"] = False
                            perms["venues"] = False
                        c = db()
                        c.execute("UPDATE users SET role=?,active=? WHERE id=?", (role, int(active), person["id"]))
                        c.execute("DELETE FROM access WHERE user_id=?", (person["id"],))
                        for v in access:
                            c.execute("INSERT INTO access(user_id,venue) VALUES(?,?)", (person["id"], v))
                        set_user_permissions(c, person["id"], role, perms)
                        if reset_password:
                            salt = secrets.token_hex(16)
                            c.execute("UPDATE users SET salt=?,pwhash=? WHERE id=?",
                                      (salt, hash_password(reset_password, salt), person["id"]))
                        c.commit()
                        c.close()
                        st.success("Account updated.")
                        st.rerun()

st.caption("Admin controls venue creation and account permissions. Managers can manage stock for assigned venues. Team / Till users receive only the permissions enabled by Admin.")
