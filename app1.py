import io
import re
from datetime import datetime, date, time, timedelta

import pandas as pd
import streamlit as st

from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils.dataframe import dataframe_to_rows


# =========================================================
# Password protection (Streamlit Secrets)
# =========================================================
def require_password():
    """
    Set in Streamlit Cloud -> App -> Settings -> Secrets:
    APP_PASSWORD="yourpassword"
    """
    pw = st.secrets.get("APP_PASSWORD", None)
    if not pw:
        st.warning(
            "Password is not configured. Add APP_PASSWORD in Streamlit Secrets to enable protection."
        )
        return True

    if "authed" not in st.session_state:
        st.session_state.authed = False

    if st.session_state.authed:
        return True

    with st.form("login"):
        entered = st.text_input("Password", type="password")
        ok = st.form_submit_button("Log in")
        if ok:
            if entered == pw:
                st.session_state.authed = True
                st.success("Logged in.")
                return True
            st.error("Incorrect password.")
    st.stop()


# =========================================================
# Robust sheet + column detection
# =========================================================
def normalize(s: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", str(s).strip().lower())


def find_sheet(xls: pd.ExcelFile, candidates):
    names = {normalize(n): n for n in xls.sheet_names}
    for c in candidates:
        key = normalize(c)
        if key in names:
            return names[key]
    # fuzzy contains
    for n in xls.sheet_names:
        nn = normalize(n)
        for c in candidates:
            if normalize(c) in nn or nn in normalize(c):
                return n
    return None


def pick_col(df: pd.DataFrame, candidates, required=True):
    cols = {normalize(c): c for c in df.columns}
    for cand in candidates:
        key = normalize(cand)
        if key in cols:
            return cols[key]
    # fuzzy contains
    for c in df.columns:
        nc = normalize(c)
        for cand in candidates:
            if normalize(cand) in nc:
                return c
    if required:
        raise KeyError(
            f"Could not find required column among: {candidates}. Available: {list(df.columns)}"
        )
    return None


def to_time(x):
    if pd.isna(x) or x == "":
        return None
    if isinstance(x, time):
        return x
    if isinstance(x, datetime):
        return x.time()
    if isinstance(x, (float, int)):
        # Excel time is fraction of day
        seconds = int(round(float(x) * 86400))
        return (datetime(2000, 1, 1) + timedelta(seconds=seconds)).time()
    s = str(x).strip()
    for fmt in ("%H:%M", "%H.%M", "%I:%M%p", "%I:%M %p"):
        try:
            return datetime.strptime(s, fmt).time()
        except ValueError:
            pass
    raise ValueError(f"Unrecognized time format: {x}")


def to_date(x):
    if pd.isna(x) or x == "":
        return None
    if isinstance(x, (datetime, pd.Timestamp)):
        return x.date()
    if isinstance(x, date):
        return x
    return pd.to_datetime(x).date()


# =========================================================
# Read template
# =========================================================
def read_template(uploaded_bytes: bytes):
    xls = pd.ExcelFile(io.BytesIO(uploaded_bytes))

    staff_sheet = find_sheet(xls, ["Staff", "Skills", "People"])
    hours_sheet = find_sheet(xls, ["WorkingHours", "Hours", "Availability"])
    hols_sheet = find_sheet(xls, ["Holidays", "Leave", "Absence"])
    params_sheet = find_sheet(xls, ["Parameters", "Params", "Rules", "Config"])

    if not staff_sheet:
        raise ValueError(f"Could not find Staff/Skills sheet. Found: {xls.sheet_names}")
    if not hours_sheet:
        raise ValueError(f"Could not find WorkingHours/Hours sheet. Found: {xls.sheet_names}")
    if not params_sheet:
        raise ValueError(f"Could not find Parameters sheet. Found: {xls.sheet_names}")

    staff_df = pd.read_excel(xls, sheet_name=staff_sheet)
    hours_df = pd.read_excel(xls, sheet_name=hours_sheet)
    hols_df = pd.read_excel(xls, sheet_name=hols_sheet) if hols_sheet else pd.DataFrame()
    params_df = pd.read_excel(xls, sheet_name=params_sheet)

    # ---- Staff + skills ----
    name_c = pick_col(staff_df, ["Name", "StaffName", "Employee", "Person"])
    home_c = pick_col(staff_df, ["HomeSite", "Site", "BaseSite"], required=False)

    staff_df = staff_df.copy()
    staff_df["Name"] = staff_df[name_c].astype(str).str.strip()
    staff_df["HomeSite"] = (
        staff_df[home_c].astype(str).str.strip().fillna("") if home_c else ""
    ).str.upper()

    # treat any additional columns as skill flags (Y/N)
    ignore = {normalize(name_c)}
    if home_c:
        ignore.add(normalize(home_c))

    def yn(v):
        if pd.isna(v):
            return False
        s = str(v).strip().lower()
        return s in {"y", "yes", "true", "1", "t"}

    for c in staff_df.columns:
        if normalize(c) in ignore:
            continue
        if staff_df[c].notna().any():
            staff_df[c] = staff_df[c].apply(yn)

    # Carol Church override
    carol_flag_col = None
    for c in staff_df.columns:
        if normalize(c) in {"iscarolchurch", "carolchurch", "carol"}:
            carol_flag_col = c
            break
    if carol_flag_col:
        staff_df["IsCarolChurch"] = staff_df[carol_flag_col].apply(bool)
    else:
        staff_df["IsCarolChurch"] = staff_df["Name"].str.lower().eq("carol church")

    # ---- Working hours ----
    hours_df = hours_df.copy()
    h_name = pick_col(hours_df, ["Name", "StaffName", "Employee", "Person"])
    hours_df["Name"] = hours_df[h_name].astype(str).str.strip()

    days = ["Mon", "Tue", "Wed", "Thu", "Fri"]
    for d in days:
        sc = pick_col(hours_df, [f"{d}Start", f"{d} Start", f"{d}_Start"], required=False)
        ec = pick_col(hours_df, [f"{d}End", f"{d} End", f"{d}_End"], required=False)
        hours_df[f"{d}Start"] = hours_df[sc].apply(to_time) if sc else None
        hours_df[f"{d}End"] = hours_df[ec].apply(to_time) if ec else None

    # ---- Holidays ----
    hols = []
    if not hols_df.empty:
        hn = pick_col(hols_df, ["Name", "StaffName", "Employee", "Person"], required=False)
        hs = pick_col(hols_df, ["StartDate", "Start"], required=False)
        he = pick_col(hols_df, ["EndDate", "End"], required=False)
        if hn and hs and he:
            for _, r in hols_df.iterrows():
                hols.append((str(r[hn]).strip(), to_date(r[hs]), to_date(r[he])))

    # ---- Parameters (robust) ----
    params = {}
    rule_c = pick_col(params_df, ["Rule", "Parameter", "Key", "Setting"], required=False)
    val_c = pick_col(params_df, ["Value", "Val"], required=False)

    if rule_c and val_c:
        for _, r in params_df.iterrows():
            k = str(r[rule_c]).strip()
            if k and k.lower() != "nan":
                params[k] = r[val_c]
    elif params_df.shape[1] >= 2:
        c1, c2 = params_df.columns[:2]
        for _, r in params_df.iterrows():
            k = str(r[c1]).strip()
            if k and k.lower() != "nan":
                params[k] = r[c2]

    return staff_df, hours_df, hols, params, xls.sheet_names


# =========================================================
# Rota rules + colours
# =========================================================
ROLE_COLORS = {
    "FrontDesk": "FFF2CC",
    "Triage_Admin": "D9EAD3",
    "Email_Box": "CFE2F3",
    "Phones": "C9DAF8",
    "Bookings": "FCE5CD",
    "EMIS": "EAD1DC",
    "Docman_PSA": "D0E0E3",
    "Docman_Awaiting": "D0E0E3",
    "Break": "E6E6E6",
    "Unassigned": "FFFFFF",
}

DAY_START = time(8, 0)
DAY_END = time(18, 30)
SLOT_MINUTES = 30

BREAK_WINDOW = (time(11, 30), time(14, 0))
BREAK_LEN_HRS = 0.5  # 30 mins
BREAK_REQUIRED_IF_SHIFT_GT = 6.0

MAX_FRONTDESK_BLOCK_HRS = 2.5
MAX_TRIAGE_BLOCK_HRS = 3.0

# Mandatory coverage
MANDATORY = [
    ("FrontDesk", "SLGP", DAY_START, DAY_END, 1),
    ("FrontDesk", "JEN", DAY_START, DAY_END, 1),
    ("FrontDesk", "BGS", DAY_START, DAY_END, 1),
    ("Triage_Admin", "SLGP", DAY_START, time(16, 0), 1),
    ("Triage_Admin", "JEN", DAY_START, time(16, 0), 1),
    ("Email_Box", None, DAY_START, DAY_END, 1),  # only JEN/BGS allowed
    ("Phones", None, DAY_START, DAY_END, 2),     # min 2, allow 3 if spare
]

# Filler tasks (allocated by remaining capacity)
FILLER = ["Bookings", "EMIS", "Docman_PSA", "Docman_Awaiting"]


# =========================================================
# Helpers
# =========================================================
def ensure_monday(d: date) -> date:
    return d - timedelta(days=d.weekday())


def timeslots(day_start=DAY_START, day_end=DAY_END, step_min=SLOT_MINUTES):
    cur = datetime(2000, 1, 1, day_start.hour, day_start.minute)
    end = datetime(2000, 1, 1, day_end.hour, day_end.minute)
    out = []
    while cur < end:
        out.append(cur.time())
        cur += timedelta(minutes=step_min)
    return out


def t_in_range(t, a, b):
    return (t >= a) and (t < b)


def hours_between(t1, t2):
    dt1 = datetime(2000, 1, 1, t1.hour, t1.minute)
    dt2 = datetime(2000, 1, 1, t2.hour, t2.minute)
    return (dt2 - dt1).total_seconds() / 3600.0


def is_on_holiday(name, d, hols):
    n0 = name.strip().lower()
    for n, s, e in hols:
        if n.strip().lower() == n0 and s and e and s <= d <= e:
            return True
    return False


def get_skill(staff_row: pd.Series, skill_name: str) -> bool:
    """
    Skill column matching by normalized name.
    If the skill column doesn't exist, return False (safe).
    """
    cols = {normalize(c): c for c in staff_row.index}
    key = normalize(skill_name)
    if key in cols:
        return bool(staff_row[cols[key]])
    return False


def skill_allowed(staff_row: pd.Series, role: str) -> bool:
    # flexible mapping (works even if your template uses slightly different headings)
    if role == "FrontDesk":
        return any(get_skill(staff_row, x) for x in ["FrontDesk", "Front Desk", "Reception"])
    if role == "Triage_Admin":
        return any(get_skill(staff_row, x) for x in ["Triage", "AdminTriage", "Admin Triage"])
    if role == "Email_Box":
        return any(get_skill(staff_row, x) for x in ["Email", "EmailBox", "Emails", "Email Box"])
    if role == "Phones":
        return any(get_skill(staff_row, x) for x in ["Phones", "Phone"])
    if role == "Bookings":
        return any(get_skill(staff_row, x) for x in ["Bookings", "Booking"])
    if role == "EMIS":
        return get_skill(staff_row, "EMIS")
    if role == "Docman_PSA":
        return any(get_skill(staff_row, x) for x in ["Docman_PSA", "Docman PSA", "PSA", "DocmanPSA"])
    if role == "Docman_Awaiting":
        return any(get_skill(staff_row, x) for x in ["Docman_Awaiting", "Awaiting Response", "AwaitingResponse", "DocmanAwaiting"])
    return False


def site_restriction_ok(staff_row: pd.Series, role: str, site: str | None) -> bool:
    home = str(staff_row.get("HomeSite", "")).strip().upper()

    # Site-specific roles must be worked by that site (per your rule)
    if role in {"FrontDesk", "Triage_Admin"} and site:
        return home == site

    # Email/Phones/EMIS/Docman only JEN+BGS
    if role in {"Email_Box", "Phones", "EMIS", "Docman_PSA", "Docman_Awaiting"}:
        return home in {"JEN", "BGS"}

    # Bookings mainly SLGP
    if role == "Bookings":
        return home == "SLGP"

    return True


# =========================================================
# Availability + scheduling
# =========================================================
def build_availability(staff_df, hours_df, hols, week_start: date):
    hmap = {r["Name"]: r for _, r in hours_df.iterrows()}
    days = ["Mon", "Tue", "Wed", "Thu", "Fri"]
    out = {}
    for i, dname in enumerate(days):
        d = week_start + timedelta(days=i)
        available = []
        for _, s in staff_df.iterrows():
            nm = s["Name"]
            if is_on_holiday(nm, d, hols):
                continue
            hr = hmap.get(nm)
            if hr is None:
                continue
            stt = hr.get(f"{dname}Start")
            end = hr.get(f"{dname}End")
            if stt and end:
                available.append(nm)
        out[d] = available
    return out, hmap


def staff_work_window(hmap_row, dname):
    return hmap_row.get(f"{dname}Start"), hmap_row.get(f"{dname}End")


def rota_generate_one_week(
    staff_df, hours_df, hols, week_start: date,
    phones_max: int = 3,
    fairness_state=None
):
    """
    Returns:
      assign_grid: dict[(date, time)] -> list of assignments (rows)
      long_rows:   list[dict] (Date, Day, Time, Name, HomeSite, Role, Site, Hours)
      gaps_rows:   list[dict] (Date, Time, Role, Site, Issue)
      fairness_state: updated cumulative fairness
    """
    if fairness_state is None:
        fairness_state = {
            "role_hours": {},   # (name, role, site)->hours
            "total_hours": {},  # name->hours
        }

    slots = timeslots()
    days = ["Mon", "Tue", "Wed", "Thu", "Fri"]
    dates = [week_start + timedelta(days=i) for i in range(5)]

    availability, hmap = build_availability(staff_df, hours_df, hols, week_start)
    staff_by_name = {r["Name"]: r for _, r in staff_df.iterrows()}

    # per-slot assignments: list of dicts
    assign_grid = {}  # (d, t) -> list[{Role, Site, Name}]
    gaps_rows = []

    # track block lengths (front desk / triage)
    fd_block = {}    # (d, name, site) -> hours in current continuous run
    tr_block = {}    # (d, name, site) -> hours in current continuous run

    # day hours (mandatory only) + break marker
    day_hours_mand = {}  # (d, name) -> hours
    break_done = {}      # (d, name) -> bool

    def can_work(name, d, t):
        hr = hmap.get(name)
        if hr is None:
            return False
        dname = days[(d - week_start).days]
        stt, end = staff_work_window(hr, dname)
        return bool(stt and end and (t >= stt) and (t < end))

    def shift_len(name, d):
        hr = hmap.get(name)
        if hr is None:
            return 0.0
        dname = days[(d - week_start).days]
        stt, end = staff_work_window(hr, dname)
        if not stt or not end:
            return 0.0
        return hours_between(stt, end)

    def score(name, role, site):
        # lower is better
        base = fairness_state["role_hours"].get((name, role, site or ""), 0.0)
        base += 0.25 * fairness_state["total_hours"].get(name, 0.0)
        # prioritize SLGP Bookings “majority”: reduce score for SLGP staff on Bookings later in filler
        return base

    def pick_candidate(d, t, role, site, used_names: set[str]):
        # Carol rule: if Carol is working, she is always FrontDesk at her site
        if role == "FrontDesk":
            for _, sr in staff_df.iterrows():
                if sr.get("IsCarolChurch", False) and can_work(sr["Name"], d, t):
                    if str(sr.get("HomeSite", "")).upper() == str(site).upper():
                        return sr["Name"]

        candidates = []
        for name in availability.get(d, []):
            if name in used_names:
                continue
            if not can_work(name, d, t):
                continue

            sr = staff_by_name[name]

            if not site_restriction_ok(sr, role, site):
                continue
            if not skill_allowed(sr, role):
                continue

            # block limits
            if role == "FrontDesk":
                cur = fd_block.get((d, name, site), 0.0)
                if cur + 0.5 > MAX_FRONTDESK_BLOCK_HRS:
                    continue
            if role == "Triage_Admin":
                cur = tr_block.get((d, name, site), 0.0)
                if cur + 0.5 > MAX_TRIAGE_BLOCK_HRS:
                    continue

            candidates.append(name)

        if not candidates:
            return None

        candidates.sort(key=lambda n: (score(n, role, site), day_hours_mand.get((d, n), 0.0)))
        return candidates[0]

    def assign_one(d, t, role, site, used_names: set[str], label=None):
        name = pick_candidate(d, t, role, site, used_names)
        if not name:
            gaps_rows.append({
                "Date": d,
                "Time": t.strftime("%H:%M"),
                "Role": role if not label else label,
                "Site": site or "",
                "Issue": "No suitable staff available",
            })
            return None

        assign_grid.setdefault((d, t), []).append({
            "Role": role,
            "Site": site or "",
            "Name": name,
            "Label": label or role,
        })
        used_names.add(name)
        day_hours_mand[(d, name)] = day_hours_mand.get((d, name), 0.0) + 0.5

        # update blocks
        if role == "FrontDesk":
            fd_block[(d, name, site)] = fd_block.get((d, name, site), 0.0) + 0.5
        if role == "Triage_Admin":
            tr_block[(d, name, site)] = tr_block.get((d, name, site), 0.0) + 0.5

        return name

    # --------- Mandatory coverage per slot ----------
    for d in dates:
        for t in slots:
            used = set()

            # reset blocks for staff not on role this slot (prevents carryover)
            # we reset by re-initialising per slot below after assignment summary
            # (kept simple; continuous runs tracked only while they keep being picked)

            # Schedule mandatory roles
            for role, site, r0, r1, need in MANDATORY:
                if not t_in_range(t, r0, r1):
                    continue

                if role == "Phones":
                    # enforce min 2
                    for i in range(need):
                        assign_one(d, t, "Phones", site, used, label=f"Phones_{i+1}")
                    # optional 3rd phone if phones_max==3 and spare
                    if phones_max >= 3:
                        assign_one(d, t, "Phones", site, used, label="Phones_3")
                    continue

                if role == "Email_Box":
                    assign_one(d, t, "Email_Box", None, used)
                    continue

                # site-specific mandatory
                assign_one(d, t, role, site, used)

            # If someone not on FrontDesk/Triage in this slot, reset their block counters
            slot = assign_grid.get((d, t), [])
            fd_now = {(a["Name"], a["Site"]) for a in slot if a["Role"] == "FrontDesk"}
            tr_now = {(a["Name"], a["Site"]) for a in slot if a["Role"] == "Triage_Admin"}

            for name in availability.get(d, []):
                # reset FrontDesk blocks for any site
                for site in ["SLGP", "JEN", "BGS"]:
                    if (name, site) not in fd_now:
                        fd_block[(d, name, site)] = 0.0
                    if (name, site) not in tr_now:
                        tr_block[(d, name, site)] = 0.0

    # --------- Break placement + filler allocations ----------
    long_rows = []
    filler_rows = []
    break_rows = []

    # Build longform rows from mandatory grid
    for (d, t), assigns in assign_grid.items():
        for a in assigns:
            nm = a["Name"]
            home = str(staff_by_name[nm].get("HomeSite", "")).upper()
            long_rows.append({
                "Date": d,
                "Day": d.strftime("%a"),
                "Time": t.strftime("%H:%M"),
                "Name": nm,
                "HomeSite": home,
                "Role": a["Label"],
                "RoleBase": a["Role"],
                "Site": a["Site"],
                "Hours": 0.5,
            })

    # Break + filler by staff day
    for i, d in enumerate(dates):
        dname = days[i]
        for nm in availability.get(d, []):
            if is_on_holiday(nm, d, hols):
                continue
            if nm not in staff_by_name:
                continue

            sr = staff_by_name[nm]
            home = str(sr.get("HomeSite", "")).upper()

            # day shift length
            sh = shift_len(nm, d)
            if sh <= 0:
                continue

            mand = day_hours_mand.get((d, nm), 0.0)
            need_break = sh > BREAK_REQUIRED_IF_SHIFT_GT

            # place break inside 11:30-14:00 if needed
            if need_break:
                placed = False
                # find a free slot where they are not already in mandatory assignment
                for t in slots:
                    if not t_in_range(t, BREAK_WINDOW[0], BREAK_WINDOW[1]):
                        continue
                    # are they already assigned in mandatory at this time?
                    slot_assigns = assign_grid.get((d, t), [])
                    if any(a["Name"] == nm for a in slot_assigns):
                        continue
                    # place break
                    break_done[(d, nm)] = True
                    placed = True
                    break_rows.append({
                        "Date": d,
                        "Day": d.strftime("%a"),
                        "Time": t.strftime("%H:%M"),
                        "Name": nm,
                        "HomeSite": home,
                        "Role": "Break",
                        "RoleBase": "Break",
                        "Site": "",
                        "Hours": BREAK_LEN_HRS,
                    })
                    break

                if not placed:
                    gaps_rows.append({
                        "Date": d,
                        "Time": "",
                        "Role": "Break",
                        "Site": home,
                        "Issue": f"{nm} shift > 6h but no free 30-min slot available in 11:30–14:00",
                    })

            # remaining capacity after mandatory + break
            used = mand + (BREAK_LEN_HRS if break_done.get((d, nm), False) else 0.0)
            remaining = max(0.0, sh - used)
            if remaining <= 0:
                continue

            # Allocate filler tasks by policy
            alloc = []

            if home == "SLGP" and skill_allowed(sr, "Bookings"):
                alloc = [("Bookings", remaining)]
            elif home in {"JEN", "BGS"}:
                parts = []
                if skill_allowed(sr, "EMIS"):
                    parts.append("EMIS")
                if skill_allowed(sr, "Docman_PSA"):
                    parts.append("Docman_PSA")
                if skill_allowed(sr, "Docman_Awaiting"):
                    parts.append("Docman_Awaiting")
                if not parts:
                    alloc = [("Unassigned", remaining)]
                else:
                    per = remaining / len(parts)
                    alloc = [(p, per) for p in parts]
            else:
                alloc = [("Unassigned", remaining)]

            for role, hrs in alloc:
                filler_rows.append({
                    "Date": d,
                    "Day": d.strftime("%a"),
                    "Time": "",  # filler not per-slot
                    "Name": nm,
                    "HomeSite": home,
                    "Role": role,
                    "RoleBase": role,
                    "Site": home,
                    "Hours": float(hrs),
                })

    # Combine longform outputs
    all_rows = long_rows + break_rows + filler_rows
    df_long = pd.DataFrame(all_rows)

    # Update fairness across weeks
    for _, r in df_long.iterrows():
        nm = r["Name"]
        role_base = r["RoleBase"]
        site = r.get("Site", "") or ""
        hrs = float(r["Hours"])
        fairness_state["role_hours"][(nm, role_base, site)] = fairness_state["role_hours"].get(
            (nm, role_base, site), 0.0
        ) + hrs
        fairness_state["total_hours"][nm] = fairness_state["total_hours"].get(nm, 0.0) + hrs

    return assign_grid, df_long, pd.DataFrame(gaps_rows), fairness_state


# =========================================================
# Excel export
# =========================================================
def fill_for_role(role_base: str):
    color = ROLE_COLORS.get(role_base, "FFFFFF")
    return PatternFill("solid", fgColor=color)


def autosize_columns(ws, max_width=55):
    for col_cells in ws.columns:
        col_letter = col_cells[0].column_letter
        lengths = []
        for c in col_cells:
            v = c.value
            if v is None:
                continue
            lengths.append(len(str(v)))
        if not lengths:
            continue
        ws.column_dimensions[col_letter].width = min(max_width, max(10, max(lengths) + 2))


def write_week_tabs(wb: Workbook, week_label: str, week_start: date, assign_grid, df_long: pd.DataFrame, df_gaps: pd.DataFrame):
    slots = timeslots()
    dates = [week_start + timedelta(days=i) for i in range(5)]

    # ---------- GRID (manager-friendly, slot view) ----------
    ws_grid = wb.create_sheet(f"{week_label}_Grid")
    ws_grid.append(["Time"] + [d.strftime("%a %d-%b") for d in dates])
    for cell in ws_grid[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")

    # Compact coverage summary string per slot/day
    for t in slots:
        row = [t.strftime("%H:%M")]
        for d in dates:
            assigns = assign_grid.get((d, t), [])
            # order for readability
            order = [
                ("FrontDesk", "SLGP"),
                ("FrontDesk", "JEN"),
                ("FrontDesk", "BGS"),
                ("Triage_Admin", "SLGP"),
                ("Triage_Admin", "JEN"),
                ("Email_Box", ""),
                ("Phones", ""),
            ]
            parts = []
            # index assigns
            for role, site in order:
                if role == "Phones":
                    phone_list = [a for a in assigns if a["Role"] == "Phones"]
                    if phone_list:
                        # show labels Phones_1 etc
                        parts.append("Phones: " + ", ".join(f'{a["Label"]}:{a["Name"]}' for a in phone_list))
                    continue
                if role == "Email_Box":
                    a = next((a for a in assigns if a["Role"] == "Email_Box"), None)
                    if a:
                        parts.append(f'Email:{a["Name"]}')
                    continue
                a = next((a for a in assigns if a["Role"] == role and a["Site"] == site), None)
                if a:
                    parts.append(f'{role}_{site}:{a["Name"]}')
            row.append("\n".join(parts))
        ws_grid.append(row)

    ws_grid.column_dimensions["A"].width = 8
    for col in range(2, 2 + len(dates)):
        ws_grid.column_dimensions[chr(64 + col)].width = 34
    for r in range(2, 2 + len(slots)):
        ws_grid.row_dimensions[r].height = 70
        for c in range(2, 2 + len(dates)):
            ws_grid.cell(r, c).alignment = Alignment(wrap_text=True, vertical="top")

    # ---------- LONGFORM (task totals + auditing) ----------
    ws_long = wb.create_sheet(f"{week_label}_LongForm")
    if df_long.empty:
        df_long = pd.DataFrame(columns=["Date","Day","Time","Name","HomeSite","Role","RoleBase","Site","Hours"])
    df_long_out = df_long.copy()
    df_long_out["Date"] = df_long_out["Date"].astype(str)

    ws_long.append(list(df_long_out.columns))
    for cell in ws_long[1]:
        cell.font = Font(bold=True)

    for r in dataframe_to_rows(df_long_out, index=False, header=False):
        ws_long.append(r)

    # colour RoleBase column
    rolebase_col_idx = list(df_long_out.columns).index("RoleBase") + 1
    for row_idx in range(2, ws_long.max_row + 1):
        rb = ws_long.cell(row_idx, rolebase_col_idx).value
        ws_long.cell(row_idx, rolebase_col_idx).fill = fill_for_role(str(rb))

    autosize_columns(ws_long)

    # ---------- COVERAGE GAPS ----------
    ws_gaps = wb.create_sheet(f"{week_label}_CoverageGaps")
    if df_gaps is None or df_gaps.empty:
        df_gaps = pd.DataFrame(columns=["Date","Time","Role","Site","Issue"])
    df_gaps_out = df_gaps.copy()
    df_gaps_out["Date"] = df_gaps_out["Date"].astype(str)

    ws_gaps.append(list(df_gaps_out.columns))
    for cell in ws_gaps[1]:
        cell.font = Font(bold=True)

    for r in dataframe_to_rows(df_gaps_out, index=False, header=False):
        ws_gaps.append(r)

    autosize_columns(ws_gaps)

    # ---------- TOTALS (daily + weekly staff totals, plus site totals) ----------
    ws_tot = wb.create_sheet(f"{week_label}_Totals")

    df = df_long.copy()
    if df.empty:
        df = pd.DataFrame(columns=["Date","Name","HomeSite","RoleBase","Hours"])

    # weekly totals per staff by RoleBase
    p_week = (
        df.groupby(["Name", "RoleBase"])["Hours"].sum()
        .reset_index()
        .pivot(index="Name", columns="RoleBase", values="Hours")
        .fillna(0.0)
    )
    p_week["WeeklyTotal"] = p_week.sum(axis=1)
    p_week = p_week.reset_index()

    # daily totals per staff by RoleBase
    p_day = (
        df.groupby(["Date", "Name", "RoleBase"])["Hours"].sum()
        .reset_index()
        .sort_values(["Date", "Name", "RoleBase"])
    )

    # weekly totals per site (by staff HomeSite)
    p_site = (
        df.groupby(["HomeSite"])["Hours"].sum()
        .reset_index()
        .rename(columns={"Hours": "WeeklyHoursTotal"})
        .sort_values("HomeSite")
    )

    ws_tot.append(["Weekly totals per staff by task (hours)"])
    ws_tot["A1"].font = Font(bold=True, size=12)
    ws_tot.append([])

    # write weekly pivot
    for r in dataframe_to_rows(p_week, index=False, header=True):
        ws_tot.append(r)

    # header styling + column fills
    header_row = 3
    headers = [c.value for c in ws_tot[header_row]]
    for cell in ws_tot[header_row]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center")

    # Apply role colour fills to role columns
    for j, h in enumerate(headers, start=1):
        if not h or h == "Name":
            continue
        base = str(h)
        if base in ROLE_COLORS:
            fill = fill_for_role(base)
            for i in range(header_row, header_row + len(p_week) + 1):
                ws_tot.cell(i, j).fill = fill

    ws_tot.append([])
    ws_tot.append(["Daily totals per staff by task (hours)"])
    ws_tot[f"A{ws_tot.max_row}"].font = Font(bold=True, size=12)
    ws_tot.append([])

    for r in dataframe_to_rows(p_day, index=False, header=True):
        ws_tot.append(r)

    ws_tot.append([])
    ws_tot.append(["Weekly hours totals per site (all work incl. filler + breaks)"])
    ws_tot[f"A{ws_tot.max_row}"].font = Font(bold=True, size=12)
    ws_tot.append([])

    for r in dataframe_to_rows(p_site, index=False, header=True):
        ws_tot.append(r)

    autosize_columns(ws_tot)


def build_excel_for_period(staff_df, hours_df, hols, start_monday: date, weeks: int, phones_max: int):
    wb = Workbook()
    wb.remove(wb.active)

    fairness_state = None

    for w in range(weeks):
        ws = start_monday + timedelta(days=7 * w)
        assign_grid, df_long, df_gaps, fairness_state = rota_generate_one_week(
            staff_df, hours_df, hols, week_start=ws, phones_max=phones_max, fairness_state=fairness_state
        )
        label = f"Week{w+1}_{ws.strftime('%d%b')}"
        write_week_tabs(wb, label, ws, assign_grid, df_long, df_gaps)

    return wb


# =========================================================
# Streamlit UI
# =========================================================
st.set_page_config(page_title="Rota Generator", layout="wide")
require_password()

st.title("Rota Generator (Excel export)")

with st.expander("Upload your completed template", expanded=True):
    uploaded = st.file_uploader("Upload rota template (.xlsx)", type=["xlsx"])

col1, col2, col3 = st.columns(3)
with col1:
    start_date = st.date_input("Week commencing (Monday)", value=date.today())
with col2:
    mode = st.selectbox("Run period", ["1 week", "4 weeks (month)", "Custom # weeks"])
with col3:
    phones_max = st.selectbox("Max phones if spare staff", [2, 3], index=1)

if mode == "Custom # weeks":
    weeks = st.number_input("Number of weeks", min_value=1, max_value=12, value=2, step=1)
elif mode == "4 weeks (month)":
    weeks = 4
else:
    weeks = 1

start_monday = ensure_monday(start_date)
st.caption(f"Will generate from Monday {start_monday.isoformat()} for {int(weeks)} week(s).")

if uploaded:
    try:
        staff_df, hours_df, hols, params, found_sheets = read_template(uploaded.getvalue())
        st.success(f"Template loaded. Sheets found: {found_sheets}")

        with st.expander("Preview: Staff + Skills", expanded=False):
            st.dataframe(staff_df, use_container_width=True, height=280)

        with st.expander("Preview: Working hours", expanded=False):
            st.dataframe(hours_df, use_container_width=True, height=280)

        with st.expander("Preview: Holidays (excluded)", expanded=False):
            if hols:
                st.write(pd.DataFrame(hols, columns=["Name", "StartDate", "EndDate"]))
            else:
                st.write("No holidays found / configured.")

        if st.button("Generate rota and export Excel"):
            wb = build_excel_for_period(
                staff_df=staff_df,
                hours_df=hours_df,
                hols=hols,
                start_monday=start_monday,
                weeks=int(weeks),
                phones_max=int(phones_max),
            )

            bio = io.BytesIO()
            wb.save(bio)
            bio.seek(0)

            out_name = f"rota_{start_monday.isoformat()}_{int(weeks)}w.xlsx"
            st.download_button(
                "📊 Download Excel rota",
                data=bio.getvalue(),
                file_name=out_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            st.success("Done — download the Excel file above.")

    except Exception as e:
        st.error("Could not process the template.")
        st.exception(e)
else:
    st.info("Upload your completed template to continue.")
