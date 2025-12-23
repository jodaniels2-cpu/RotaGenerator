import io
import re
from datetime import datetime, date, time, timedelta
import pandas as pd
import streamlit as st

from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils.dataframe import dataframe_to_rows


# -----------------------------
# Password protection (Streamlit Secrets)
# -----------------------------
def require_password():
    # Set in Streamlit Cloud -> App -> Settings -> Secrets:
    # APP_PASSWORD="yourpassword"
    pw = st.secrets.get("APP_PASSWORD", None)
    if not pw:
        st.warning("Password is not configured. Add APP_PASSWORD in Streamlit Secrets to enable protection.")
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


# -----------------------------
# Helpers: robust sheet + column detection
# -----------------------------
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
        raise KeyError(f"Could not find required column among: {candidates}. Available: {list(df.columns)}")
    return None

def to_time(x):
    if pd.isna(x) or x == "":
        return None
    if isinstance(x, time):
        return x
    if isinstance(x, datetime):
        return x.time()
    if isinstance(x, (float, int)):
        # Excel time as fraction of day
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


# -----------------------------
# Parse template (robust)
# -----------------------------
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
        raise ValueError(f"Could not find Parameters/Params sheet. Found: {xls.sheet_names}")

    staff_df = pd.read_excel(xls, sheet_name=staff_sheet)
    hours_df = pd.read_excel(xls, sheet_name=hours_sheet)
    hols_df = pd.read_excel(xls, sheet_name=hols_sheet) if hols_sheet else pd.DataFrame()
    params_df = pd.read_excel(xls, sheet_name=params_sheet)

    # --- staff cols ---
    name_c = pick_col(staff_df, ["Name", "StaffName"])
    home_c = pick_col(staff_df, ["HomeSite", "Site", "BaseSite"], required=False)

    staff_df = staff_df.copy()
    staff_df["Name"] = staff_df[name_c].astype(str).str.strip()
    staff_df["HomeSite"] = staff_df[home_c].astype(str).str.strip().fillna("") if home_c else ""

    # skills: any Y/N columns become skill flags
    skill_cols = [c for c in staff_df.columns if normalize(c) not in [normalize(name_c), normalize(home_c or "")]]
    # keep only meaningful skill columns (avoid blanks)
    skill_cols = [c for c in skill_cols if staff_df[c].notna().any()]

    def yn(v):
        if pd.isna(v): return False
        s = str(v).strip().lower()
        return s in ["y", "yes", "true", "1"]

    for c in skill_cols:
        staff_df[c] = staff_df[c].apply(yn)

    # detect Carol Church column (optional)
    carol_c = None
    for c in staff_df.columns:
        if normalize(c) in ["iscarolchurch", "carolchurch", "carol"]:
            carol_c = c
            break
    if carol_c:
        staff_df["IsCarolChurch"] = staff_df[carol_c].apply(bool)
    else:
        staff_df["IsCarolChurch"] = staff_df["Name"].str.lower().eq("carol church")

    # --- working hours ---
    hn = normalize
    hours_df = hours_df.copy()
    hours_name_c = pick_col(hours_df, ["Name", "StaffName"])
    hours_df["Name"] = hours_df[hours_name_c].astype(str).str.strip()

    # day columns: MonStart MonEnd ...
    days = ["Mon", "Tue", "Wed", "Thu", "Fri"]
    for d in days:
        sc = pick_col(hours_df, [f"{d}Start", f"{d} Start", f"{d}_Start"], required=False)
        ec = pick_col(hours_df, [f"{d}End", f"{d} End", f"{d}_End"], required=False)
        hours_df[f"{d}Start"] = hours_df[sc].apply(to_time) if sc else None
        hours_df[f"{d}End"] = hours_df[ec].apply(to_time) if ec else None

    # --- holidays ---
    hols = []
    if not hols_df.empty:
        hname = pick_col(hols_df, ["Name", "StaffName"], required=False)
        hs = pick_col(hols_df, ["StartDate", "Start"], required=False)
        he = pick_col(hols_df, ["EndDate", "End"], required=False)
        if hname and hs and he:
            for _, r in hols_df.iterrows():
                hols.append((str(r[hname]).strip(), to_date(r[hs]), to_date(r[he])))

    # --- parameters ---
    # accept many formats: Rule/Value, Parameter/Value, Key/Value
    rule_c = pick_col(params_df, ["Rule", "Parameter", "Key", "Setting"], required=False)
    val_c = pick_col(params_df, ["Value", "Val"], required=False)
    params = {}
    if rule_c and val_c:
        for _, r in params_df.iterrows():
            k = str(r[rule_c]).strip()
            if k and k.lower() != "nan":
                params[k] = r[val_c]
    else:
        # fallback: first two columns
        if params_df.shape[1] >= 2:
            c1, c2 = params_df.columns[:2]
            for _, r in params_df.iterrows():
                k = str(r[c1]).strip()
                if k and k.lower() != "nan":
                    params[k] = r[c2]

    return staff_df, hours_df, hols, params, xls.sheet_names


# -----------------------------
# Core rota model
# -----------------------------
SITES = ["SLGP", "JEN", "BGS"]

# Role keys used in output + color map
ROLE_COLORS = {
    "FrontDesk_SLGP": "FFF2CC",
    "FrontDesk_JEN":  "FFF2CC",
    "FrontDesk_BGS":  "FFF2CC",
    "Triage_Admin_SLGP": "D9EAD3",
    "Triage_Admin_JEN":  "D9EAD3",
    "Email_Box": "CFE2F3",
    "Phones": "C9DAF8",
    "Bookings": "FCE5CD",
    "EMIS": "EAD1DC",
    "Docman_PSA": "D0E0E3",
    "Docman_Awaiting": "D0E0E3",
    "Break": "E6E6E6",
    "Unassigned": "FFFFFF",
}

MANDATORY = [
    # all day front desks
    ("FrontDesk_SLGP", time(8,0), time(18,30), 1),
    ("FrontDesk_JEN",  time(8,0), time(18,30), 1),
    ("FrontDesk_BGS",  time(8,0), time(18,30), 1),

    # triage at JEN + SLGP until 16:00
    ("Triage_Admin_SLGP", time(8,0), time(16,0), 1),
    ("Triage_Admin_JEN",  time(8,0), time(16,0), 1),

    # email always by JEN/BGS
    ("Email_Box", time(8,0), time(18,30), 1),

    # phones 2–3 by JEN/BGS (min 2 enforced; we allow 3 if spare)
    ("Phones", time(8,0), time(18,30), 2),
]

FILLER = [
    ("Bookings", time(8,0), time(18,30), None),         # SLGP majority
    ("EMIS", time(8,0), time(18,30), None),             # JEN/BGS
    ("Docman_PSA", time(8,0), time(18,30), None),       # JEN/BGS
    ("Docman_Awaiting", time(8,0), time(18,30), None),  # JEN/BGS
]

def daterange(d0, d1):
    d = d0
    while d <= d1:
        yield d
        d += timedelta(days=1)

def is_on_holiday(name, d, hols):
    for n, s, e in hols:
        if n.strip().lower() == name.strip().lower() and s and e and s <= d <= e:
            return True
    return False

def timeslots(day_start=time(8,0), day_end=time(18,30), step_min=30):
    start_dt = datetime(2000,1,1, day_start.hour, day_start.minute)
    end_dt = datetime(2000,1,1, day_end.hour, day_end.minute)
    cur = start_dt
    out = []
    while cur < end_dt:
        out.append(cur.time())
        cur += timedelta(minutes=step_min)
    return out

def t_in_range(t, a, b):
    return (t >= a) and (t < b)

def hours_between(t1, t2):
    dt1 = datetime(2000,1,1,t1.hour,t1.minute)
    dt2 = datetime(2000,1,1,t2.hour,t2.minute)
    return (dt2 - dt1).total_seconds() / 3600

def skill_allowed(staff_row, role):
    # map role -> skill column name possibilities
    role_map = {
        "FrontDesk": ["FrontDesk", "Front Desk", "Reception"],
        "Triage": ["Triage", "AdminTriage", "Admin Triage"],
        "Email_Box": ["Email", "EmailBox", "Emails", "Email Box"],
        "Phones": ["Phones", "Phone"],
        "Bookings": ["Bookings", "Booking"],
        "EMIS": ["EMIS"],
        "Docman_PSA": ["DocmanPSA", "Docman PSA", "PSA"],
        "Docman_Awaiting": ["DocmanAwaiting", "AwaitingResponse", "Awaiting Response"],
    }

    if role.startswith("FrontDesk"):
        keys = role_map["FrontDesk"]
    elif role.startswith("Triage_Admin"):
        keys = role_map["Triage"]
    else:
        keys = role_map.get(role, [role])

    cols = {normalize(c): c for c in staff_row.index}
    for k in keys:
        nk = normalize(k)
        if nk in cols:
            return bool(staff_row[cols[nk]])
    # if skill column missing, default to False (safer)
    return False

def site_restriction_ok(staff_row, role):
    home = str(staff_row.get("HomeSite","")).strip().upper()
    # Email + Phones only JEN/BGS
    if role in ["Email_Box", "Phones", "EMIS", "Docman_PSA", "Docman_Awaiting"]:
        return home in ["JEN", "BGS"]
    # Bookings mainly SLGP (allow SLGP only here)
    if role == "Bookings":
        return home == "SLGP"
    # site-specific roles
    if role.endswith("_SLGP"):
        return home == "SLGP"
    if role.endswith("_JEN"):
        return home == "JEN"
    if role.endswith("_BGS"):
        return home == "BGS"
    return True

def build_availability(staff_df, hours_df, hols, week_start: date):
    # returns dict day-> list of available staff names
    hmap = {r["Name"]: r for _, r in hours_df.iterrows()}

    days = ["Mon","Tue","Wed","Thu","Fri"]
    out = {}
    for i, dname in enumerate(days):
        d = week_start + timedelta(days=i)
        available = []
        for _, s in staff_df.iterrows():
            name = s["Name"]
            if is_on_holiday(name, d, hols):
                continue
            hr = hmap.get(name)
            if hr is None:
                continue
            stt = hr.get(f"{dname}Start")
            end = hr.get(f"{dname}End")
            if stt and end:
                available.append(name)
        out[d] = available
    return out, hmap

def staff_work_window(hmap_row, dname):
    stt = hmap_row.get(f"{dname}Start")
    end = hmap_row.get(f"{dname}End")
    return stt, end

def rota_generate_one_week(
    staff_df, hours_df, hols, week_start: date,
    phones_max=3,
    break_window=(time(11,30), time(14,0)),
    max_frontdesk_block_hours=2.5,
    max_triage_block_hours=3.0,
    fairness_state=None,
):
    """
    fairness_state: dict tracking cumulative penalties/hours across weeks, optional.
    """
    if fairness_state is None:
        fairness_state = {
            "role_hours": {},      # (name, role)->hours
            "frontdesk_hours": {}, # name->hours
            "triage_hours": {},    # name->hours
        }

    slots = timeslots()
    days = ["Mon","Tue","Wed","Thu","Fri"]

    availability, hmap = build_availability(staff_df, hours_df, hols, week_start)
    staff_by_name = {r["Name"]: r for _, r in staff_df.iterrows()}

    # assignments: dict[(date, time)] -> dict role -> staffname
    assign = {}
    gaps = []

    # track consecutive blocks per staff per role-group (front desk / triage)
    frontdesk_block = {} # (date, staff)->consecutive hours
    triage_block = {}    # (date, staff)->consecutive hours

    # track total assigned hours per day per staff and whether break placed
    day_hours = {}       # (date, staff)->hours
    break_done = {}      # (date, staff)->bool

    def can_work(name, d, t):
        hr = hmap.get(name)
        if hr is None:
            return False
        dname = days[(d - week_start).days]
        stt, end = staff_work_window(hr, dname)
        return stt and end and (t >= stt) and (t < end)

    def score_candidate(name, role):
        # lower is better
        key = (name, role)
        base = fairness_state["role_hours"].get(key, 0.0)
        # penalize overuse of front desk & triage in general
        if role.startswith("FrontDesk"):
            base += 2.0 * fairness_state["frontdesk_hours"].get(name, 0.0)
        if role.startswith("Triage_Admin"):
            base += 1.5 * fairness_state["triage_hours"].get(name, 0.0)
        return base

    def pick_staff(d, t, role, already_used_set):
        # enforce Carol rule: if Carol working, always front desk at her site
        if role.startswith("FrontDesk"):
            for _, sr in staff_df.iterrows():
                if sr.get("IsCarolChurch", False) and can_work(sr["Name"], d, t):
                    # Carol must be on front desk "when she is in"
                    # Only if the role matches her home site
                    home = str(sr.get("HomeSite","")).upper()
                    if role.endswith(home):
                        return sr["Name"]

        candidates = []
        for name in availability.get(d, []):
            if name in already_used_set:
                continue
            if not can_work(name, d, t):
                continue
            sr = staff_by_name[name]
            if not site_restriction_ok(sr, role):
                continue
            # skills
            if role.startswith("FrontDesk"):
                if not skill_allowed(sr, "FrontDesk"):
                    continue
                # block limit
                cur = frontdesk_block.get((d, name), 0.0)
                if cur + 0.5 > max_frontdesk_block_hours:
                    continue
            elif role.startswith("Triage_Admin"):
                if not skill_allowed(sr, "Triage"):
                    continue
                cur = triage_block.get((d, name), 0.0)
                if cur + 0.5 > max_triage_block_hours:
                    continue
            else:
                if not skill_allowed(sr, role):
                    continue

            candidates.append(name)

        if not candidates:
            return None

        candidates.sort(key=lambda n: (score_candidate(n, role), day_hours.get((d,n), 0.0)))
        return candidates[0]

    # schedule breaks: if staff has >6 hours that day, force one 30-min slot in window
    def maybe_force_breaks(d, t):
        bw0, bw1 = break_window
        if not t_in_range(t, bw0, bw1):
            return
        for name in availability.get(d, []):
            if not can_work(name, d, t):
                continue
            if break_done.get((d, name), False):
                continue
            # only force once we already know they are likely to exceed 6 hours:
            # heuristic: if their shift length > 6 hours, require break
            hr = hmap[name]
            dname = days[(d - week_start).days]
            stt, end = staff_work_window(hr, dname)
            if not stt or not end:
                continue
            shift_len = hours_between(stt, end)
            if shift_len <= 6:
                continue

            # if person is currently assigned to something in this slot, skip forcing
            slot_roles = assign.get((d, t), {})
            if name in slot_roles.values():
                continue

            # mark break
            assign.setdefault((d, t), {})
            # allow multiple breaks in same slot for different staff using longform output later
            # For grid, we'll record breaks in longform and totals; grid remains role-based coverage.
            break_done[(d, name)] = True

    # Fill mandatory roles each slot
    for d in sorted(availability.keys()):
        for t in slots:
            maybe_force_breaks(d, t)
            used = set()
            slot_roles = assign.setdefault((d, t), {})

            # mandatory roles
            for role, r0, r1, need in MANDATORY:
                if not t_in_range(t, r0, r1):
                    continue

                # Phones: we enforce min=2; later optionally add 3rd if spare
                if role == "Phones":
                    for k in range(need):
                        pick = pick_staff(d, t, role, used)
                        if pick:
                            slot_roles[f"Phones_{k+1}"] = pick
                            used.add(pick)
                            day_hours[(d, pick)] = day_hours.get((d, pick), 0.0) + 0.5
                        else:
                            gaps.append((d, t, "Phones", "No available JEN/BGS phone-trained staff"))
                    # try add a 3rd phone if spare capacity and phones_max==3
                    if phones_max >= 3:
                        pick = pick_staff(d, t, role, used)
                        if pick:
                            slot_roles["Phones_3"] = pick
                            used.add(pick)
                            day_hours[(d, pick)] = day_hours.get((d, pick), 0.0) + 0.5
                    continue

                # normal mandatory single coverage
                pick = pick_staff(d, t, role, used)
                if pick:
                    slot_roles[role] = pick
                    used.add(pick)
                    day_hours[(d, pick)] = day_hours.get((d, pick), 0.0) + 0.5

                    # update blocks
                    if role.startswith("FrontDesk"):
                        frontdesk_block[(d, pick)] = frontdesk_block.get((d, pick), 0.0) + 0.5
                    else:
                        # reset if not on front desk this slot
                        for n in list(frontdesk_block.keys()):
                            dd, nm = n
                            if dd == d and nm == pick:
                                pass
                    if role.startswith("Triage_Admin"):
                        triage_block[(d, pick)] = triage_block.get((d, pick), 0.0) + 0.5
                else:
                    gaps.append((d, t, role, "No suitable staff available"))

            # reset blocks for staff not currently on that role (simple approximation)
            # (prevents blocks carrying over incorrectly)
            fd_names = {v for k, v in slot_roles.items() if k.startswith("FrontDesk")}
            tr_names = {v for k, v in slot_roles.items() if k.startswith("Triage_Admin")}
            for name in availability.get(d, []):
                if name not in fd_names:
                    frontdesk_block[(d, name)] = 0.0
                if name not in tr_names:
                    triage_block[(d, name)] = 0.0

    # Filler tasks: assign remaining working time that isn’t already used
    # We do this in longform output rather than overloading the grid with unlimited roles.
    # We'll calculate "unassigned capacity" per staff per day and allocate to tasks by policy.
    filler_alloc = []  # rows: date, name, role, hours
    for i, dname in enumerate(days):
        d = week_start + timedelta(days=i)
        for name in availability.get(d, []):
            hr = hmap[name]
            stt, end = staff_work_window(hr, dname)
            if not stt or not end:
                continue
            shift_len = hours_between(stt, end)
            already = day_hours.get((d, name), 0.0)
            # subtract break (0.5h) if needed and break was scheduled
            if shift_len > 6:
                if break_done.get((d, name), False):
                    already += 0.5
                else:
                    # if no break could be placed, record gap
                    gaps.append((d, None, "Break", f"{name} shift > 6h but no break placed in 11:30–14:00"))
            remaining = max(0.0, shift_len - already)

            if remaining <= 0:
                continue

            # task policy:
            # SLGP -> Bookings majority; JEN/BGS -> EMIS + Docman
            home = str(staff_by_name[name].get("HomeSite","")).upper()
            if home == "SLGP" and skill_allowed(staff_by_name[name], "Bookings"):
                filler_alloc.append((d, name, "Bookings", remaining))
            elif home in ["JEN", "BGS"]:
                # split between EMIS and Docman if trained
                parts = []
                if skill_allowed(staff_by_name[name], "EMIS"):
                    parts.append("EMIS")
                if skill_allowed(staff_by_name[name], "Docman_PSA"):
                    parts.append("Docman_PSA")
                if skill_allowed(staff_by_name[name], "Docman_Awaiting"):
                    parts.append("Docman_Awaiting")

                if not parts:
                    filler_alloc.append((d, name, "Unassigned", remaining))
                else:
                    per = remaining / len(parts)
                    for r in parts:
                        filler_alloc.append((d, name, r, per))
            else:
                filler_alloc.append((d, name, "Unassigned", remaining))

    # Update fairness_state (so next week balances better)
    # We count the mandatory grid roles by slot + filler allocations
    for (d, t), slot_roles in assign.items():
        for role_key, name in slot_roles.items():
            role = role_key if not role_key.startswith("Phones_") else "Phones"
            fairness_state["role_hours"][(name, role)] = fairness_state["role_hours"].get((name, role), 0.0) + 0.5
            if role.startswith("FrontDesk"):
                fairness_state["frontdesk_hours"][name] = fairness_state["frontdesk_hours"].get(name, 0.0) + 0.5
            if role.startswith("Triage_Admin"):
                fairness_state["triage_hours"][name] = fairness_state["triage_hours"].get(name, 0.0) + 0.5

    for d, name, role, hrs in filler_alloc:
        fairness_state["role_hours"][(name, role)] = fairness_state["role_hours"].get((name, role), 0.0) + float(hrs)

    return assign, filler_alloc, gaps, fairness_state


# -----------------------------
# Excel Output
# -----------------------------
def fill_for_role(role):
    col = ROLE_COLORS.get(role, "FFFFFF")
    return PatternFill("solid", fgColor=col)

def write_week_to_workbook(wb: Workbook, title: str, week_start: date, assign, filler_alloc, gaps):
    ws_grid = wb.create_sheet(f"{title}_Grid")
    ws_gap = wb.create_sheet(f"{title}_CoverageGaps")
    ws_tot = wb.create_sheet(f"{title}_Totals")

    slots = timeslots()
    days = ["Mon","Tue","Wed","Thu","Fri"]
    dates = [week_start + timedelta(days=i) for i in range(5)]

    # --- GRID ---
    # Columns: Time, then each day
    ws_grid.append(["Time"] + [d.strftime("%a %d-%b") for d in dates])
    for cell in ws_grid[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")

    # For each slot, we put a compact string summarizing coverage
    # (Front desks, triage, email, phones)
    for t in slots:
        row = [t.strftime("%H:%M")]
        for d in dates:
            slot_roles = assign.get((d, t), {})
            # build readable summary
            parts = []
            for k in ["FrontDesk_SLGP","FrontDesk_JEN","FrontDesk_BGS","Triage_Admin_SLGP","Triage_Admin_JEN","Email_Box"]:
                if k in slot_roles:
                    parts.append(f"{k}:{slot_roles[k]}")
            phones = [slot_roles.get("Phones_1"), slot_roles.get("Phones_2"), slot_roles.get("Phones_3")]
            phones = [p for p in phones if p]
            if phones:
                parts.append("Phones:" + ", ".join(phones))

            txt = "\n".join(parts) if parts else ""
            row.append(txt)
        ws_grid.append(row)

    # style grid
    ws_grid.column_dimensions["A"].width = 8
    for col in range(2, 7):
        ws_grid.column_dimensions[chr(64+col)].width = 32
    for r in range(2, 2+len(slots)):
        ws_grid.row_dimensions[r].height = 60
        for c in range(2, 7):
            ws_grid.cell(r, c).alignment = Alignment(wrap_text=True, vertical="top")

    # --- COVERAGE GAPS ---
    ws_gap.append(["Date", "Time", "Role", "Issue"])
    ws_gap[1][0].font = ws_gap[1][1].font = ws_gap[1][2].font = ws_gap[1][3].font = Font(bold=True)

    for d, t, role, issue in gaps:
        t_str = "" if t is None else t.strftime("%H:%M")
        ws_gap.append([d.isoformat(), t_str, role, issue])

    # --- TOTALS ---
    # Daily & weekly totals per staff by task + weekly hours per site
    # Build longform from mandatory grid + filler_alloc
    rows = []
    for (d, t), slot_roles in assign.items():
        for role_key, name in slot_roles.items():
            role = "Phones" if role_key.startswith("Phones_") else role_key
            rows.append([d, name, role, 0.5])
    for d, name, role, hrs in filler_alloc:
        rows.append([d, name, role, float(hrs)])

    df = pd.DataFrame(rows, columns=["Date","Name","Role","Hours"])
    if df.empty:
        df = pd.DataFrame(columns=["Date","Name","Role","Hours"])

    # Staff weekly totals
    pivot_week = (df.groupby(["Name","Role"])["Hours"].sum()
                    .reset_index()
                    .pivot(index="Name", columns="Role", values="Hours")
                    .fillna(0.0))
    pivot_week["WeeklyTotal"] = pivot_week.sum(axis=1)
    pivot_week = pivot_week.reset_index()

    # Daily totals
    pivot_day = (df.groupby(["Date","Name","Role"])["Hours"].sum()
                  .reset_index())

    # Weekly hours totals per site (approx: infer from role name suffix or staff home site not included here,
    # so we estimate by role: FrontDesk_SITE and Triage_SITE as site hours; other tasks as staff home site unknown in this tab)
    site_hours = []
    for role in ["FrontDesk_SLGP","FrontDesk_JEN","FrontDesk_BGS","Triage_Admin_SLGP","Triage_Admin_JEN"]:
        hrs = df.loc[df["Role"]==role, "Hours"].sum()
        if role.endswith("_SLGP"): site="SLGP"
        elif role.endswith("_JEN"): site="JEN"
        else: site="BGS"
        site_hours.append([site, role, float(hrs)])
    site_df = pd.DataFrame(site_hours, columns=["Site","Role","Hours"])
    site_sum = site_df.groupby("Site")["Hours"].sum().reset_index().rename(columns={"Hours":"WeeklyHoursTotal"})

    # Write totals sheet
    ws_tot.append(["Weekly totals per staff by task (hours)"])
    ws_tot["A1"].font = Font(bold=True, size=12)

    start_row = 3
    for r in dataframe_to_rows(pivot_week, index=False, header=True):
        ws_tot.append(r)

    # format header row
    header_row = start_row
    for cell in ws_tot[header_row]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")

    # colour-code columns by role
    # header row is at row=3
    headers = [c.value for c in ws_tot[header_row]]
    for j, h in enumerate(headers, start=1):
        if not h or h == "Name": 
            continue
        fill = fill_for_role(h)
        for i in range(header_row, header_row + len(pivot_week) + 1):
            ws_tot.cell(i, j).fill = fill

    # daily totals block
    ws_tot.append([])
    ws_tot.append(["Daily totals per staff by task"])
    ws_tot[f"A{ws_tot.max_row}"].font = Font(bold=True, size=12)

    for r in dataframe_to_rows(pivot_day, index=False, header=True):
        ws_tot.append(r)

    # site summary block
    ws_tot.append([])
    ws_tot.append(["Weekly hours totals per site (coverage roles)"])
    ws_tot[f"A{ws_tot.max_row}"].font = Font(bold=True, size=12)

    for r in dataframe_to_rows(site_sum, index=False, header=True):
        ws_tot.append(r)

    return wb


def build_excel_for_period(staff_df, hours_df, hols, start_monday: date, weeks: int, phones_max=3):
    wb = Workbook()
    # remove default sheet
    wb.remove(wb.active)

    fairness_state = None
    all_gap_rows = []
    for w in range(weeks):
        ws = start_monday + timedelta(days=7*w)
        assign, filler_alloc, gaps, fairness_state = rota_generate_one_week(
            staff_df, hours_df, hols,
            week_start=ws,
            phones_max=phones_max,
            fairness_state=fairness_state
        )
        title = f"Week{w+1}_{ws.strftime('%d%b')}"
        write_week_to_workbook(wb, title, ws, assign, filler_alloc, gaps)
        all_gap_rows.extend([(ws,)+g for g in gaps])

    return wb


# -----------------------------
# Streamlit UI
# -----------------------------
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

def ensure_monday(d: date) -> date:
    return d - timedelta(days=d.weekday())

if uploaded:
    try:
        staff_df, hours_df, hols, params, found_sheets = read_template(uploaded.getvalue())
        st.success(f"Template loaded. Sheets found: {found_sheets}")

        # show quick preview
        with st.expander("Preview: Staff + Skills"):
            st.dataframe(staff_df.head(50), use_container_width=True)
        with st.expander("Preview: Working Hours"):
            st.dataframe(hours_df.head(50), use_container_width=True)

        start_monday = ensure_monday(start_date)
        st.info(f"Generating from Monday: {start_monday.isoformat()} for {weeks} week(s).")

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

            out_name = f"rota_{start_monday.isoformat()}_{weeks}w.xlsx"
            st.download_button(
                "📊 Download Excel rota",
                data=bio.getvalue(),
                file_name=out_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            st.success("Done — download the Excel file above.")

    except Exception as e:
        st.error("Could not process the template.")
        st.exception(e)
else:
    st.warning("Upload your completed template to continue.")
