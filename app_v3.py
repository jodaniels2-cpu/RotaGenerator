import io
import re
from datetime import datetime, date, time, timedelta
import pandas as pd
import streamlit as st

from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils.dataframe import dataframe_to_rows


# =========================================================
# Password protection (Streamlit Secrets in TOML)
# =========================================================
def require_password():
    """
    Set in Streamlit Secrets (TOML):

    # .streamlit/secrets.toml
    APP_PASSWORD = "your-strong-password"

    On Streamlit Cloud:
      App -> Settings -> Secrets -> paste TOML above
    """
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


# =========================================================
# Helpers: robust sheet + column detection
# =========================================================
def normalize(s: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", str(s).strip().lower())

def find_sheet(xls: pd.ExcelFile, candidates):
    names = {normalize(n): n for n in xls.sheet_names}
    for c in candidates:
        key = normalize(c)
        if key in names:
            return names[key]
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
    for c in df.columns:
        nc = normalize(c)
        for cand in candidates:
            if normalize(cand) in nc:
                return c
    if required:
        raise KeyError(f"Missing required column. Need one of {candidates}. Have {list(df.columns)}")
    return None

def to_time(x):
    if pd.isna(x) or x == "":
        return None
    if isinstance(x, time):
        return x
    if isinstance(x, datetime):
        return x.time()
    if isinstance(x, (float, int)):
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

def h_between(t1: time, t2: time) -> float:
    dt1 = datetime(2000,1,1,t1.hour,t1.minute)
    dt2 = datetime(2000,1,1,t2.hour,t2.minute)
    return (dt2-dt1).total_seconds()/3600.0

def dt_of(d: date, t: time) -> datetime:
    return datetime(d.year, d.month, d.day, t.hour, t.minute)

def add_minutes(t: time, mins: int) -> time:
    dt = datetime(2000,1,1,t.hour,t.minute) + timedelta(minutes=mins)
    return dt.time()

def t_in_range(t: time, a: time, b: time) -> bool:
    return (t >= a) and (t < b)


# =========================================================
# Parse template (robust)
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

    # --- staff cols ---
    name_c = pick_col(staff_df, ["Name", "StaffName"])
    home_c = pick_col(staff_df, ["HomeSite", "Site", "BaseSite"], required=False)

    staff_df = staff_df.copy()
    staff_df["Name"] = staff_df[name_c].astype(str).str.strip()
    staff_df["HomeSite"] = staff_df[home_c].astype(str).str.strip().fillna("") if home_c else ""

    # skills: any Y/N columns become flags
    base_cols = {normalize(name_c)}
    if home_c:
        base_cols.add(normalize(home_c))
    skill_cols = [c for c in staff_df.columns if normalize(c) not in base_cols]
    skill_cols = [c for c in skill_cols if staff_df[c].notna().any()]

    def yn(v):
        if pd.isna(v): return False
        s = str(v).strip().lower()
        return s in ["y","yes","true","1"]

    for c in skill_cols:
        staff_df[c] = staff_df[c].apply(yn)

    # Carol marker
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
    hours_df = hours_df.copy()
    hours_name_c = pick_col(hours_df, ["Name", "StaffName"])
    hours_df["Name"] = hours_df[hours_name_c].astype(str).str.strip()

    days = ["Mon","Tue","Wed","Thu","Fri"]
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

    # --- parameters (optional) ---
    rule_c = pick_col(params_df, ["Rule", "Parameter", "Key", "Setting"], required=False)
    val_c = pick_col(params_df, ["Value", "Val"], required=False)
    params = {}
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
# Rota rules
# =========================================================
ROLE_COLORS = {
    "FrontDesk_SLGP": "FFF2CC",
    "FrontDesk_JEN":  "FFF2CC",
    "FrontDesk_BGS":  "FFF2CC",
    "Triage_Admin_SLGP": "D9EAD3",
    "Triage_Admin_JEN":  "D9EAD3",
    "Email_Box": "CFE2F3",
    "Phones": "C9DAF8",
    "Bookings": "FCE5CD",
    "Emis_Tasks": "EAD1DC",
    "Docman_Tasks": "D0E0E3",
    "Awaiting_PSA_Admin": "D0E0E3",
    "Break": "E6E6E6",
    "Unassigned": "FFFFFF",
}

DAY_START = time(8,0)
DAY_END   = time(18,30)

SLOT_MIN = 30
MIN_STINT_HOURS = 1.5
MIN_STINT_SLOTS = int(MIN_STINT_HOURS * 60 / SLOT_MIN)

BREAK_WINDOW = (time(12,0), time(14,0))
BREAK_LEN_SLOTS = 1
SHIFT_BREAK_THRESHOLD_HOURS = 6.0

MANDATORY_RULES = [
    ("FrontDesk_SLGP", DAY_START, DAY_END, 1),
    ("FrontDesk_JEN",  DAY_START, DAY_END, 1),
    ("FrontDesk_BGS",  DAY_START, DAY_END, 1),
    ("Triage_Admin_SLGP", DAY_START, time(16,0), 1),
    ("Triage_Admin_JEN",  DAY_START, time(16,0), 1),
    ("Email_Box", DAY_START, DAY_END, 1),
    ("Phones", DAY_START, DAY_END, 2),
    ("Bookings", DAY_START, DAY_END, 3),
]

def psa_site_for_day(d: date):
    wd = d.weekday()
    if wd in (0, 4):
        return "SLGP"
    if wd in (1, 3):
        return "JEN"
    return "BGS"

def awaiting_required(d: date, t: time) -> bool:
    return t_in_range(t, time(10,0), time(16,0))

def awaiting_optional(d: date, t: time) -> bool:
    return t_in_range(t, time(16,0), DAY_END)


# =========================================================
# Skills + restrictions
# =========================================================
def skill_allowed(staff_row, role_key: str) -> bool:
    role_map = {
        "FrontDesk": ["FrontDesk", "Front Desk", "Reception"],
        "Triage": ["Triage", "AdminTriage", "Admin Triage"],
        "Email_Box": ["Email", "EmailBox", "Emails", "Email Box"],
        "Phones": ["Phones", "Phone"],
        "Bookings": ["Bookings", "Booking"],
        "Emis_Tasks": ["EMIS", "Emis", "Emis Tasks"],
        "Docman_Tasks": ["Docman", "Docman Tasks", "Docman_PSA", "Docman PSA", "Docman_Awaiting", "Awaiting Response", "AwaitingResponse"],
        "Awaiting_PSA_Admin": ["PSA", "PSA Admin", "Awaiting Response", "AwaitingResponse", "Docman_PSA", "Docman PSA"],
    }

    if role_key.startswith("FrontDesk"):
        keys = role_map["FrontDesk"]
    elif role_key.startswith("Triage_Admin"):
        keys = role_map["Triage"]
    else:
        keys = role_map.get(role_key, [role_key])

    cols = {normalize(c): c for c in staff_row.index}
    for k in keys:
        nk = normalize(k)
        if nk in cols:
            return bool(staff_row[cols[nk]])
    return False

def site_restriction_ok(staff_row, role_key: str) -> bool:
    home = str(staff_row.get("HomeSite","")).strip().upper()
    if role_key in ["Email_Box", "Phones", "Emis_Tasks", "Docman_Tasks", "Awaiting_PSA_Admin"]:
        return home in ["JEN", "BGS"]
    if role_key == "Bookings":
        return home == "SLGP"
    if role_key.endswith("_SLGP"):
        return home == "SLGP"
    if role_key.endswith("_JEN"):
        return home == "JEN"
    if role_key.endswith("_BGS"):
        return home == "BGS"
    return True


# =========================================================
# Availability + holidays + breaks
# =========================================================
def is_on_holiday(name: str, d: date, hols):
    for n, s, e in hols:
        if n.strip().lower() == name.strip().lower() and s and e and s <= d <= e:
            return True
    return False

def build_availability(staff_df, hours_df, hols, week_start: date):
    hmap = {r["Name"]: r for _, r in hours_df.iterrows()}
    days = ["Mon","Tue","Wed","Thu","Fri"]
    avail = {}
    for i, dname in enumerate(days):
        d = week_start + timedelta(days=i)
        names = []
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
                names.append(nm)
        avail[d] = names
    return avail, hmap

def ensure_monday(d: date) -> date:
    return d - timedelta(days=d.weekday())

def timeslots(day_start=DAY_START, day_end=DAY_END, step_min=SLOT_MIN):
    cur = datetime(2000,1,1,day_start.hour,day_start.minute)
    end = datetime(2000,1,1,day_end.hour,day_end.minute)
    out = []
    while cur < end:
        out.append(cur.time())
        cur += timedelta(minutes=step_min)
    return out

def shift_window(hmap_row, dname):
    return hmap_row.get(f"{dname}Start"), hmap_row.get(f"{dname}End")

def pick_break_slot_near_midpoint(d: date, stt: time, end: time):
    if not stt or not end:
        return None
    if h_between(stt, end) <= SHIFT_BREAK_THRESHOLD_HOURS:
        return None
    midpoint = dt_of(d, stt) + (dt_of(d, end) - dt_of(d, stt)) / 2
    candidates = []
    for t in [time(12,0), time(12,30), time(13,0), time(13,30)]:
        if t >= stt and add_minutes(t, 30) <= end and t_in_range(t, BREAK_WINDOW[0], BREAK_WINDOW[1]):
            dist = abs((dt_of(d, t) - midpoint).total_seconds())
            candidates.append((dist, t))
    if not candidates:
        return None
    candidates.sort(key=lambda x: x[0])
    return candidates[0][1]


# =========================================================
# Scheduler with min 1.5h stints
# =========================================================
def rota_generate_week(staff_df, hours_df, hols, week_start: date):
    slots = timeslots()
    days = ["Mon","Tue","Wed","Thu","Fri"]
    dates = [week_start + timedelta(days=i) for i in range(5)]

    avail, hmap = build_availability(staff_df, hours_df, hols, week_start)
    staff_by_name = {r["Name"]: r for _, r in staff_df.iterrows()}

    breaks = {}
    for d in dates:
        dname = days[(d-week_start).days]
        for name in avail.get(d, []):
            hr = hmap[name]
            stt, end = shift_window(hr, dname)
            b = pick_break_slot_near_midpoint(d, stt, end)
            if b:
                breaks.setdefault((d, b), set()).add(name)

    last_role_state = {}  # (d, name)-> (role, remaining_slots_lock)
    cont_fd = {}
    cont_triage = {}
    role_hours = {}

    assign = {}  # (d,t)-> list[(role,name)]
    gaps = []

    def is_working(name, d, t):
        dname = days[(d-week_start).days]
        hr = hmap.get(name)
        if hr is None:
            return False
        stt, end = shift_window(hr, dname)
        return bool(stt and end and (t >= stt) and (t < end))

    def is_on_break(name, d, t):
        return name in breaks.get((d, t), set())

    def already_used(name, d, t):
        return any(nm == name for _, nm in assign.get((d, t), []))

    def role_site_ok(role, d, staff_row):
        if role == "Awaiting_PSA_Admin":
            target = psa_site_for_day(d)
            return str(staff_row.get("HomeSite","")).strip().upper() == target
        return True

    def score(name, role, d):
        base = role_hours.get((name, role), 0.0)
        if role.startswith("FrontDesk"):
            base += 2.0 * sum(v for (nm, rl), v in role_hours.items() if nm == name and rl.startswith("FrontDesk"))
        cur = last_role_state.get((d, name))
        if cur and cur[0] == role and cur[1] > 0:
            base -= 0.5
        return base

    def candidate_ok(name, role, d, t):
        if not is_working(name, d, t) or is_on_break(name, d, t) or already_used(name, d, t):
            return False
        sr = staff_by_name[name]
        if not skill_allowed(sr, role):
            return False
        if not site_restriction_ok(sr, role):
            return False
        if not role_site_ok(role, d, sr):
            return False

        cur = last_role_state.get((d, name))
        if cur and cur[1] > 0 and cur[0] != role:
            return False

        if role.startswith("FrontDesk") and cont_fd.get((d, name), 0.0) + 0.5 > 2.5:
            return False
        if role.startswith("Triage_Admin") and cont_triage.get((d, name), 0.0) + 0.5 > 3.0:
            return False
        return True

    def force_carol(role, d, t):
        if not role.startswith("FrontDesk"):
            return None
        for _, sr in staff_df.iterrows():
            if sr.get("IsCarolChurch", False):
                nm = sr["Name"]
                if is_working(nm, d, t) and not is_on_break(nm, d, t):
                    home = str(sr.get("HomeSite","")).strip().upper()
                    if role.endswith(home):
                        return nm
        return None

    for d in dates:
        for t in slots:
            assign.setdefault((d, t), [])

            # mandatory list for this slot
            wanted = []
            for role, a, b, need in MANDATORY_RULES:
                if t_in_range(t, a, b):
                    wanted.append((role, need))
            if awaiting_required(d, t) or awaiting_optional(d, t):
                wanted.append(("Awaiting_PSA_Admin", 1))

            # fill mandatory
            for role, need in wanted:
                already = sum(1 for rl, _ in assign[(d, t)] if rl == role)
                while already < need:
                    forced = force_carol(role, d, t)
                    if forced and candidate_ok(forced, role, d, t):
                        pick = forced
                    else:
                        candidates = [n for n in avail.get(d, []) if candidate_ok(n, role, d, t)]
                        if not candidates:
                            gaps.append((d, t, role, "No suitable staff available"))
                            break
                        candidates.sort(key=lambda n: score(n, role, d))
                        pick = candidates[0]

                    assign[(d, t)].append((role, pick))

                    cur = last_role_state.get((d, pick))
                    if (not cur) or (cur[0] != role) or (cur[1] <= 0):
                        last_role_state[(d, pick)] = (role, MIN_STINT_SLOTS-1)

                    if role.startswith("FrontDesk"):
                        cont_fd[(d, pick)] = cont_fd.get((d, pick), 0.0) + 0.5
                    else:
                        cont_fd[(d, pick)] = 0.0
                    if role.startswith("Triage_Admin"):
                        cont_triage[(d, pick)] = cont_triage.get((d, pick), 0.0) + 0.5
                    else:
                        cont_triage[(d, pick)] = 0.0

                    role_hours[(pick, role)] = role_hours.get((pick, role), 0.0) + 0.5
                    already += 1

            # fillers: keep idle staff busy, with requested priorities
            filler_order = ["Phones", "Bookings", "Emis_Tasks", "Docman_Tasks", "Awaiting_PSA_Admin"]
            if d.weekday() in (0, 4):  # Mon/Fri
                filler_order = ["Phones", "Bookings", "Awaiting_PSA_Admin", "Emis_Tasks", "Docman_Tasks"]

            max_phones = 3
            current_phones = sum(1 for rl, _ in assign[(d, t)] if rl == "Phones")

            for role in filler_order:
                if role == "Phones" and current_phones >= max_phones:
                    continue

                if role == "Phones":
                    target = max_phones - current_phones
                elif role == "Awaiting_PSA_Admin":
                    target = 1 if awaiting_optional(d, t) else 0
                else:
                    target = 999

                for _ in range(target):
                    candidates = []
                    for name in avail.get(d, []):
                        if already_used(name, d, t):
                            continue
                        if not is_working(name, d, t) or is_on_break(name, d, t):
                            continue
                        cur = last_role_state.get((d, name))
                        if cur and cur[1] > 0 and cur[0] != role:
                            continue
                        sr = staff_by_name[name]
                        if not skill_allowed(sr, role):
                            continue
                        if not site_restriction_ok(sr, role):
                            continue
                        if role == "Awaiting_PSA_Admin" and not role_site_ok(role, d, sr):
                            continue
                        candidates.append(name)
                    if not candidates:
                        break

                    # SLGP: majority time on bookings (if trained)
                    if role != "Bookings":
                        filtered = []
                        for n in candidates:
                            sr = staff_by_name[n]
                            if str(sr.get("HomeSite","")).strip().upper() == "SLGP" and skill_allowed(sr, "Bookings"):
                                continue
                            filtered.append(n)
                        if filtered:
                            candidates = filtered

                    candidates.sort(key=lambda n: score(n, role, d))
                    pick = candidates[0]
                    assign[(d, t)].append((role, pick))

                    cur = last_role_state.get((d, pick))
                    if (not cur) or (cur[0] != role) or (cur[1] <= 0):
                        last_role_state[(d, pick)] = (role, MIN_STINT_SLOTS-1)

                    role_hours[(pick, role)] = role_hours.get((pick, role), 0.0) + 0.5
                    if role == "Phones":
                        current_phones += 1

    # Break check
    for d in dates:
        dname = ["Mon","Tue","Wed","Thu","Fri"][(d-week_start).days]
        for name in avail.get(d, []):
            hr = hmap[name]
            stt, end = shift_window(hr, dname)
            if not stt or not end:
                continue
            if h_between(stt, end) <= SHIFT_BREAK_THRESHOLD_HOURS:
                continue
            had_break = any(name in breaks.get((d, t), set()) for t in [time(12,0), time(12,30), time(13,0), time(13,30)])
            if not had_break:
                gaps.append((d, None, "Break", f"{name} shift > 6h but no break could be placed 12:00–14:00"))

    return assign, breaks, gaps


# =========================================================
# Excel output
# =========================================================
def fill_for_role(role):
    col = ROLE_COLORS.get(role, "FFFFFF")
    return PatternFill("solid", fgColor=col)

def role_to_site(role, d):
    if role.endswith("_SLGP"):
        return "SLGP"
    if role.endswith("_JEN"):
        return "JEN"
    if role.endswith("_BGS"):
        return "BGS"
    if role == "Bookings":
        return "SLGP"
    if role in ["Phones", "Email_Box", "Emis_Tasks", "Docman_Tasks"]:
        return "JEN/BGS"
    if role == "Awaiting_PSA_Admin":
        return psa_site_for_day(d)
    return ""

def build_staff_timeline(assign, breaks, week_start: date, staff_names):
    slots = timeslots()
    dates = [week_start + timedelta(days=i) for i in range(5)]
    rows = []
    for d in dates:
        for t in slots:
            for name in staff_names:
                if name in breaks.get((d, t), set()):
                    rows.append([name, d, t, "Break"])
                    continue
                roles = [rl for rl, nm in assign.get((d,t), []) if nm == name]
                rows.append([name, d, t, " + ".join(roles) if roles else ""])
    return pd.DataFrame(rows, columns=["Name","Date","Time","Task"])

def write_period_workbook(staff_df, hours_df, hols, start_monday: date, weeks: int):
    wb = Workbook()
    wb.remove(wb.active)

    for w in range(weeks):
        week_start = start_monday + timedelta(days=7*w)
        assign, breaks, gaps = rota_generate_week(staff_df, hours_df, hols, week_start)

        slots = timeslots()
        dates = [week_start + timedelta(days=i) for i in range(5)]

        # Grid
        ws_grid = wb.create_sheet(f"Week{w+1}_Grid")
        ws_grid.append(["Time"] + [d.strftime("%a %d-%b") for d in dates])
        for c in ws_grid[1]:
            c.font = Font(bold=True)
            c.alignment = Alignment(horizontal="center", vertical="center")

        for t in slots:
            row = [t.strftime("%H:%M")]
            for d in dates:
                slot_roles = assign.get((d, t), [])
                parts = []
                for role in [
                    "FrontDesk_SLGP","FrontDesk_JEN","FrontDesk_BGS",
                    "Triage_Admin_SLGP","Triage_Admin_JEN",
                    "Email_Box","Phones","Bookings","Awaiting_PSA_Admin",
                    "Emis_Tasks","Docman_Tasks",
                ]:
                    ppl = [nm for rl, nm in slot_roles if rl == role]
                    if ppl:
                        parts.append(f"{role}: " + ", ".join(ppl))
                row.append("\\n".join(parts))
            ws_grid.append(row)

        ws_grid.column_dimensions["A"].width = 8
        for col in range(2, 7):
            ws_grid.column_dimensions[chr(64+col)].width = 40
        for r in range(2, 2+len(slots)):
            ws_grid.row_dimensions[r].height = 90
            for c in range(2, 7):
                ws_grid.cell(r, c).alignment = Alignment(wrap_text=True, vertical="top")

        # Coverage gaps
        ws_gaps = wb.create_sheet(f"Week{w+1}_CoverageGaps")
        ws_gaps.append(["Date","Time","Role","Issue"])
        for c in ws_gaps[1]:
            c.font = Font(bold=True)
        for d, t, role, issue in gaps:
            ws_gaps.append([d.isoformat(), "" if t is None else t.strftime("%H:%M"), role, issue])

        # Totals
        ws_tot = wb.create_sheet(f"Week{w+1}_Totals")
        rows = []
        for d in dates:
            for t in slots:
                for rl, nm in assign.get((d,t), []):
                    rows.append([d, nm, rl, 0.5, role_to_site(rl, d)])
                for nm in breaks.get((d,t), set()):
                    rows.append([d, nm, "Break", 0.5, ""])
        df = pd.DataFrame(rows, columns=["Date","Name","Task","Hours","Site"])
        if df.empty:
            df = pd.DataFrame(columns=["Date","Name","Task","Hours","Site"])

        pivot_w = (df.groupby(["Name","Task"])["Hours"].sum()
                     .reset_index()
                     .pivot(index="Name", columns="Task", values="Hours")
                     .fillna(0.0))
        pivot_w["WeeklyTotal"] = pivot_w.sum(axis=1)
        pivot_w = pivot_w.reset_index()

        pivot_d = df.groupby(["Date","Name","Task"])["Hours"].sum().reset_index()
        site_tot = df[df["Site"].notna() & (df["Site"]!="")].groupby("Site")["Hours"].sum().reset_index().rename(columns={"Hours":"WeeklyHours"})

        ws_tot.append(["Weekly totals per staff by task (hours)"])
        ws_tot["A1"].font = Font(bold=True, size=12)
        ws_tot.append([])
        for r in dataframe_to_rows(pivot_w, index=False, header=True):
            ws_tot.append(r)

        header_row = 3
        for cell in ws_tot[header_row]:
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal="center", vertical="center")

        headers = [c.value for c in ws_tot[header_row]]
        for j, h in enumerate(headers, start=1):
            if not h or h == "Name":
                continue
            fill = fill_for_role(h)
            for irow in range(header_row, header_row + len(pivot_w) + 1):
                ws_tot.cell(irow, j).fill = fill

        ws_tot.append([])
        ws_tot.append(["Daily totals per staff by task (hours)"])
        ws_tot[f"A{ws_tot.max_row}"].font = Font(bold=True, size=12)
        ws_tot.append([])
        for r in dataframe_to_rows(pivot_d, index=False, header=True):
            ws_tot.append(r)

        ws_tot.append([])
        ws_tot.append(["Weekly hours totals per site (hours)"])
        ws_tot[f"A{ws_tot.max_row}"].font = Font(bold=True, size=12)
        ws_tot.append([])
        for r in dataframe_to_rows(site_tot, index=False, header=True):
            ws_tot.append(r)

        # Timelines: overall + per site
        staff_names = list(staff_df["Name"].astype(str))
        tl = build_staff_timeline(assign, breaks, week_start, staff_names)
        m = {(r.Date, r.Time, r.Name): r.Task for r in tl.itertuples(index=False)}

        ws_tl_all = wb.create_sheet(f"Week{w+1}_Timelines_All")
        ws_tl_all.append(["Date","Time"] + staff_names)
        for c in ws_tl_all[1]:
            c.font = Font(bold=True)
            c.alignment = Alignment(horizontal="center", vertical="center")
        ws_tl_all.freeze_panes = "C2"

        for d in dates:
            for t in slots:
                row = [d.strftime("%a %d-%b"), t.strftime("%H:%M")]
                for nm in staff_names:
                    row.append(m.get((d, t, nm), ""))
                ws_tl_all.append(row)

        for r in range(2, ws_tl_all.max_row+1):
            for c in range(3, ws_tl_all.max_column+1):
                val = ws_tl_all.cell(r,c).value
                if val:
                    role = val.split(" + ")[0]
                    ws_tl_all.cell(r,c).fill = fill_for_role(role)
                    ws_tl_all.cell(r,c).alignment = Alignment(wrap_text=True, vertical="top")

        for site in ["SLGP","JEN","BGS"]:
            site_staff = list(staff_df.loc[staff_df["HomeSite"].astype(str).str.upper()==site, "Name"].astype(str))
            if not site_staff:
                continue
            ws_site = wb.create_sheet(f"Week{w+1}_{site}_Timelines")
            ws_site.append(["Date","Time"] + site_staff)
            for c in ws_site[1]:
                c.font = Font(bold=True)
                c.alignment = Alignment(horizontal="center", vertical="center")
            ws_site.freeze_panes = "C2"

            for d in dates:
                for t in slots:
                    row = [d.strftime("%a %d-%b"), t.strftime("%H:%M")]
                    for nm in site_staff:
                        row.append(m.get((d, t, nm), ""))
                    ws_site.append(row)

            for r in range(2, ws_site.max_row+1):
                for c in range(3, ws_site.max_column+1):
                    val = ws_site.cell(r,c).value
                    if val:
                        role = val.split(" + ")[0]
                        ws_site.cell(r,c).fill = fill_for_role(role)
                        ws_site.cell(r,c).alignment = Alignment(wrap_text=True, vertical="top")

    return wb


# =========================================================
# Streamlit UI (no previews)
# =========================================================
st.set_page_config(page_title="Rota Generator", layout="wide")
require_password()

st.title("Rota Generator (Excel export)")

uploaded = st.file_uploader("Upload rota template (.xlsx)", type=["xlsx"])

c1, c2 = st.columns(2)
with c1:
    start_date = st.date_input("Week commencing (Monday)", value=date.today())
with c2:
    mode = st.selectbox("Run period", ["1 week", "4 weeks (month)", "Custom # weeks"])

if mode == "Custom # weeks":
    weeks = int(st.number_input("Number of weeks", min_value=1, max_value=12, value=2, step=1))
elif mode == "4 weeks (month)":
    weeks = 4
else:
    weeks = 1

start_monday = ensure_monday(start_date)

st.caption(
    "Enforced: min 1.5h stints, breaks 12:00–14:00 near midpoint, "
    "Front Desk max 2.5h continuous, Triage max 3h continuous, "
    "Phones (min 2) + Email (1) by JEN/BGS only, Bookings (min 3) SLGP, "
    "Awaiting/PSA Admin 10:00–16:00 rotating site (Mon/Fri SLGP, Tue/Thu JEN, Wed BGS) "
    "and optional after 16:00 if available, plus SLGP filler time forced to Bookings where trained."
)

if uploaded:
    try:
        staff_df, hours_df, hols, params, sheets = read_template(uploaded.getvalue())
        st.success("Template loaded. Click Generate to create the Excel rota.")

        if st.button("Generate rota and download Excel", type="primary"):
            wb = write_period_workbook(staff_df, hours_df, hols, start_monday, weeks)
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
    except Exception as e:
        st.error("Could not process the template.")
        st.exception(e)
else:
    st.info("Upload your completed template to continue.")
