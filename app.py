import io
import re
from datetime import datetime, date, time, timedelta

import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils.dataframe import dataframe_to_rows

# =========================================================
# Rota Generator v9 — block-based engine
# =========================================================
# What’s new vs v7/v8:
# - A true block scheduler: assigns tasks in 2h minimum blocks (unless end-of-shift/break).
# - Max block 3h for all tasks except Email (Email can run longer).
# - Front Desk is enforced as exactly 1 per site per slot.
# - Cross-site is only used when NO home-site staff can cover that role at that time (never for Front Desk; Bookings stays SLGP-only).
# - Timeline cells explicitly show Holiday / Sick / Bank Holiday / Not working and are color-coded.
#
# Template expectations (as you’ve been using):
# Staff sheet: Name, HomeSite, CanFrontDesk, CanTriage, CanEmail, CanPhones, CanBookings, CanEMIS, CanDocman_PSA, CanDocman_AWAIT, (optional) IsCarolChurch(Y/N)
# WorkingHours sheet: Name, MonStart/MonEnd ... FriStart/FriEnd
# Holidays sheet (optional): Name, Start, End, Notes (Notes => default Holiday unless contains "sick"/"sickness" or "bank")
#
# Upload this file as app.py on Streamlit Cloud.

# ---------------- Password (Streamlit secrets TOML) ----------------
def require_password():
    """
    In Streamlit Secrets (TOML):
      APP_PASSWORD = "your-strong-password"
    """
    pw = st.secrets.get("APP_PASSWORD")
    if not pw:
        st.warning("APP_PASSWORD not set in Streamlit Secrets. App is not password-protected.")
        return True
    if st.session_state.get("authed"):
        return True
    with st.form("login"):
        entered = st.text_input("Password", type="password")
        ok = st.form_submit_button("Log in")
        if ok and entered == pw:
            st.session_state.authed = True
            st.success("Logged in.")
            return True
        if ok:
            st.error("Incorrect password.")
    st.stop()

# ---------------- Helpers ----------------
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
    # substring fallback
    for c in df.columns:
        nc = normalize(c)
        for cand in candidates:
            if normalize(cand) in nc:
                return c
    if required:
        raise KeyError(f"Missing required column among {candidates}. Available: {list(df.columns)}")
    return None

def to_time(x):
    if pd.isna(x) or x == "":
        return None
    if isinstance(x, time):
        return x
    if isinstance(x, datetime):
        return x.time()
    if isinstance(x, (float, int)):
        # excel time fraction of day
        seconds = int(round(float(x) * 86400))
        return (datetime(2000, 1, 1) + timedelta(seconds=seconds)).time()
    return pd.to_datetime(str(x)).time()

def to_date(x):
    if pd.isna(x) or x == "":
        return None
    if isinstance(x, (datetime, pd.Timestamp)):
        return x.date()
    if isinstance(x, date):
        return x
    return pd.to_datetime(x).date()

def dt_of(d: date, t: time) -> datetime:
    return datetime(d.year, d.month, d.day, t.hour, t.minute)

def add_minutes(t: time, mins: int) -> time:
    return (datetime(2000, 1, 1, t.hour, t.minute) + timedelta(minutes=mins)).time()

def t_in_range(t: time, a: time, b: time) -> bool:
    return (t >= a) and (t < b)

def ensure_monday(d: date) -> date:
    return d - timedelta(days=d.weekday())

# ---------------- Read template ----------------
def read_template(uploaded_bytes: bytes):
    xls = pd.ExcelFile(io.BytesIO(uploaded_bytes))
    staff_sheet = find_sheet(xls, ["Staff", "Skills", "People"])
    hours_sheet = find_sheet(xls, ["WorkingHours", "Hours", "Availability"])
    hols_sheet = find_sheet(xls, ["Holidays", "Leave", "Absence"])

    if not staff_sheet:
        raise ValueError(f"Could not find Staff sheet. Found: {xls.sheet_names}")
    if not hours_sheet:
        raise ValueError(f"Could not find WorkingHours sheet. Found: {xls.sheet_names}")

    staff_df = pd.read_excel(xls, sheet_name=staff_sheet)
    hours_df = pd.read_excel(xls, sheet_name=hours_sheet)
    hols_df = pd.read_excel(xls, sheet_name=hols_sheet) if hols_sheet else pd.DataFrame()

    name_c = pick_col(staff_df, ["Name", "StaffName"])
    home_c = pick_col(staff_df, ["HomeSite", "Site", "BaseSite"], required=False)

    staff_df = staff_df.copy()
    staff_df["Name"] = staff_df[name_c].astype(str).str.strip()
    staff_df["HomeSite"] = staff_df[home_c].astype(str).str.strip().str.upper() if home_c else ""

    def yn(v):
        if pd.isna(v): return False
        s = str(v).strip().lower()
        return s in ["y", "yes", "true", "1"]

    for c in staff_df.columns:
        if normalize(c).startswith("can") or normalize(c) in ["iscarolchurchyn", "iscarolchurch"]:
            staff_df[c] = staff_df[c].apply(yn)

    carol_c = pick_col(staff_df, ["IsCarolChurch(Y/N)", "IsCarolChurch", "Carol"], required=False)
    if carol_c:
        staff_df["IsCarolChurch"] = staff_df[carol_c].apply(bool)
    else:
        staff_df["IsCarolChurch"] = staff_df["Name"].str.lower().eq("carol church")

    hours_df = hours_df.copy()
    hours_name_c = pick_col(hours_df, ["Name", "StaffName"])
    hours_df["Name"] = hours_df[hours_name_c].astype(str).str.strip()

    for d in ["Mon", "Tue", "Wed", "Thu", "Fri"]:
        sc = pick_col(hours_df, [f"{d}Start", f"{d} Start", f"{d}_Start"], required=False)
        ec = pick_col(hours_df, [f"{d}End", f"{d} End", f"{d}_End"], required=False)
        hours_df[f"{d}Start"] = hours_df[sc].apply(to_time) if sc else None
        hours_df[f"{d}End"] = hours_df[ec].apply(to_time) if ec else None

    # Holidays with Notes => default Holiday unless contains "sickness/sick" or "bank"
    hols = []
    if not hols_df.empty:
        hn = pick_col(hols_df, ["Name", "StaffName"], required=False) or hols_df.columns[0]
        hs = pick_col(hols_df, ["StartDate", "Start"], required=False) or hols_df.columns[1]
        he = pick_col(hols_df, ["EndDate", "End"], required=False) or hols_df.columns[2]
        notes_c = pick_col(hols_df, ["Notes", "Note", "Reason"], required=False)
        for _, r in hols_df.iterrows():
            nm = str(r[hn]).strip()
            sd = to_date(r[hs])
            ed = to_date(r[he])
            note = "" if (not notes_c or pd.isna(r[notes_c])) else str(r[notes_c]).strip().lower()
            kind = "Holiday"
            if "sick" in note or "sickness" in note:
                kind = "Sick"
            elif "bank" in note:
                kind = "Bank Holiday"
            hols.append((nm, sd, ed, kind))

    return staff_df, hours_df, hols

# ---------------- Business rules ----------------
DAY_START = time(8, 0)
DAY_END = time(18, 30)
SLOT_MIN = 30

# User requirement: min 2 hours, max 3 hours (except email)
MIN_BLOCK_SLOTS = 4  # 2h
MAX_BLOCK_SLOTS = 6  # 3h

# Breaks: start times fixed to 12:00–14:00 window, pick nearest midpoint
BREAK_WINDOW = (time(12, 0), time(14, 0))
BREAK_THRESHOLD_HOURS = 6.0

def timeslots():
    cur = datetime(2000, 1, 1, DAY_START.hour, DAY_START.minute)
    end = datetime(2000, 1, 1, DAY_END.hour, DAY_END.minute)
    out = []
    while cur < end:
        out.append(cur.time())
        cur += timedelta(minutes=SLOT_MIN)
    return out

MANDATORY_RULES = [
    ("FrontDesk_SLGP", DAY_START, DAY_END, 1),
    ("FrontDesk_JEN",  DAY_START, DAY_END, 1),
    ("FrontDesk_BGS",  DAY_START, DAY_END, 1),

    ("Triage_Admin_SLGP", DAY_START, time(16, 0), 1),
    ("Triage_Admin_JEN",  DAY_START, time(16, 0), 1),

    ("Email_Box", DAY_START, DAY_END, 1),

    ("Phones", DAY_START, DAY_END, 2),
    ("Bookings", DAY_START, DAY_END, 3),
]

def awaiting_site_for_day(d: date) -> str:
    # Mon/Fri SLGP, Tue/Thu JEN, Wed BGS
    wd = d.weekday()
    if wd in (0, 4):
        return "SLGP"
    if wd in (1, 3):
        return "JEN"
    return "BGS"

def awaiting_required(d: date, t: time) -> bool:
    return t_in_range(t, time(10, 0), time(16, 0))

def awaiting_optional(d: date, t: time) -> bool:
    return t_in_range(t, time(16, 0), DAY_END)

def filler_order_for_day(d: date):
    # Mon/Fri: phones priority
    if d.weekday() in (0, 4):
        return ["Phones", "Bookings", "Awaiting_PSA_Admin", "Emis_Tasks", "Docman_Tasks"]
    return ["Phones", "Bookings", "Emis_Tasks", "Docman_Tasks", "Awaiting_PSA_Admin"]

# ---------------- Skills mapping ----------------
def can(sr, col: str) -> bool:
    return bool(sr.get(col, False))

def skill_allowed(sr, role: str) -> bool:
    if role.startswith("FrontDesk"):
        return can(sr, "CanFrontDesk")
    if role.startswith("Triage_Admin"):
        return can(sr, "CanTriage")
    if role == "Email_Box":
        return can(sr, "CanEmail")
    if role == "Phones":
        return can(sr, "CanPhones")
    if role == "Bookings":
        return can(sr, "CanBookings")
    if role == "Emis_Tasks":
        return can(sr, "CanEMIS")
    if role == "Docman_Tasks":
        return can(sr, "CanDocman_PSA") or can(sr, "CanDocman_AWAIT")
    if role == "Awaiting_PSA_Admin":
        return can(sr, "CanDocman_PSA") or can(sr, "CanDocman_AWAIT")
    if role == "Unassigned":
        return True
    return False

def role_site(role: str, d: date) -> str:
    if role.endswith("_SLGP"):
        return "SLGP"
    if role.endswith("_JEN"):
        return "JEN"
    if role.endswith("_BGS"):
        return "BGS"
    if role == "Bookings":
        return "SLGP"
    if role == "Awaiting_PSA_Admin":
        return awaiting_site_for_day(d)
    return ""  # shared

def is_frontdesk(role: str) -> bool:
    return role.startswith("FrontDesk_")

def site_allowed(sr, role: str, d: date, allow_cross: bool) -> bool:
    home = str(sr.get("HomeSite", "")).strip().upper()
    if is_frontdesk(role):
        return home == role_site(role, d)
    if role == "Bookings":
        return home == "SLGP"
    if role.startswith("Triage_Admin"):
        target = role_site(role, d)
        return (home == target) or allow_cross
    if role == "Awaiting_PSA_Admin":
        target = awaiting_site_for_day(d)
        return (home == target) or allow_cross
    if role in ["Email_Box", "Phones", "Emis_Tasks", "Docman_Tasks"]:
        if home in ["JEN", "BGS"]:
            return True
        return allow_cross
    return True

# ---------------- Holidays / working hours ----------------
def holiday_kind(name: str, d: date, hols):
    for n, s, e, k in hols:
        if n.strip().lower() == name.strip().lower() and s and e and s <= d <= e:
            return k
    return None

def build_maps(staff_df, hours_df):
    staff_by_name = {r["Name"]: r for _, r in staff_df.iterrows()}
    hmap = {r["Name"]: r for _, r in hours_df.iterrows()}
    return staff_by_name, hmap

def is_working(hmap, week_start, d, t, name):
    days = ["Mon", "Tue", "Wed", "Thu", "Fri"]
    dname = days[(d - week_start).days]
    hr = hmap.get(name)
    if hr is None:
        return False
    stt = hr.get(f"{dname}Start")
    end = hr.get(f"{dname}End")
    return bool(stt and end and (t >= stt) and (t < end))

def shift_window(hmap, week_start, d, name):
    days = ["Mon", "Tue", "Wed", "Thu", "Fri"]
    dname = days[(d - week_start).days]
    hr = hmap.get(name)
    if hr is None:
        return None, None
    return hr.get(f"{dname}Start"), hr.get(f"{dname}End")

def pick_break_slot_near_midpoint(d: date, stt: time, end: time):
    if not stt or not end:
        return None
    if (dt_of(d, end) - dt_of(d, stt)).total_seconds() / 3600.0 <= BREAK_THRESHOLD_HOURS:
        return None
    midpoint = dt_of(d, stt) + (dt_of(d, end) - dt_of(d, stt)) / 2
    candidates = []
    for t in [time(12, 0), time(12, 30), time(13, 0), time(13, 30)]:
        if t >= stt and add_minutes(t, 30) <= end and t_in_range(t, BREAK_WINDOW[0], BREAK_WINDOW[1]):
            dist = abs((dt_of(d, t) - midpoint).total_seconds())
            candidates.append((dist, t))
    if not candidates:
        return None
    candidates.sort(key=lambda x: x[0])
    return candidates[0][1]

# =========================================================
# Block-based scheduler
# =========================================================
def rota_generate_week(staff_df, hours_df, hols, week_start: date):
    slots = timeslots()
    dates = [week_start + timedelta(days=i) for i in range(5)]
    staff_by_name, hmap = build_maps(staff_df, hours_df)

    # breaks: (d,t)->set(names)
    breaks = {}
    for d in dates:
        for name in staff_by_name.keys():
            if holiday_kind(name, d, hols):
                continue
            stt, end = shift_window(hmap, week_start, d, name)
            b = pick_break_slot_near_midpoint(d, stt, end)
            if b:
                breaks.setdefault((d, b), set()).add(name)

    # assignment map: (d,t,name)->role
    a = {}

    # track block for each (d,name): (role, site, end_idx_exclusive)
    block = {}

    # per-day role usage counts to avoid "same all day"
    role_minutes = {}  # (d,name,role)->minutes

    gaps = []

    def on_break(name, d, t):
        return name in breaks.get((d, t), set())

    def is_free(name, d, t):
        return (d, t, name) not in a

    def slot_index(t):
        return slots.index(t)

    def has_role_at(role, d, t):
        return [nm for (dd, tt, nm), rr in a.items() if dd == d and tt == t and rr == role]

    def needed_for_slot(d, t):
        wanted = []
        for role, ta, tb, need in MANDATORY_RULES:
            if t_in_range(t, ta, tb):
                wanted.append((role, need))
        if awaiting_required(d, t) or awaiting_optional(d, t):
            wanted.append(("Awaiting_PSA_Admin", 1))
        return wanted

    def no_home_site_candidates(role, d, t):
        # Determine if any staff on the target site (for site-bound roles) can cover at this time.
        target_site = role_site(role, d)
        for name, sr in staff_by_name.items():
            if holiday_kind(name, d, hols):
                continue
            if not is_working(hmap, week_start, d, t, name):
                continue
            if on_break(name, d, t):
                continue
            if not is_free(name, d, t):
                continue
            if not skill_allowed(sr, role):
                continue
            # home-site only check
            if target_site and str(sr.get("HomeSite","")).upper() != target_site:
                continue
            # for shared roles, treat "home-site" as their actual HomeSite (we don't force)
            return False
        return True

    def candidate_ok(name, role, d, t, allow_cross):
        sr = staff_by_name[name]

        if holiday_kind(name, d, hols):
            return False
        if not is_working(hmap, week_start, d, t, name):
            return False
        if on_break(name, d, t):
            return False
        if not is_free(name, d, t):
            return False
        if not skill_allowed(sr, role):
            return False

        # existing block?
        b = block.get((d, name))
        if b is not None:
            # if already in a block, must keep that role until block end
            if b[0] != role:
                return False

        # cross-site only if no home-site candidates exist (and never for FrontDesk)
        if not site_allowed(sr, role, d, allow_cross=allow_cross):
            return False

        # front desk must stay on their own site always
        if is_frontdesk(role) and str(sr.get("HomeSite","")).upper() != role_site(role, d):
            return False

        return True

    def pick_block_len(name, role, d, start_idx):
        """Choose a block length in slots: >= MIN_BLOCK_SLOTS unless end-of-shift/break arrives sooner.
           For non-email, cap at MAX_BLOCK_SLOTS. Prefer 3h early, 2h later, and avoid same all day."""
        stt, end = shift_window(hmap, week_start, d, name)
        if not stt or not end:
            return 1
        # latest index we can work
        end_idx = start_idx
        while end_idx < len(slots) and slots[end_idx] < end:
            end_idx += 1

        # if break within next N slots, cap to that
        break_idx = None
        for k in range(start_idx, min(len(slots), end_idx)):
            if name in breaks.get((d, slots[k]), set()):
                break_idx = k
                break
        hard_end = break_idx if break_idx is not None else end_idx

        remaining = hard_end - start_idx
        if remaining <= 0:
            return 0

        # end-of-shift remainder can be < MIN
        if remaining < MIN_BLOCK_SLOTS:
            return remaining

        # avoid "same all day" by limiting repeats:
        used_mins = role_minutes.get((d, name, role), 0)
        # if already did >= 6h of this role, force 2h blocks from now (except email)
        if role != "Email_Box" and used_mins >= 360:
            return MIN_BLOCK_SLOTS

        # prefer 3h blocks if it fits and not near end
        if role == "Email_Box":
            # Email: can be longer, but still don’t force all day; cap to 4h blocks for gentler rotation
            return min(8, remaining)  # 4h max for email blocks here
        return min(MAX_BLOCK_SLOTS, remaining)

    def start_block(name, role, d, start_idx):
        L = pick_block_len(name, role, d, start_idx)
        if L <= 0:
            return False
        s = role_site(role, d) or str(staff_by_name[name].get("HomeSite","")).upper()
        block[(d, name)] = (role, s, start_idx + L)
        return True

    def fill_block(name, d, idx):
        """If staff has an active block, apply its role to this slot."""
        b = block.get((d, name))
        if not b:
            return False
        role, _, end_idx = b
        if idx >= end_idx:
            del block[(d, name)]
            return False
        t = slots[idx]
        a[(d, t, name)] = role
        role_minutes[(d, name, role)] = role_minutes.get((d, name, role), 0) + SLOT_MIN
        return True

    def enforce_frontdesk_exactly_one(d, idx):
        t = slots[idx]
        for site in ["SLGP", "JEN", "BGS"]:
            role = f"FrontDesk_{site}"
            current = has_role_at(role, d, t)
            if len(current) > 1:
                # should never happen; mark gaps and keep first
                gaps.append((d, t, role, f"More than 1 assigned ({', '.join(current)})"))
                for extra in current[1:]:
                    a[(d, t, extra)] = "Unassigned"
            if len(current) == 1:
                continue
            # need exactly 1
            # candidate search: first within home site; no cross-site permitted for FD
            cands = []
            for name, sr in staff_by_name.items():
                if str(sr.get("HomeSite","")).upper() != site:
                    continue
                if candidate_ok(name, role, d, t, allow_cross=False):
                    cands.append(name)
            # prefer Carol Church if applicable
            if "Carol Church" in cands:
                pick = "Carol Church"
            elif cands:
                pick = cands[0]
            else:
                gaps.append((d, t, role, "No suitable staff available"))
                continue
            if (d, pick) not in block:
                start_block(pick, role, d, idx)
            fill_block(pick, d, idx)

    def enforce_role_need(role, need, d, idx):
        t = slots[idx]
        current = has_role_at(role, d, t)
        while len(current) < need:
            # decide whether cross-site is allowed
            allow_cross = (not is_frontdesk(role)) and no_home_site_candidates(role, d, t)
            # candidate pool
            cands = []
            for name in staff_by_name.keys():
                if candidate_ok(name, role, d, t, allow_cross=allow_cross):
                    # site-sticky: prefer same site blocks
                    prev_block = block.get((d, name))
                    penalty = 0
                    if prev_block:
                        # if they are blocked to other role, they won't be candidate_ok anyway
                        penalty = 0
                    # deprioritise if they've already done lots of this role today
                    penalty += role_minutes.get((d, name, role), 0) / 60.0
                    cands.append((penalty, name))
            if not cands:
                if role != "Awaiting_PSA_Admin" or awaiting_required(d, t):
                    gaps.append((d, t, role, "No suitable staff available"))
                break
            cands.sort(key=lambda x: x[0])
            pick = cands[0][1]
            if (d, pick) not in block:
                start_block(pick, role, d, idx)
            fill_block(pick, d, idx)
            current = has_role_at(role, d, t)

    # main loop: per day, per slot
    for d in dates:
        for idx in range(len(slots)):
            t = slots[idx]

            # 1) extend existing blocks into this slot
            for name in staff_by_name.keys():
                # don't overwrite break/holiday; blocks won't exist if not working
                if (d, t, name) in a:
                    continue
                fill_block(name, d, idx)

            # 2) Enforce Front Desk exactly one per site
            enforce_frontdesk_exactly_one(d, idx)

            # 3) Enforce other mandatory coverage
            wanted = needed_for_slot(d, t)
            for role, need in wanted:
                if role.startswith("FrontDesk_"):
                    # already handled above as exact 1, but ensure need=1 coverage
                    continue
                enforce_role_need(role, need, d, idx)

            # 4) Fill everyone who is working and not on break and not assigned yet
            filler_roles = filler_order_for_day(d)
            for name, sr in staff_by_name.items():
                if holiday_kind(name, d, hols):
                    continue
                if not is_working(hmap, week_start, d, t, name):
                    continue
                if on_break(name, d, t):
                    continue
                if not is_free(name, d, t):
                    continue

                # if they have a continuing block, it would have filled already; so start a new filler block
                chosen = None

                # SLGP majority on bookings: if SLGP and bookings trained, prefer bookings
                if str(sr.get("HomeSite","")).upper() == "SLGP" and skill_allowed(sr, "Bookings"):
                    if site_allowed(sr, "Bookings", d, allow_cross=False) and skill_allowed(sr, "Bookings"):
                        chosen = "Bookings"

                if not chosen:
                    # prefer NOT repeating same role all day: pick role with least minutes so far today
                    best = None
                    for role in filler_roles:
                        if role == "Awaiting_PSA_Admin" and not awaiting_optional(d, t):
                            continue
                        if not skill_allowed(sr, role):
                            continue
                        allow_cross = (not is_frontdesk(role)) and no_home_site_candidates(role, d, t)
                        if not site_allowed(sr, role, d, allow_cross=allow_cross):
                            continue
                        # avoid ultra-long same-role days (except email): if already >= 6h on this role, deprioritise
                        used = role_minutes.get((d, name, role), 0)
                        if role != "Email_Box" and used >= 360:
                            continue
                        score = used
                        if best is None or score < best[0]:
                            best = (score, role)
                    if best:
                        chosen = best[1]

                if not chosen:
                    chosen = "Unassigned"

                if (d, name) not in block:
                    start_block(name, chosen, d, idx)
                fill_block(name, d, idx)

    # verify break placement if >6h
    for d in dates:
        for name in staff_by_name.keys():
            if holiday_kind(name, d, hols):
                continue
            stt, end = shift_window(hmap, week_start, d, name)
            if not stt or not end:
                continue
            if (dt_of(d, end) - dt_of(d, stt)).total_seconds() / 3600.0 <= BREAK_THRESHOLD_HOURS:
                continue
            had_break = any(name in breaks.get((d, bt), set()) for bt in [time(12,0), time(12,30), time(13,0), time(13,30)])
            if not had_break:
                gaps.append((d, None, "Break", f"{name}: shift > 6h but no break could be placed 12:00–14:00"))

    return a, breaks, gaps, hmap

# =========================================================
# Excel output
# =========================================================
ROLE_COLORS = {
    # roles
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
    "Unassigned": "FFFFFF",
    # statuses
    "Break": "CFE2F3",
    "Not working": "DDDDDD",
    "Holiday": "FFF2CC",
    "Bank Holiday": "FFE599",
    "Sick": "F4CCCC",
}

def fill_for(value: str):
    return PatternFill("solid", fgColor=ROLE_COLORS.get(value, "FFFFFF"))

def build_workbook(staff_df, hours_df, hols, start_monday: date, weeks: int):
    wb = Workbook()
    wb.remove(wb.active)
    slots = timeslots()

    for w in range(weeks):
        week_start = start_monday + timedelta(days=7*w)
        dates = [week_start + timedelta(days=i) for i in range(5)]
        staff_names = list(staff_df["Name"].astype(str))

        a, breaks, gaps, hmap = rota_generate_week(staff_df, hours_df, hols, week_start)

        # Week grid
        ws_grid = wb.create_sheet(f"Week{w+1}_Grid")
        ws_grid.append(["Time"] + [d.strftime("%a %d-%b") for d in dates])
        for c in ws_grid[1]:
            c.font = Font(bold=True)
            c.alignment = Alignment(horizontal="center", vertical="center")

        for t in slots:
            row = [t.strftime("%H:%M")]
            for d in dates:
                parts = []
                for role in [
                    "FrontDesk_SLGP","FrontDesk_JEN","FrontDesk_BGS",
                    "Triage_Admin_SLGP","Triage_Admin_JEN",
                    "Email_Box","Phones","Bookings",
                    "Awaiting_PSA_Admin","Emis_Tasks","Docman_Tasks",
                ]:
                    ppl = [nm for nm in staff_names if a.get((d, t, nm)) == role]
                    if ppl:
                        parts.append(f"{role}: " + ", ".join(ppl))
                row.append("\n".join(parts))
            ws_grid.append(row)

        ws_grid.column_dimensions["A"].width = 8
        for col in range(2, 7):
            ws_grid.column_dimensions[chr(64+col)].width = 42
        for r in range(2, 2+len(slots)):
            ws_grid.row_dimensions[r].height = 95
            for c in range(2, 7):
                ws_grid.cell(r, c).alignment = Alignment(wrap_text=True, vertical="top")

        # Coverage gaps
        ws_gaps = wb.create_sheet(f"Week{w+1}_CoverageGaps")
        ws_gaps.append(["Date", "Time", "Role", "Issue"])
        for c in ws_gaps[1]:
            c.font = Font(bold=True)
        for d, t, role, issue in gaps:
            ws_gaps.append([d.isoformat(), "" if t is None else t.strftime("%H:%M"), role, issue])

        # Totals
        ws_tot = wb.create_sheet(f"Week{w+1}_Totals")
        rows = []
        for d in dates:
            for t in slots:
                for nm in staff_names:
                    role = a.get((d, t, nm))
                    if role:
                        rows.append([d, nm, role, 0.5])
                for nm in breaks.get((d, t), set()):
                    rows.append([d, nm, "Break", 0.5])
        df = pd.DataFrame(rows, columns=["Date","Name","Task","Hours"])
        if df.empty:
            df = pd.DataFrame(columns=["Date","Name","Task","Hours"])

        pivot_w = (df.groupby(["Name","Task"])["Hours"].sum()
                     .reset_index()
                     .pivot(index="Name", columns="Task", values="Hours")
                     .fillna(0.0))
        pivot_w["WeeklyTotal"] = pivot_w.sum(axis=1)
        pivot_w = pivot_w.reset_index()

        pivot_d = df.groupby(["Date","Name","Task"])["Hours"].sum().reset_index()

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
            f = fill_for(h)
            for irow in range(header_row, header_row + len(pivot_w) + 1):
                ws_tot.cell(irow, j).fill = f

        ws_tot.append([])
        ws_tot.append(["Daily totals per staff by task (hours)"])
        ws_tot[f"A{ws_tot.max_row}"].font = Font(bold=True, size=12)
        ws_tot.append([])
        for r in dataframe_to_rows(pivot_d, index=False, header=True):
            ws_tot.append(r)

        # Timelines
        ws_tl = wb.create_sheet(f"Week{w+1}_Timelines_All")
        ws_tl.append(["Date", "Time"] + staff_names)
        for c in ws_tl[1]:
            c.font = Font(bold=True)
            c.alignment = Alignment(horizontal="center", vertical="center")
        ws_tl.freeze_panes = "C2"

        for d in dates:
            for t in slots:
                row = [d.strftime("%a %d-%b"), t.strftime("%H:%M")]
                for nm in staff_names:
                    hk = holiday_kind(nm, d, hols)
                    if hk:
                        row.append(hk)
                    elif not is_working(hmap, week_start, d, t, nm):
                        row.append("Not working")
                    elif nm in breaks.get((d, t), set()):
                        row.append("Break")
                    else:
                        row.append(a.get((d, t, nm), "Unassigned"))
                ws_tl.append(row)

        for r in range(2, ws_tl.max_row + 1):
            for c in range(3, ws_tl.max_column + 1):
                val = str(ws_tl.cell(r, c).value)
                ws_tl.cell(r, c).fill = fill_for(val)
                ws_tl.cell(r, c).alignment = Alignment(wrap_text=True, vertical="top")

        # By site
        for site in ["SLGP","JEN","BGS"]:
            site_staff = list(staff_df.loc[staff_df["HomeSite"].astype(str).str.upper()==site, "Name"].astype(str))
            if not site_staff:
                continue
            ws_site = wb.create_sheet(f"Week{w+1}_{site}_Timelines")
            ws_site.append(["Date", "Time"] + site_staff)
            for c in ws_site[1]:
                c.font = Font(bold=True)
                c.alignment = Alignment(horizontal="center", vertical="center")
            ws_site.freeze_panes = "C2"
            for d in dates:
                for t in slots:
                    row = [d.strftime("%a %d-%b"), t.strftime("%H:%M")]
                    for nm in site_staff:
                        hk = holiday_kind(nm, d, hols)
                        if hk:
                            row.append(hk)
                        elif not is_working(hmap, week_start, d, t, nm):
                            row.append("Not working")
                        elif nm in breaks.get((d, t), set()):
                            row.append("Break")
                        else:
                            row.append(a.get((d, t, nm), "Unassigned"))
                    ws_site.append(row)
            for r in range(2, ws_site.max_row + 1):
                for c in range(3, ws_site.max_column + 1):
                    val = str(ws_site.cell(r, c).value)
                    ws_site.cell(r, c).fill = fill_for(val)
                    ws_site.cell(r, c).alignment = Alignment(wrap_text=True, vertical="top")

    return wb

# =========================================================
# Streamlit UI
# =========================================================
st.set_page_config(page_title="Rota Generator v9 (block engine)", layout="wide")
require_password()

st.title("Rota Generator v9 — Block-based engine")

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
    "v9: assigns tasks in 2h minimum blocks (unless end-of-shift/break), max 3h blocks (except Email). "
    "Front Desk is exactly 1 per site at all times. Awaiting/PSA Admin mandatory 10:00–16:00 "
    "(Mon/Fri SLGP, Tue/Thu JEN, Wed BGS). Holidays Notes: default Holiday unless includes 'sick/sickness' or 'bank'."
)

if uploaded:
    try:
        staff_df, hours_df, hols = read_template(uploaded.getvalue())
        st.success("Template loaded.")

        if st.button("Generate rota and download Excel", type="primary"):
            wb = build_workbook(staff_df, hours_df, hols, start_monday, weeks)
            bio = io.BytesIO()
            wb.save(bio)
            bio.seek(0)
            out_name = f"rota_{start_monday.isoformat()}_{weeks}w.xlsx"
            st.download_button(
                "📊 Download Excel rota",
                data=bio.getvalue(),
                file_name=out_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    except Exception as e:
        st.error("Could not process the template.")
        st.exception(e)
else:
    st.info("Upload your completed template to continue.")
