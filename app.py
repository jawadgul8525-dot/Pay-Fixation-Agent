import streamlit as st
import datetime
import io
import os
import re
import json
import zipfile
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage
from reportlab.lib.styles import getSampleStyleSheet

# ==============================================================================
# 0. PERSISTENCE LAYER & STORAGE DIRECTORIES
# ==============================================================================
DATA_DIR = os.path.join(os.path.dirname(__file__), "data")
REGISTRY_FILE = os.path.join(DATA_DIR, "chart_registry.json")
os.makedirs(DATA_DIR, exist_ok=True)

# Default historical bootstrap (Used ONLY until the first chart is activated)
BOOTSTRAP_CAF_SCALES = {
    "NCs E": {
        "1994-06-01": (1245, 35, 1770), "2001-12-01": (1870, 55, 3520),
        "2005-07-01_REV": (2130, 65, 4100), "2005-07-01_UPG": (2130, 65, 4100),
        "2007-07-01": (2475, 75, 4725), "2008-07-01": (2970, 90, 5670),
        "2011-07-01": (4800, 150, 9300), "2015-07-01": (6210, 195, 12060),
        "2016-07-01": (7610, 240, 14840), "2017-07-01": (9125, 285, 17675),
        "2022-07-01": (13540, 430, 26440), "2024-06-01": (15620, 490, 30320),
        "2026-07-01": (18780, 590, 36480)
    },
    "GD Sep": {
        "1994-06-01": (1254, 30, 1800), "2001-12-01": (1875, 75, 3525),
        "2005-07-01_REV": (2155, 85, 4025), "2005-07-01_UPG": (2415, 115, 5865),
        "2007-07-01": (2780, 135, 6830), "2008-07-01": (3340, 160, 8140),
        "2011-07-01": (5400, 270, 13500), "2015-07-01": (6985, 350, 17485),
        "2016-07-01": (8590, 430, 21490), "2017-07-01": (10260, 510, 25560),
        "2022-07-01": (15230, 760, 38030), "2024-06-01": (18160, 950, 46660),
        "2026-07-01": (21830, 1150, 56330)
    },
    "GD Lnk": {
        "1994-06-01": (1262, 61, 1933), "2001-12-01": (1895, 95, 3985),
        "2005-07-01_REV": (2180, 110, 4600), "2005-07-01_UPG": (2485, 125, 6235),
        "2007-07-01": (2860, 145, 7210), "2008-07-01": (3430, 175, 8680),
        "2011-07-01": (5500, 290, 14200), "2015-07-01": (7115, 375, 18365),
        "2016-07-01": (8750, 470, 22850), "2017-07-01": (10465, 560, 27265),
        "2022-07-01": (15530, 840, 40730), "2024-06-01": (18780, 1040, 49980),
        "2026-07-01": (22580, 1250, 60080)
    },
    "GD Nk": {
        "1994-06-01": (1400, 66, 2390), "2001-12-01": (2400, 100, 5100),
        "2005-07-01_REV": (2415, 115, 5865), "2005-07-01_UPG": (2555, 140, 6755),
        "2007-07-01": (2940, 160, 7740), "2008-07-01": (3530, 190, 9230),
        "2011-07-01": (5700, 310, 15000), "2015-07-01": (7370, 405, 19520),
        "2016-07-01": (9070, 500, 24070), "2017-07-01": (10835, 595, 28665),
        "2022-07-01": (16080, 890, 42780), "2024-06-01": (19450, 1140, 53650),
        "2026-07-01": (23380, 1380, 64780)
    },
    "GD Hav": {
        "1994-06-01": (1440, 73, 2535), "2001-12-01": (2160, 110, 5460),
        "2005-07-01_REV": (2485, 125, 6235), "2005-07-01_UPG": (2655, 150, 7155),
        "2007-07-01": (3055, 175, 8305), "2008-07-01": (3665, 210, 9965),
        "2011-07-01": (5900, 350, 16400), "2015-07-01": (7630, 452, 21190),
        "2016-07-01": (9390, 560, 26190), "2017-07-01": (11225, 665, 31175),
        "2022-07-01": (16660, 990, 46360), "2024-06-01": (20140, 1240, 57340),
        "2026-07-01": (24210, 1500, 69210)
    },
    "GD N/Sub": {
        "1994-06-01": (1725, 116, 3465), "2001-12-01": (2590, 175, 7840),
        "2005-07-01_REV": (2980, 200, 8980), "2007-07-01": (3430, 230, 10330),
        "2008-07-01": (4115, 275, 12365), "2011-07-01": (6600, 120, 20100),
        "2015-07-01": (8540, 585, 26090), "2016-07-01": (10310, 720, 32110),
        "2017-07-01": (12570, 860, 38370), "2022-07-01": (18650, 1280, 57050),
        "2024-06-01": (26640, 1790, 80340), "2026-07-01": (32030, 2160, 96830)
    },
    "Sub": {
        "1994-06-01": (2338, 144, 4110), "2001-12-01": (3360, 215, 8950),
        "2005-07-01_REV": (3865, 245, 10235), "2007-07-01": (4445, 280, 11725),
        "2008-07-01": (5335, 335, 14045), "2011-07-01": (8600, 550, 25100),
        "2015-07-01": (11120, 710, 32420), "2016-07-01": (13680, 880, 40080),
        "2017-07-01": (16350, 1630, 47850), "2022-07-01": (24260, 1560, 71060),
        "2024-06-01": (30520, 2270, 98620), "2026-07-01": (36690, 2730, 118590)
    },
    "SM": {
        "1994-06-01": (3800, 197, 5376), "2001-12-01": (3700, 295, 10420),
        "2005-07-01_REV": (6555, 340, 11995), "2007-07-01": (7540, 390, 13780),
        "2008-07-01": (9050, 470, 16570), "2011-07-01": (14600, 770, 37700),
        "2015-07-01": (18880, 995, 48730), "2016-07-01": (23220, 1230, 60120),
        "2017-07-01": (27745, 1465, 71695), "2022-07-01": (41170, 2180, 106570),
        "2024-06-01": (46610, 2580, 124010), "2026-07-01": (56030, 3110, 149330)
    }
}

DEFAULT_CIVIL_SCALES = {
    "BPS-1": {"2011-07-01": (4800, 150, 9300), "2015-07-01": (6210, 195, 12060), "2016-07-01": (7610, 240, 14840), "2017-07-01": (9125, 285, 17675), "2022-07-01": (13540, 430, 26440)},
    "BPS-5": {"2011-07-01": (5400, 270, 13500), "2015-07-01": (6985, 350, 17485), "2016-07-01": (8590, 430, 21490), "2017-07-01": (10260, 510, 25560), "2022-07-01": (15230, 760, 38030)},
    "BPS-7": {"2011-07-01": (5700, 310, 15000), "2015-07-01": (7370, 405, 19520), "2016-07-01": (9070, 500, 24070), "2017-07-01": (10835, 595, 28665), "2022-07-01": (16080, 890, 42780)},
    "BPS-11": {"2011-07-01": (6600, 460, 20400), "2015-07-01": (8540, 595, 26390), "2016-07-01": (10310, 740, 32510), "2017-07-01": (12570, 880, 38970), "2022-07-01": (18650, 1310, 57950)},
    "BPS-14": {"2011-07-01": (8000, 610, 26300), "2015-07-01": (10340, 790, 34040), "2016-07-01": (12500, 980, 41900), "2017-07-01": (15180, 1170, 50280), "2022-07-01": (22530, 1740, 74730)},
    "BPS-16": {"2011-07-01": (10000, 800, 34000), "2015-07-01": (12880, 1030, 43780), "2016-07-01": (15600, 1280, 54000), "2017-07-01": (18910, 1520, 64510), "2022-07-01": (28070, 2260, 95870)},
    "BPS-17": {"2011-07-01": (16000, 1200, 40000), "2015-07-01": (20680, 1550, 51680), "2016-07-01": (25440, 1930, 64040), "2017-07-01": (30370, 2300, 76370), "2022-07-01": (45070, 3420, 113470)}
}

DEFAULT_RECT_RATES = {
    "Matric": [
        ("1994-06-01", 1230), ("2001-12-01", 1845), ("2005-07-01", 2128),
        ("2007-07-01", 2440), ("2008-07-01", 2930), ("2011-07-01", 4700),
        ("2015-07-01", 6080), ("2016-07-01", 7480), ("2017-07-01", 8945),
        ("2022-07-01", 13280)
    ],
    "Non-Matric": [
        ("1994-06-01", 1110), ("2001-12-01", 1665), ("2005-07-01", 1915),
        ("2007-07-01", 2205), ("2008-07-01", 2645), ("2011-07-01", 4300),
        ("2015-07-01", 5560), ("2016-07-01", 6840), ("2017-07-01", 8165),
        ("2022-07-01", 12120)
    ]
}

def load_registry():
    if os.path.exists(REGISTRY_FILE):
        try:
            with open(REGISTRY_FILE, "r") as f:
                return json.load(f)
        except Exception:
            pass
    # Initialize Registry
    reg = {
        "CAF": {
            "active": {
                "filename": "Bundled_Verified_CAF_Chart.xlsx",
                "upload_date": "2026-10-29 10:00:00",
                "approval_date": "2026-10-29 10:00:00",
                "effective_years": [1994, 2001, 2005, 2007, 2008, 2011, 2015, 2016, 2017, 2022, 2024, 2026],
                "ranks": list(BOOTSTRAP_CAF_SCALES.keys()),
                "data_file": "caf_active.json"
            },
            "backup": None
        },
        "Civil Servants": {
            "active": {
                "filename": "Bundled_Verified_BPS_Chart.xlsx",
                "upload_date": "2026-10-29 10:00:00",
                "approval_date": "2026-10-29 10:00:00",
                "effective_years": [2011, 2015, 2016, 2017, 2022],
                "ranks": list(DEFAULT_CIVIL_SCALES.keys()),
                "data_file": "civil_active.json"
            },
            "backup": None
        }
    }
    save_registry(reg)
    save_category_data("CAF", BOOTSTRAP_CAF_SCALES, DEFAULT_RECT_RATES, "caf_active.json")
    save_category_data("Civil Servants", DEFAULT_CIVIL_SCALES, {}, "civil_active.json")
    return reg

def save_registry(reg):
    with open(REGISTRY_FILE, "w") as f:
        json.dump(reg, f, indent=2)

def save_category_data(category, scales, rect_rates, filename):
    path = os.path.join(DATA_DIR, filename)
    with open(path, "w") as f:
        json.dump({"scales": scales, "rect_rates": rect_rates}, f, indent=2)

def load_category_data(category, filename):
    path = os.path.join(DATA_DIR, filename)
    if os.path.exists(path):
        try:
            with open(path, "r") as f:
                d = json.load(f)
                return d.get("scales", {}), d.get("rect_rates", {})
        except Exception:
            pass
    if category == "Civil Servants":
        return DEFAULT_CIVIL_SCALES, {}
    return BOOTSTRAP_CAF_SCALES, DEFAULT_RECT_RATES

# ==============================================================================
# 1. PARSER ENGINE FOR DYNAMIC UPLOADED CHARTS
# ==============================================================================
def format_scale_str(sc):
    return f"{sc[0]}-{sc[1]}-{sc[2]}" if sc else ""

def parse_uploaded_chart(uploaded_file):
    """
    Parses both Wide Matrix charts and Row-based (Tabular) charts.
    Extracts Ranks, Scale Brackets (Min-Inc-Max), Effective Years, and Revision dates.
    """
    try:
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)

        df = df.dropna(how='all')
        if df.empty:
            return None, "The uploaded file is empty."

        scales = {}
        rect_rates = {"Matric": [], "Non-Matric": []}
        detected_years = set()
        warnings = []

        date_pat = re.compile(r'(\d{1,2}[-/]\d{1,2}[-/]\d{2,4})')
        year_pat = re.compile(r'\b(19\d\d|20\d\d)\b')
        scale_pat = re.compile(r'(\d+)\s*[-/]\s*(\d+)\s*[-/]\s*(\d+)')
        num_pat = re.compile(r'^\s*(\d+)\s*$')

        rank_col = None
        for col in df.columns:
            if "rank" in str(col).lower() or "designation" in str(col).lower() or "bps" in str(col).lower():
                rank_col = col
                break
        if rank_col is None:
            rank_col = df.columns[0]

        eff_year_col = None
        for col in df.columns:
            if "effective" in str(col).lower() and "year" in str(col).lower():
                eff_year_col = col
                break

        # Check if Tabular / Long format (Rows contain Rank, Effective Year, Scale)
        scale_col = None
        for col in df.columns:
            if "scale" in str(col).lower() or "pay band" in str(col).lower():
                scale_col = col
                break

        if eff_year_col and scale_col:
            # TABULAR / LONG FORMAT
            for _, row in df.iterrows():
                rk = str(row[rank_col]).strip()
                if not rk or rk.lower() == 'nan':
                    continue
                raw_y = str(row[eff_year_col]).strip()
                m_y = year_pat.search(raw_y)
                if not m_y:
                    continue
                year_int = int(m_y.group(1))
                detected_years.add(year_int)

                # Look for explicit date in row
                eff_date = f"{year_int}-07-01"
                for c in df.columns:
                    if "date" in str(c).lower():
                        m_d = date_pat.search(str(row[c]))
                        if m_d:
                            eff_date = normalize_date_str(m_d.group(1))

                val = str(row[scale_col]).strip()
                m_sc = scale_pat.search(val)
                if m_sc:
                    min_p, inc_p, max_p = map(int, m_sc.groups())
                    if rk not in scales:
                        scales[rk] = {}
                    scales[rk][eff_date] = (min_p, inc_p, max_p)
        else:
            # WIDE MATRIX FORMAT (Standard CAF / Ilyas Sb layout)
            col_date_map = {}
            for col in df.columns:
                col_str = str(col).strip()
                m_d = date_pat.search(col_str)
                m_y = year_pat.search(col_str)
                if m_d:
                    iso_dt = normalize_date_str(m_d.group(1))
                    if "upgradation" in col_str.lower():
                        iso_dt += "_UPG"
                    elif "revision" in col_str.lower() and iso_dt == "2005-07-01":
                        iso_dt += "_REV"
                    col_date_map[col] = iso_dt
                    if m_y:
                        detected_years.add(int(m_y.group(1)))
                elif m_y:
                    y_int = int(m_y.group(1))
                    col_date_map[col] = f"{y_int}-07-01"
                    detected_years.add(y_int)

            for _, row in df.iterrows():
                rk = str(row[rank_col]).strip()
                if not rk or rk.lower() == 'nan':
                    continue

                # Recruit rates detection
                if "rect" in rk.lower():
                    cat = "Non-Matric" if "non" in rk.lower() else "Matric"
                    for col, dt in col_date_map.items():
                        v = str(row[col]).strip()
                        m_n = num_pat.search(v)
                        if m_n:
                            amt = int(m_n.group(1))
                            if amt > 0:
                                pure_dt = dt.split("_")[0]
                                rect_rates[cat].append((pure_dt, amt))
                    continue

                for col, dt in col_date_map.items():
                    val = str(row[col]).strip()
                    m_sc = scale_pat.search(val)
                    if m_sc:
                        min_p, inc_p, max_p = map(int, m_sc.groups())
                        if rk not in scales:
                            scales[rk] = {}
                        scales[rk][dt] = (min_p, inc_p, max_p)

        if not scales:
            return None, "No valid scale entries found (Format must be 'Min-Inc-Max' e.g. 5400-270-13500)."

        # Sort recruit rates
        for cat in rect_rates:
            rect_rates[cat].sort(key=lambda x: x[0])

        return {
            "scales": scales,
            "rect_rates": rect_rates,
            "ranks": sorted(list(scales.keys())),
            "effective_years": sorted(list(detected_years)),
            "warnings": warnings,
            "rows_read": len(df)
        }, None
    except Exception as e:
        return None, f"Failed to interpret chart: {str(e)}"

def normalize_date_str(raw_d):
    raw_d = raw_d.replace('/', '-')
    parts = raw_d.split('-')
    if len(parts[0]) == 4:
        return f"{parts[0]}-{int(parts[1]):02d}-{int(parts[2]):02d}"
    return f"{int(parts[2]):04d}-{int(parts[1]):02d}-{int(parts[0]):02d}"

# ==============================================================================
# 2. RUNTIME CALCULATION ENGINE
# ==============================================================================
def get_rect_pay(qualification, enrol_date_str, rect_rates_dict):
    rates = rect_rates_dict.get(qualification, rect_rates_dict.get("Matric", []))
    applicable_pay = None
    for rev_dt, rate in rates:
        if rev_dt <= enrol_date_str:
            applicable_pay = rate
    return applicable_pay

def get_scale_at_date(rank, target_date_str, scales_dict, rev_tag=None):
    if rank not in scales_dict:
        clean = rank.replace("GD ", "")
        for k in scales_dict:
            if clean.lower() == k.lower() or rank.lower() == k.lower():
                rank = k
                break
    if rank not in scales_dict:
        return None
    r_scales = scales_dict[rank]
    if rev_tag and rev_tag in r_scales:
        return r_scales[rev_tag]
    target_dt = datetime.datetime.strptime(target_date_str, "%Y-%m-%d").date()
    applicable = None
    for k, sc in r_scales.items():
        dt_part = k.split("_")[0]
        if datetime.datetime.strptime(dt_part, "%Y-%m-%d").date() <= target_dt:
            applicable = sc
    return applicable

def run_pay_fixation_engine(personnel, scales_dict, rect_rates, special_instructions=""):
    doe = personnel["doe"]
    enrol_rank = personnel.get("enrol_rank", "GD Sep")
    qualification = personnel.get("qualification", "Matric")
    d_remust = personnel.get("d_remust")
    promotions = personnel.get("promotions", [])
    end_date_str = personnel.get("end_date", "2026-07-01")
    warnings = []

    # Detect dynamic revision schedule from active chart scales
    active_revisions = set()
    for rk, dt_dict in scales_dict.items():
        for d_key in dt_dict.keys():
            active_revisions.add(d_key)

    events = []
    is_ncs_e = (enrol_rank == "NCs E")

    if is_ncs_e:
        events.append({"date": doe, "type": "ENROLLMENT", "rank": "NCs E", "remarks": "Enrolled directly as NCs E"})
    else:
        events.append({"date": doe, "type": "ENROLLMENT", "rank": "GD Rect", "remarks": f"Enrolled as Rect ({qualification})"})
        if d_remust:
            events.append({"date": d_remust, "type": "REMUSTERING", "rank": enrol_rank, "remarks": "Remustered as train sldr"})

    for p in promotions:
        p_rank = p["rank"]
        events.append({"date": p["date"], "type": "PROMOTION", "rank": p_rank, "remarks": f"Pmtd {p_rank.replace('GD ', '')}"})

    for r_key in sorted(list(active_revisions)):
        r_dt = r_key.split("_")[0]
        r_type = "UPGRADATION" if "UPG" in r_key else "REVISION"
        if doe <= r_dt <= end_date_str:
            events.append({"date": r_dt, "type": r_type, "rev_key": r_key, "remarks": r_type.capitalize()})

    start_year = datetime.datetime.strptime(doe, "%Y-%m-%d").year
    end_year = datetime.datetime.strptime(end_date_str, "%Y-%m-%d").year
    for y in range(start_year, end_year + 1):
        inc_dt = f"{y}-12-01"
        if doe <= inc_dt <= end_date_str:
            events.append({"date": inc_dt, "type": "ANNUAL_INCREMENT", "remarks": "A/Inc"})

    priority = {"ENROLLMENT": 0, "REMUSTERING": 1, "REVISION": 2, "UPGRADATION": 3, "PROMOTION": 4, "ANNUAL_INCREMENT": 5}
    events.sort(key=lambda x: (x["date"], priority.get(x["type"], 9)))

    ledger = []
    current_pay = 0
    current_rank = enrol_rank if is_ncs_e else "GD Sep"
    current_scale = None
    substantive_scale = None
    last_change_dt = None

    for ev in events:
        ev_dt = datetime.datetime.strptime(ev["date"], "%Y-%m-%d").date()
        ev_t = ev["type"]

        if ev_t == "ENROLLMENT":
            if is_ncs_e:
                current_scale = get_scale_at_date("NCs E", ev["date"], scales_dict)
                substantive_scale = current_scale
                if not current_scale:
                    warnings.append(f"HOLD — REVIEW REQUIRED: NCs E scale unavailable on {ev['date']} in active chart.")
                    current_pay = "Review Req"
                else:
                    current_pay = current_scale[0]
                last_change_dt = ev_dt
                ledger.append({
                    "date": ev_dt.strftime("%d-%m-%Y"), "pay": current_pay,
                    "scale": format_scale_str(current_scale), "remarks": ev["remarks"]
                })
            else:
                rect_pay = get_rect_pay(qualification, ev["date"], rect_rates)
                current_scale = get_scale_at_date(enrol_rank, ev["date"], scales_dict)
                substantive_scale = current_scale
                if rect_pay is None:
                    warnings.append(f"HOLD — REVIEW REQUIRED: Recruit rate ({qualification}) missing for date {ev['date']}.")
                    current_pay = "Review Req"
                else:
                    current_pay = rect_pay
                last_change_dt = ev_dt
                ledger.append({
                    "date": ev_dt.strftime("%d-%m-%Y"), "pay": current_pay,
                    "scale": "Fixed", "remarks": ev["remarks"]
                })

        elif ev_t == "REMUSTERING":
            current_rank = enrol_rank
            current_scale = get_scale_at_date(current_rank, ev["date"], scales_dict)
            substantive_scale = current_scale
            if not current_scale:
                warnings.append(f"HOLD — REVIEW REQUIRED: Scale missing for {current_rank} on {ev['date']}.")
                current_pay = "Review Req"
            else:
                current_pay = current_scale[0]
            last_change_dt = ev_dt
            ledger.append({
                "date": ev_dt.strftime("%d-%m-%Y"), "pay": current_pay,
                "scale": format_scale_str(current_scale), "remarks": ev["remarks"]
            })

        elif ev_t == "ANNUAL_INCREMENT":
            if not isinstance(current_pay, (int, float)):
                ledger.append({"date": ev_dt.strftime("%d-%m-%Y"), "pay": current_pay, "scale": '"', "remarks": "HOLD — Scale Unresolved"})
                continue
            days_served = (ev_dt - last_change_dt).days if last_change_dt else 365
            if days_served < 180:
                ledger.append({"date": ev_dt.strftime("%d-%m-%Y"), "pay": current_pay, "scale": '"', "remarks": "Not entitled (< 6 mos in scale)"})
            else:
                active_inc = substantive_scale[1] if substantive_scale else (current_scale[1] if current_scale else 0)
                current_pay += active_inc
                ledger.append({"date": ev_dt.strftime("%d-%m-%Y"), "pay": current_pay, "scale": '"', "remarks": "A/Inc"})

        elif ev_t == "REVISION":
            active_rank = current_rank
            if active_rank in ["H/Lt", "H/Capt"]:
                active_rank = "SM" if "SM" in scales_dict else "Sub"
            new_sc = get_scale_at_date(active_rank, ev["date"], scales_dict, rev_tag=ev.get("rev_key"))
            if new_sc and current_scale and isinstance(current_pay, (int, float)):
                stages = (current_pay - current_scale[0]) // current_scale[1]
                current_pay = new_sc[0] + (stages * new_sc[1])
                current_scale = new_sc
                substantive_scale = new_sc
                ledger.append({
                    "date": ev_dt.strftime("%d-%m-%Y"), "pay": current_pay,
                    "scale": format_scale_str(current_scale), "remarks": "Revision"
                })
            else:
                if not new_sc:
                    warnings.append(f"HOLD — REVIEW REQUIRED: Revision scale missing for {active_rank} on {ev['date']}.")
                ledger.append({
                    "date": ev_dt.strftime("%d-%m-%Y"), "pay": current_pay,
                    "scale": format_scale_str(new_sc) if new_sc else "Missing Scale", "remarks": "Revision (Review Req)"
                })

        elif ev_t == "UPGRADATION":
            active_rank = current_rank
            new_sc = get_scale_at_date(active_rank, ev["date"], scales_dict, rev_tag=ev.get("rev_key"))
            if new_sc and isinstance(current_pay, (int, float)):
                p_step = new_sc[0]
                while p_step < current_pay:
                    p_step += new_sc[1]
                current_pay = p_step
                current_scale = new_sc
                substantive_scale = new_sc
                ledger.append({
                    "date": ev_dt.strftime("%d-%m-%Y"), "pay": current_pay,
                    "scale": format_scale_str(current_scale), "remarks": "Upgradation"
                })

        elif ev_t == "PROMOTION":
            new_rank = ev["rank"]
            if new_rank in ["H/Lt", "H/Capt"]:
                sm_scale = get_scale_at_date("SM", ev["date"], scales_dict)
                if not sm_scale:
                    warnings.append(f"HOLD — REVIEW REQUIRED: SM Scale unavailable on {ev['date']} for {new_rank} advance increment.")
                    honorary_inc = 0
                else:
                    honorary_inc = sm_scale[1]
                if isinstance(current_pay, (int, float)):
                    current_pay += honorary_inc
                last_change_dt = ev_dt
                ledger.append({
                    "date": ev_dt.strftime("%d-%m-%Y"), "pay": current_pay,
                    "scale": format_scale_str(substantive_scale) if substantive_scale else '"',
                    "remarks": f"1 x adv inc on prom as {new_rank} (SM scale basis)"
                })
            else:
                current_rank = new_rank
                new_sc = get_scale_at_date(current_rank, ev["date"], scales_dict)
                if not new_sc:
                    warnings.append(f"HOLD — REVIEW REQUIRED: Missing scale for rank '{current_rank}' on {ev['date']}.")
                    current_pay = "Review Req"
                else:
                    if isinstance(current_pay, (int, float)):
                        p_step = new_sc[0]
                        while p_step < current_pay:
                            p_step += new_sc[1]
                        current_pay = p_step + new_sc[1]
                    current_scale = new_sc
                    substantive_scale = new_sc
                last_change_dt = ev_dt
                ledger.append({
                    "date": ev_dt.strftime("%d-%m-%Y"), "pay": current_pay,
                    "scale": format_scale_str(current_scale), "remarks": ev["remarks"]
                })

    return {
        "header": personnel, "ledger": ledger,
        "special_instructions": special_instructions, "warnings": warnings
    }

# ==============================================================================
# 3. MULTI-FORMAT EXPORT GENERATORS
# ==============================================================================
def generate_excel_bytes(data):
    wb = Workbook()
    ws = wb.active
    ws.title = "PAY FIXATION"
    ws.views.sheetView[0].showGridLines = True

    h = data["header"]
    ws.merge_cells("A1:E1")
    ws["A1"] = f"PAY FIXATION IRO FC NO. {h.get('fc_no','')} {h.get('rank','')} {h.get('name','').upper()}"
    ws["A1"].font = Font(name="Calibri", size=11, bold=True)
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")

    bio = [
        ("a", "Category & Active Chart", f"{h.get('category','CAF')} ({h.get('chart_name','')})"),
        ("b", "D.O E", h.get("doe")),
        ("c", "D.O.Remustration", h.get("d_remust", "N/A (Direct)"))
    ]
    letters = ["d", "e", "f", "g", "h", "i"]
    for idx, p in enumerate(h.get("promotions", [])):
        bio.append((letters[idx], f"D.O.Promotion as {p['rank'].replace('GD ', '')}", p["date"]))

    r = 3
    for mark, lbl, val in bio:
        ws.cell(row=r, column=1, value=mark).alignment = Alignment(horizontal="center")
        ws.cell(row=r, column=2, value=lbl)
        ws.cell(row=r, column=3, value=val).alignment = Alignment(horizontal="center")
        r += 1

    if data.get("special_instructions"):
        r += 1
        ws.cell(row=r, column=1, value="Note:")
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=5)
        ws.cell(row=r, column=2, value=f"Special Instructions: {data['special_instructions']} (Verification Required)").font = Font(italic=True, color="000080")

    r += 2
    cols = ["Ser", "Date", "Pay Fixed", "Pay Scale", "Remarks"]
    for ci, col in enumerate(cols, start=1):
        cell = ws.cell(row=r, column=ci, value=col)
        cell.font = Font(bold=True)
        cell.fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
        cell.alignment = Alignment(horizontal="center")

    border = Border(left=Side(style='thin', color='A6A6A6'), right=Side(style='thin', color='A6A6A6'),
                    top=Side(style='thin', color='A6A6A6'), bottom=Side(style='thin', color='A6A6A6'))
    r += 1
    for idx, row in enumerate(data["ledger"], start=1):
        ws.cell(row=r, column=1, value=idx).alignment = Alignment(horizontal="center")
        ws.cell(row=r, column=2, value=row["date"]).alignment = Alignment(horizontal="center")
        ws.cell(row=r, column=3, value=row["pay"]).alignment = Alignment(horizontal="center")
        ws.cell(row=r, column=4, value=row["scale"]).alignment = Alignment(horizontal="center")
        ws.cell(row=r, column=5, value=row["remarks"])
        for c in range(1, 6):
            ws.cell(row=r, column=c).border = border
        r += 1

    ws.column_dimensions['A'].width = 8
    ws.column_dimensions['B'].width = 14
    ws.column_dimensions['C'].width = 14
    ws.column_dimensions['D'].width = 24
    ws.column_dimensions['E'].width = 36

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

def generate_word_bytes(data):
    doc = Document()
    h = data["header"]
    p_title = doc.add_paragraph()
    run = p_title.add_run(f"PAY FIXATION IRO FC NO. {h.get('fc_no','')} {h.get('rank','')} {h.get('name','').upper()}")
    run.font.bold = True
    run.font.size = Pt(12)
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    bio_table = doc.add_table(rows=0, cols=3)
    bio_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    def add_bio(mark, lbl, val):
        r = bio_table.add_row().cells
        r[0].text = mark
        r[1].text = lbl
        r[2].text = str(val) if val else "N/A"
        r[0].width = Inches(0.5)
        r[1].width = Inches(2.6)
        r[2].width = Inches(2.2)

    add_bio("a.", "Chart Category", f"{h.get('category','CAF')} [{h.get('chart_name','')}]")
    add_bio("b.", "D.O E", h.get("doe"))
    add_bio("c.", "D.O.Remustration", h.get("d_remust", "N/A"))
    letters = ["d.", "e.", "f.", "g.", "h.", "i."]
    for idx, p in enumerate(h.get("promotions", [])):
        add_bio(letters[idx], f"D.O.Promotion as {p['rank'].replace('GD ', '')}", p["date"])

    if data.get("special_instructions"):
        p_note = doc.add_paragraph()
        r_note = p_note.add_run(f"\nSpecial Instructions: {data['special_instructions']} (Cross-check against service book)")
        r_note.font.italic = True
        r_note.font.color.rgb = RGBColor(0, 32, 96)

    doc.add_paragraph().text = ""
    table = doc.add_table(rows=1, cols=5)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = 'Table Grid'
    hdr_cells = table.rows[0].cells
    cols = ["Ser", "Date", "Pay Fixed", "Pay Scale", "Remarks"]
    for i, title in enumerate(cols):
        hdr_cells[i].text = title
        hdr_cells[i].paragraphs[0].runs[0].font.bold = True
        shading = parse_xml(r'<w:shd {} w:fill="D9E1F2"/>'.format(nsdecls('w')))
        hdr_cells[i]._tc.get_or_add_tcPr().append(shading)

    for idx, row in enumerate(data["ledger"], start=1):
        row_cells = table.add_row().cells
        row_cells[0].text = str(idx)
        row_cells[1].text = str(row["date"])
        row_cells[2].text = str(row["pay"])
        row_cells[3].text = str(row["scale"])
        row_cells[4].text = str(row["remarks"])
        for c in row_cells:
            c.paragraphs[0].runs[0].font.size = Pt(9.5)

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()

def generate_pdf_bytes(data):
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=letter, leftMargin=36, rightMargin=36, topMargin=36, bottomMargin=36)
    styles = getSampleStyleSheet()

    elements = []
    if os.path.exists("logo.png"):
        elements.append(RLImage("logo.png", width=65, height=65))
        elements.append(Spacer(1, 4))

    h = data["header"]
    title_text = f"<b>PAY FIXATION IRO FC NO. {h.get('fc_no','')} {h.get('rank','')} {h.get('name','').upper()}</b>"
    elements.append(Paragraph(title_text, styles['Title']))
    elements.append(Spacer(1, 8))

    bio_data = [
        ["a.", "Chart Category", f"{h.get('category','CAF')} [{h.get('chart_name','')}]"],
        ["b.", "D.O E", str(h.get("doe"))],
        ["c.", "D.O.Remustration", str(h.get("d_remust", "N/A"))]
    ]
    letters = ["d.", "e.", "f.", "g.", "h.", "i."]
    for idx, p in enumerate(h.get("promotions", [])):
        bio_data.append([letters[idx], f"D.O.Promotion as {p['rank'].replace('GD ', '')}", str(p["date"])])

    t_bio = Table(bio_data, colWidths=[30, 210, 160])
    t_bio.setStyle(TableStyle([
        ('FONTNAME', (0,0), (-1,-1), 'Helvetica'),
        ('FONTSIZE', (0,0), (-1,-1), 8.5),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
    ]))
    elements.append(t_bio)

    if data.get("special_instructions"):
        elements.append(Spacer(1, 6))
        elements.append(Paragraph(f"<i>Special Instructions: {data['special_instructions']} (Verification Required)</i>", styles['Normal']))

    elements.append(Spacer(1, 10))

    tbl_data = [["Ser", "Date", "Pay Fixed", "Pay Scale", "Remarks"]]
    for idx, row in enumerate(data["ledger"], start=1):
        tbl_data.append([str(idx), str(row["date"]), str(row["pay"]), str(row["scale"]), str(row["remarks"])])

    t_main = Table(tbl_data, colWidths=[30, 75, 75, 130, 230])
    t_main.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#D9E1F2')),
        ('TEXTCOLOR', (0,0), (-1,0), colors.black),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('FONTSIZE', (0,0), (-1,0), 9),
        ('FONTNAME', (0,1), (-1,-1), 'Helvetica'),
        ('FONTSIZE', (0,1), (-1,-1), 8),
        ('ALIGN', (0,0), (3,-1), 'CENTER'),
        ('ALIGN', (4,0), (4,-1), 'LEFT'),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#A6A6A6')),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0), (-1,-1), 2.5),
        ('BOTTOMPADDING', (0,0), (-1,-1), 2.5),
    ]))
    elements.append(t_main)
    doc.build(elements)
    return buf.getvalue()

def create_registry_zip():
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for fname in os.listdir(DATA_DIR):
            fpath = os.path.join(DATA_DIR, fname)
            if os.path.isfile(fpath):
                z.write(fpath, arcname=fname)
    return buf.getvalue()

# ==============================================================================
# 4. STREAMLIT APPLICATION INTERFACE
# ==============================================================================
def main():
    st.set_page_config(page_title="Pay Fixation Agent", layout="wide", page_icon="logo.png")

    # Header with exact logo.png
    col_logo, col_title = st.columns([1, 6])
    with col_logo:
        if os.path.exists("logo.png"):
            st.image("logo.png", width=95)
    with col_title:
        st.title("Pay Fixation Agent")
        st.caption("Developed by **SC Rehman Gul** to assist Civil Armed Forces (CAF) & Civil Servants in automated pay fixation.")

    st.info("⚠️ **Audit Disclaimer:** This tool processes official scales dynamically. Always cross-check results against official departmental orders and service records.")

    # Load persistent registry
    registry = load_registry()

    # --------------------------------------------------------------------------
    # SIDEBAR: DYNAMIC CHART MANAGEMENT & AUDIT VAULT
    # --------------------------------------------------------------------------
    st.sidebar.header("📁 Pay Chart Vault & Configuration")
    target_category = st.sidebar.selectbox("Select Pay Scale Category", ["CAF", "Civil Servants"], index=0)

    cat_meta = registry[target_category]["active"]
    st.sidebar.markdown(f"**Active Chart:** `{cat_meta['filename']}`")
    st.sidebar.markdown(f"**Approval Date:** `{cat_meta['approval_date']}`")
    st.sidebar.markdown(f"**Effective Years:** `{', '.join(map(str, cat_meta['effective_years']))}`")

    # Upload replacement chart
    st.sidebar.markdown("---")
    st.sidebar.subheader(f"Upload Replacement for {target_category}")
    uploaded_chart = st.sidebar.file_uploader(
        f"Upload New {target_category} Chart (.xlsx, .xls, .csv)",
        type=["xlsx", "xls", "csv"],
        key="chart_uploader"
    )

    if uploaded_chart is not None:
        parsed_res, err = parse_uploaded_chart(uploaded_chart)
        if err:
            st.sidebar.error(f"❌ Validation Error: {err}")
        else:
            st.sidebar.success(f"✅ Readable: {parsed_res['rows_read']} rows, {len(parsed_res['ranks'])} ranks.")
            st.sidebar.info(f"Years Detected: {parsed_res['effective_years']}")

            with st.sidebar.expander("👁️ Preview Parsed Ranks"):
                st.write(parsed_res["ranks"])

            if st.sidebar.button(f"Confirm & Activate New {target_category} Chart", type="primary"):
                # Move current active to backup
                registry[target_category]["backup"] = cat_meta
                # Save new file
                new_data_file = f"{target_category.lower().replace(' ', '_')}_active.json"
                save_category_data(target_category, parsed_res["scales"], parsed_res["rect_rates"], new_data_file)

                # Update registry
                registry[target_category]["active"] = {
                    "filename": uploaded_chart.name,
                    "upload_date": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "approval_date": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "effective_years": parsed_res["effective_years"],
                    "ranks": parsed_res["ranks"],
                    "data_file": new_data_file
                }
                save_registry(registry)
                st.sidebar.success(f"🎉 New {target_category} chart successfully activated!")
                st.rerun()

    # Rollback support
    if registry[target_category].get("backup"):
        bak = registry[target_category]["backup"]
        if st.sidebar.button(f"⏪ Rollback to Previous ({bak['filename']})"):
            registry[target_category]["active"] = bak
            registry[target_category]["backup"] = None
            save_registry(registry)
            st.sidebar.warning(f"Reverted to previous chart: {bak['filename']}")
            st.rerun()

    # Persistent Storage Download
    st.sidebar.markdown("---")
    st.sidebar.caption("💾 **Cloud Deployment Persistence:**")
    st.sidebar.download_button(
        label="Download Registry Backup (.zip)",
        data=create_registry_zip(),
        file_name="chart_registry_backup.zip",
        mime="application/zip",
        help="Commit this zip's extracted data/ folder to your GitHub repo to ensure permanent zero-loss persistence on Streamlit Cloud."
    )

    # --------------------------------------------------------------------------
    # MAIN FORM: PERSONNEL SERVICE HISTORY
    # --------------------------------------------------------------------------
    active_scales, active_rect_rates = load_category_data(target_category, cat_meta["data_file"])
    available_ranks = cat_meta["ranks"]
    if not available_ranks:
        available_ranks = ["GD Sep", "GD Lnk", "GD Nk", "GD Hav", "GD N/Sub", "Sub", "SM"]

    col1, col2 = st.columns(2)
    with col1:
        st.subheader("👤 Personnel Profile")
        fc_no = st.text_input("Service / FC / Personnel No.", value="102052")
        name = st.text_input("Full Name", value="Jamal")

        # Rank selector directly powered by active chart
        final_rank = st.selectbox("Substantive / Final Rank", available_ranks, index=min(5, len(available_ranks)-1))
        enrol_rank = st.selectbox("Enrolled Rank", available_ranks, index=0)

        qualification = "Matric"
        if target_category == "CAF":
            qualification = st.radio("Recruit Educational Category", ["Matric", "Non-Matric"], horizontal=True)

        doe = st.date_input("Date of Enrollment (D.O.E)",
                            value=datetime.date(2007, 3, 15),
                            min_value=datetime.date(1990, 1, 1),
                            max_value=datetime.date(2035, 12, 31)).strftime("%Y-%m-%d")

        d_remust = None
        if target_category == "CAF" and enrol_rank != "NCs E":
            d_remust = st.date_input("Date of Remustering",
                                     value=datetime.date(2007, 9, 10),
                                     min_value=datetime.date(1990, 1, 1),
                                     max_value=datetime.date(2035, 12, 31)).strftime("%Y-%m-%d")

    with col2:
        st.subheader("📈 Promotion History")
        st.caption("Promotions will be checked against the active dynamic chart:")
        num_promotions = st.number_input("Number of Promotions to Add", min_value=0, max_value=8, value=4, step=1)
        promotions = []
        default_promos = [
            (datetime.date(2012, 3, 15), "GD Lnk"),
            (datetime.date(2014, 2, 20), "GD Nk"),
            (datetime.date(2017, 12, 2), "GD Hav"),
            (datetime.date(2026, 1, 8), "GD N/Sub"),
            (datetime.date(2028, 1, 1), "Sub"),
            (datetime.date(2030, 1, 1), "SM")
        ]
        for i in range(num_promotions):
            p_c1, p_c2 = st.columns([1, 1])
            def_dt, def_rk = default_promos[i] if i < len(default_promos) else (datetime.date(2020, 1, 1), available_ranks[0])
            with p_c1:
                p_dt = st.date_input(f"Promo {i+1} Date", value=def_dt,
                                     min_value=datetime.date(1990, 1, 1),
                                     max_value=datetime.date(2035, 12, 31),
                                     key=f"p_dt_{i}").strftime("%Y-%m-%d")
            with p_c2:
                # Fallback to available ranks
                cur_idx = available_ranks.index(def_rk) if def_rk in available_ranks else 0
                p_rk = st.selectbox(f"Promo {i+1} Rank", available_ranks + ["H/Lt", "H/Capt"], index=cur_idx, key=f"p_rk_{i}")
            promotions.append({"date": p_dt, "rank": p_rk})

    st.markdown("---")
    st.subheader("📝 Directives & Special Instructions")
    special_instructions = st.text_input(
        "Enter case-specific orders or notes (e.g., '1 x adv increment for FA/Matric', 'Withheld increment 2013', 'Pay protection order'):",
        placeholder="Type order reference here..."
    )

    if st.button("🚀 Calculate Pay Fixation Statement", type="primary"):
        personnel_payload = {
            "name": name,
            "fc_no": fc_no,
            "rank": final_rank,
            "enrol_rank": enrol_rank,
            "qualification": qualification,
            "doe": doe,
            "d_remust": d_remust,
            "promotions": promotions,
            "end_date": "2026-07-01",
            "category": target_category,
            "chart_name": cat_meta["filename"]
        }

        # Calculate using the active dynamic chart
        result = run_pay_fixation_engine(personnel_payload, active_scales, active_rect_rates, special_instructions)

        if result.get("warnings"):
            for w in result["warnings"]:
                st.warning(f"⚠️ {w}")

        st.success(f"✅ Pay fixation calculated using approved {target_category} chart `{cat_meta['filename']}`!")
        df_ledger = pd.DataFrame(result["ledger"])
        df_ledger.index = df_ledger.index + 1
        st.dataframe(df_ledger, use_container_width=True)

        st.markdown("### 📥 Download Formal Statements")
        dcol1, dcol2, dcol3 = st.columns(3)

        dcol1.download_button(
            label="📊 Download Excel (.xlsx)",
            data=generate_excel_bytes(result),
            file_name=f"Pay_Fixation_{fc_no}_{name}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        dcol2.download_button(
            label="📄 Download Word (.docx)",
            data=generate_word_bytes(result),
            file_name=f"Pay_Fixation_{fc_no}_{name}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
        dcol3.download_button(
            label="📑 Download PDF (.pdf)",
            data=generate_pdf_bytes(result),
            file_name=f"Pay_Fixation_{fc_no}_{name}.pdf",
            mime="application/pdf"
        )

        st.markdown("---")
        # Direct Access & QR Share Card
        share_col_logo, share_col_info, share_col_qr = st.columns([1, 3, 2])
        with share_col_logo:
            if os.path.exists("logo.png"):
                st.image("logo.png", width=80)
        with share_col_info:
            st.markdown("**Pay Fixation Agent — Direct Access & Share**")
            st.markdown("Share this application link with CAF & Civil Service personnel to audit pay fixation statements.")
            app_url = st.text_input("App / Repository Direct Link:", value="https://github.com", key="app_share_url")
            st.markdown(f"🔗 **Direct Link:** [{app_url}]({app_url})")
        with share_col_qr:
            qr_api_url = f"https://api.qrserver.com/v1/create-qr-code/?size=120x120&data={app_url}"
            st.image(qr_api_url, caption="Scan QR to open app", width=120)

if __name__ == "__main__":
    main()
