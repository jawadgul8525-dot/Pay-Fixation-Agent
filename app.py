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
# 0. PERSISTENCE & STORAGE CONFIGURATION
# ==============================================================================
DATA_DIR = os.path.join(os.path.dirname(__file__), "data")
REGISTRY_FILE = os.path.join(DATA_DIR, "chart_registry.json")
os.makedirs(DATA_DIR, exist_ok=True)

# ------------------------------------------------------------------------------
# Verified CAF Uniformed Scales (1994 – 2026)
# ------------------------------------------------------------------------------
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

# ------------------------------------------------------------------------------
# Official Federal Government Civil Servants Scales (1972 – 2026: BPS-01 to 19)
# ------------------------------------------------------------------------------
RAW_CIVIL_SERVANT_DATA = """
BPS / Rank,01-03-1972,01-05-1977,01-07-1981,01-07-1983,01-07-1987,01-06-1991,01-06-1991_Rev,1994-06-01,01-12-2001,01-07-2005,01-07-2007,01-07-2008,01-07-2011,01-07-2015,01-07-2016,01-07-2017,01-07-2022,01-07-2026
BPS-01,100-2-140,250-5-340,250-5-375,440-10-640,600-13-860,605-17-860,920-26-1310,1245-35-1770,1870-55-3520,2150-65-4100,2475-75-4725,2970-90-5670,4800-150-9300,6210-195-12060,7640-240-14840,9130-290-17830,13550-430-26450,16280-520-31880
BPS-02,110-3-160,260-6-365,260-6-405,460-12-700,625-16-945,630-21-945,945-32-1425,1275-44-1935,1915-65-3865,2200-75-4450,2530-85-5080,3035-100-6035,4900-170-10000,6335-220-12935,7790-275-16040,9310-330-19210,13820-490-28520,16600-590-34300
BPS-03,120-3-180,270-7-390,270-7-435,480-14-760,650-19-1030,655-25-1030,975-37-1530,1320-50-2070,1980-75-4230,2275-85-4825,2615-100-5615,3140-120-6740,5050-200-11050,6535-260-14335,8040-325-17790,9610-390-21310,14260-580-31660,17130-700-38130
BPS-04,130-4-200,280-8-415,280-8-475,500-16-820,675-22-1115,680-29-1115,1005-43-1650,1360-58-2230,2040-85-4590,2345-100-5345,2700-115-6150,3240-140-7440,5200-230-12100,6730-300-15730,8280-370-19380,9900-440-23100,14690-660-34490,17650-800-41650
BPS-05,150-6-280,290-10-470,290-10-540,520-18-880,700-25-1200,705-33-1200,1035-49-1770,1400-66-2390,2100-100-5100,2415-115-5865,2780-135-6830,3340-160-8140,5400-260-13200,6985-340-17185,8590-420-21190,10260-500-25260,15230-750-37730,18300-910-45600
BPS-06,165-8-315,315-12-525,315-12-605,540-20-940,725-28-1285,730-37-1285,1065-54-1875,1440-73-2535,2160-110-5460,2485-125-6235,2860-145-7210,3430-175-8680,5600-290-14300,7235-375-18485,8900-470-23000,10620-590-28320,15760-840-40960,18930-1010-49230
BPS-07,180-10-370,335-14-575,335-14-665,560-23-1020,750-31-1370,755-41-1370,1095-60-1995,1480-81-2695,2220-120-5820,2555-140-6755,2940-160-7740,3530-190-9230,5800-320-15400,7490-415-19940,9220-510-24520,10990-650-30490,16310-910-43610,19590-1100-52590
BPS-08,200-12-425,370-16-640,370-16-750,590-26-1110,790-34-1470,795-45-1470,1140-65-2115,1540-88-2860,2310-130-6210,2655-150-7155,3055-175-8305,3665-210-9965,6000-350-16500,7750-455-21400,9540-560-26340,11380-700-32380,16890-1000-46890,20290-1210-56590
BPS-09,225-15-480,390-20-700,390-20-820,620-29-1200,830-38-1590,840-50-1590,1185-72-2265,1605-97-3060,2410-145-6760,2770-165-7720,3185-190-8885,3820-230-10720,6200-380-17600,8015-495-22865,9860-610-28160,11770-760-34570,17470-1090-50170,20990-1310-60290
BPS-10,250-18-540,410-22-760,410-22-900,660-32-1300,870-42-1710,870-56-1710,1230-79-2415,1660-107-3265,2490-160-7290,2865-185-8415,3295-215-9745,3955-260-11755,6400-420-19000,8275-544-24595,10180-670-30280,12160-820-36760,18050-1190-53750,21680-1430-64580
BPS-11,275-20-600,430-24-830,430-24-980,700-35-1400,910-46-1830,915-61-1830,1275-86-2565,1725-116-3465,2590-175-7840,2980-200-8980,3430-230-10330,4115-275-12365,6600-460-20400,8540-595-26390,10510-740-32710,12570-900-39570,18650-1310-57950,22410-1580-69810
BPS-12,300-20-650,460-28-900,460-28-1060,750-40-1550,970-52-2010,975-69-2010,1355-96-2795,1830-130-3780,2745-195-8595,3155-225-9905,3630-260-11430,4355-310-13655,7000-500-22000,9055-650-28555,11140-800-35140,13320-960-42120,19770-1430-62670,23750-1720-75350
BPS-13,325-25-700,490-30-950,490-30-1125,800-45-1700,1035-58-2195,1040-77-2195,1440-107-3045,1950-144-4110,2925-215-9375,3365-245-10715,3870-285-12420,4645-340-14845,7500-550-24000,9700-715-31150,11930-880-38330,14260-1080-46660,21160-1560-67960,25420-1880-81820
BPS-14,350-25-750,520-30-1010,520-30-1210,850-50-1850,1100-64-2380,1105-85-2380,1530-119-3315,2065-161-4480,3100-240-10300,3565-275-11815,4100-315-13550,4920-380-16320,8000-610-26300,10340-790-34040,12720-980-42120,15180-1280-53580,22530-1740-74730,27060-2090-89760
BPS-15,375-25-825,550-35-1100,550-35-1350,900-55-2000,1165-71-2585,1160-95-2585,1620-131-3585,2190-177-4845,3285-265-11235,3780-305-12930,4350-350-14850,5220-420-17820,8500-700-29500,10985-905-38135,13510-1120-47110,16120-1330-56020,23920-1980-83320,28730-2380-100130
BPS-16,400-35-1000,625-40-1325,625-40-1625,1050-80-2250,1350-105-2925,1350-105-2925,1875-146-4065,2535-197-5490,3805-295-12655,4375-340-14575,5050-390-16750,6060-470-20160,10000-800-34000,12910-1035-43960,15880-1280-54280,18910-1520-64510,28070-2260-95870,33720-2720-115320
BPS-17,500-50-1250,900-50-1750,900-50-2250,1600-120-3040,2065-155-3925,2065-155-3925,2870-215-5450,3880-290-7360,6210-465-15510,7140-535-17840,8210-615-20510,9850-740-24650,16000-1200-40000,20680-1555-51780,25440-1930-64040,30370-2300-76370,45070-3420-113470,54140-4110-136340
BPS-18,1000-75-1750,1350-75-2150,1350-75-2650,2100-150-3600,2710-195-4660,2710-195-4660,3765-271-6475,5085-366-8745,8135-585-19835,9355-675-22855,10760-775-26260,12910-930-31510,20000-1500-50000,25940-1950-64940,31890-2400-79890,38350-2870-95750,56880-4260-142080,68330-5120-170730
BPS-19,1800-80-2200,2250-100-2750,2250-100-3050,3200-160-4480,4130-205-5770,4130-205-5770,5740-285-8590,7750-385-11600,12400-615-24700,14260-705-28360,16400-810-32600,19680-970-39080,31000-1600-63000,40155-2075-81655,49370-2560-100570,59210-3050-120210,87840-4530-178440,105510-5450-214510
"""

def parse_embedded_civil_scales():
    lines = RAW_CIVIL_SERVANT_DATA.strip().split("\n")
    headers = [h.strip() for h in lines[0].split(",")]
    date_cols = headers[1:]
    scales = {}
    for line in lines[1:]:
        parts = [p.strip() for p in line.split(",")]
        rk = parts[0]
        # Store both "BPS-01" and "BPS-1"
        rk_clean = f"BPS-{int(rk.split('-')[1])}"
        scales[rk] = {}
        scales[rk_clean] = {}
        for idx, dt in enumerate(date_cols):
            val = parts[idx+1]
            if val and "-" in val:
                nums = [int(n) for n in val.split("-")]
                sc = (nums[0], nums[1], nums[2])
                norm_dt = dt.replace('/', '-')
                if len(norm_dt) == 10 and norm_dt[2] == '-' and norm_dt[5] == '-':
                    # DD-MM-YYYY to YYYY-MM-DD
                    norm_dt = f"{norm_dt[6:]}-{norm_dt[3:5]}-{norm_dt[0:2]}"
                scales[rk][norm_dt] = sc
                scales[rk_clean][norm_dt] = sc
    return scales

DEFAULT_CIVIL_SCALES = parse_embedded_civil_scales()

# ==============================================================================
# 1. REGISTRY & FILE PERSISTENCE
# ==============================================================================
def load_registry():
    if os.path.exists(REGISTRY_FILE):
        try:
            with open(REGISTRY_FILE, "r") as f:
                return json.load(f)
        except Exception:
            pass
    # Initialize Registry with both categories
    reg = {
        "CAF": {
            "active": {
                "filename": "Official_CAF_Pay_Scales_1994_2026.xlsx",
                "approval_date": "2026-10-30 00:00:00",
                "effective_years": [1994, 2001, 2005, 2007, 2008, 2011, 2015, 2016, 2017, 2022, 2024, 2026],
                "ranks": list(BOOTSTRAP_CAF_SCALES.keys()),
                "data_file": "caf_active.json"
            },
            "backup": None
        },
        "Civil Servants": {
            "active": {
                "filename": "Official_Civil_Servants_BPS_1972_2026.csv",
                "approval_date": "2026-10-30 00:00:00",
                "effective_years": [1972, 1977, 1981, 1983, 1987, 1991, 1994, 2001, 2005, 2007, 2008, 2011, 2015, 2016, 2017, 2022, 2026],
                "ranks": [f"BPS-{i:02d}" for i in range(1, 20)],
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
# 2. ADVANCED UPLOAD PARSER (HANDLES EB, SLASHD RATES, MATRIX & LONG)
# ==============================================================================
def format_scale_str(sc):
    return f"{sc[0]}-{sc[1]}-{sc[2]}" if sc else ""

def normalize_rank_str(rk):
    rk = rk.strip()
    m = re.match(r'^bps-?0*(\d+)$', rk, re.IGNORECASE)
    if m:
        return f"BPS-{int(m.group(1)):02d}"
    return rk

def parse_scale_value(val_str):
    val_str = str(val_str).strip()
    if not val_str or val_str.lower() in ['nan', 'none', '-']:
        return None
    # Regular 3-number scale (e.g. 5400-270-13500)
    m = re.match(r'^(\d+)\s*[-/]\s*(\d+)\s*[-/]\s*(\d+)$', val_str)
    if m:
        return int(m.group(1)), int(m.group(2)), int(m.group(3))
    # Handle older scales with Efficiency Bar (EB) or multi-increment
    nums = re.findall(r'\b\d+\b', val_str)
    if len(nums) >= 3:
        # Min = first number, Max = last number, Inc = second number
        return int(nums[0]), int(nums[1]), int(nums[-1])
    return None

def parse_uploaded_chart(uploaded_file):
    try:
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file, header=None)
        else:
            df = pd.read_excel(uploaded_file, header=None)

        df = df.dropna(how='all')
        if df.empty:
            return None, "File is empty."

        # Find header row containing 'BPS' or 'Rank'
        hdr_idx = None
        for r_i in range(min(5, len(df))):
            row_vals = [str(x).lower() for x in df.iloc[r_i].values]
            if any("bps" in x or "rank" in x for x in row_vals):
                hdr_idx = r_i
                break
        if hdr_idx is None:
            hdr_idx = 0

        df.columns = df.iloc[hdr_idx]
        df = df.iloc[hdr_idx+1:].reset_index(drop=True)

        rank_col = None
        for col in df.columns:
            if any(k in str(col).lower() for k in ["bps", "rank", "designation"]):
                rank_col = col
                break
        if rank_col is None:
            rank_col = df.columns[0]

        scales = {}
        rect_rates = {"Matric": [], "Non-Matric": []}
        detected_years = set()
        date_cols = {}

        year_pat = re.compile(r'\b(19\d\d|20\d\d)\b')
        date_pat = re.compile(r'(\d{1,2}[-/]\d{1,2}[-/]\d{2,4})')

        for col in df.columns:
            if col == rank_col:
                continue
            col_str = str(col).strip()
            m_d = date_pat.search(col_str)
            m_y = year_pat.search(col_str)
            if m_d:
                raw_d = m_d.group(1).replace('/', '-')
                parts = raw_d.split('-')
                iso_dt = f"{parts[0]}-{int(parts[1]):02d}-{int(parts[2]):02d}" if len(parts[0]) == 4 else f"{int(parts[2]):04d}-{int(parts[1]):02d}-{int(parts[0]):02d}"
                date_cols[col] = iso_dt
                if m_y:
                    detected_years.add(int(m_y.group(1)))
            elif m_y:
                y_val = int(m_y.group(1))
                date_cols[col] = f"{y_val}-07-01"
                detected_years.add(y_val)

        for _, row in df.iterrows():
            raw_rk = str(row[rank_col]).strip()
            if not raw_rk or raw_rk.lower() == 'nan':
                continue
            norm_rk = normalize_rank_str(raw_rk)

            # Recruit stipend capture
            if "rect" in raw_rk.lower():
                cat = "Non-Matric" if "non" in raw_rk.lower() else "Matric"
                for col, dt in date_cols.items():
                    nums = re.findall(r'\b\d+\b', str(row[col]))
                    if nums:
                        rect_rates[cat].append((dt, int(nums[0])))
                continue

            for col, dt in date_cols.items():
                sc = parse_scale_value(row[col])
                if sc:
                    if norm_rk not in scales:
                        scales[norm_rk] = {}
                    scales[norm_rk][dt] = sc

        if not scales:
            return None, "No recognizable pay scale values found in document."

        return {
            "scales": scales,
            "rect_rates": rect_rates,
            "ranks": sorted(list(scales.keys())),
            "effective_years": sorted(list(detected_years)),
            "rows_read": len(df)
        }, None
    except Exception as e:
        return None, f"Parsing error: {str(e)}"

# ==============================================================================
# 3. RUNTIME CALCULATION ENGINE
# ==============================================================================
def get_scale_at_date(rank, target_date_str, scales_dict, rev_tag=None):
    norm_rk = normalize_rank_str(rank)
    if norm_rk not in scales_dict:
        # Try alternate rank notation
        alt = f"BPS-{int(norm_rk.split('-')[1])}" if "bps" in norm_rk.lower() else norm_rk
        if alt in scales_dict:
            norm_rk = alt
        else:
            return None
    r_scales = scales_dict[norm_rk]
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
    enrol_rank = normalize_rank_str(personnel.get("enrol_rank", "BPS-01"))
    qualification = personnel.get("qualification", "Matric")
    d_remust = personnel.get("d_remust")
    promotions = personnel.get("promotions", [])
    end_date_str = personnel.get("end_date", "2026-07-01")
    is_civil = (personnel.get("category") == "Civil Servants")
    warnings = []

    active_revisions = set()
    for rk, dt_dict in scales_dict.items():
        for d_key in dt_dict.keys():
            active_revisions.add(d_key)

    events = []
    if is_civil or enrol_rank == "NCs E":
        events.append({"date": doe, "type": "ENROLLMENT", "rank": enrol_rank, "remarks": f"Appointed as {enrol_rank}"})
    else:
        events.append({"date": doe, "type": "ENROLLMENT", "rank": "GD Rect", "remarks": f"Enrolled as Rect ({qualification})"})
        if d_remust:
            events.append({"date": d_remust, "type": "REMUSTERING", "rank": enrol_rank, "remarks": "Remustered as train sldr"})

    for p in promotions:
        p_rank = normalize_rank_str(p["rank"])
        events.append({"date": p["date"], "type": "PROMOTION", "rank": p_rank, "remarks": f"Pmtd to {p_rank}"})

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
    current_rank = enrol_rank
    current_scale = None
    substantive_scale = None
    last_change_dt = None

    for ev in events:
        ev_dt = datetime.datetime.strptime(ev["date"], "%Y-%m-%d").date()
        ev_t = ev["type"]

        if ev_t == "ENROLLMENT":
            if is_civil or enrol_rank == "NCs E":
                current_scale = get_scale_at_date(enrol_rank, ev["date"], scales_dict)
                substantive_scale = current_scale
                if not current_scale:
                    warnings.append(f"HOLD — REVIEW REQUIRED: Scale missing for {enrol_rank} on {ev['date']}.")
                    current_pay = "Review Req"
                else:
                    current_pay = current_scale[0]
                last_change_dt = ev_dt
                ledger.append({
                    "date": ev_dt.strftime("%d-%m-%Y"), "pay": current_pay,
                    "scale": format_scale_str(current_scale), "remarks": ev["remarks"]
                })
            else:
                rates = rect_rates.get(qualification, [])
                rect_pay = next((r[1] for r in reversed(rates) if r[0] <= ev["date"]), None)
                current_scale = get_scale_at_date(enrol_rank, ev["date"], scales_dict)
                substantive_scale = current_scale
                current_pay = rect_pay if rect_pay else (current_scale[0] if current_scale else "Review Req")
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
                inc_val = substantive_scale[1] if substantive_scale else (current_scale[1] if current_scale else 0)
                current_pay += inc_val
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
                honorary_inc = sm_scale[1] if sm_scale else 0
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
                        current_pay = p_step + new_sc[1]  # Equal/Over + 1 Inc
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
# 4. EXPORT HANDLERS (EXCEL, WORD, PDF)
# ==============================================================================
def generate_excel_bytes(data):
    wb = Workbook()
    ws = wb.active
    ws.title = "PAY FIXATION"
    ws.views.sheetView[0].showGridLines = True

    h = data["header"]
    ws.merge_cells("A1:E1")
    ws["A1"] = f"PAY FIXATION IRO {h.get('rank','')} {h.get('name','').upper()} ({h.get('fc_no','')})"
    ws["A1"].font = Font(name="Calibri", size=11, bold=True)
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")

    bio = [
        ("a", "Category & Active Chart", f"{h.get('category','Civil Servants')} [{h.get('chart_name','')}]"),
        ("b", "Date of Appointment / Enrolment", h.get("doe")),
        ("c", "Date of Remustering / Regularization", h.get("d_remust", "N/A (Direct Appt)"))
    ]
    letters = ["d", "e", "f", "g", "h", "i"]
    for idx, p in enumerate(h.get("promotions", [])):
        bio.append((letters[idx], f"Promotion to {p['rank']}", p["date"]))

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
    run = p_title.add_run(f"PAY FIXATION STATEMENT — {h.get('rank','')} {h.get('name','').upper()}")
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
        r[1].width = Inches(2.8)
        r[2].width = Inches(2.2)

    add_bio("a.", "Cadre / Chart", f"{h.get('category','Civil Servants')} [{h.get('chart_name','')}]")
    add_bio("b.", "Personnel / Service ID", h.get("fc_no"))
    add_bio("c.", "Date of Entry / Enrolment", h.get("doe"))
    letters = ["d.", "e.", "f.", "g.", "h."]
    for idx, p in enumerate(h.get("promotions", [])):
        add_bio(letters[idx], f"Promotion to {p['rank']}", p["date"])

    if data.get("special_instructions"):
        p_note = doc.add_paragraph()
        r_note = p_note.add_run(f"\nSpecial Directives: {data['special_instructions']} (Cross-check with Service Book)")
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
    title_text = f"<b>PAY FIXATION IRO {h.get('rank','')} {h.get('name','').upper()} ({h.get('fc_no','')})</b>"
    elements.append(Paragraph(title_text, styles['Title']))
    elements.append(Spacer(1, 8))

    bio_data = [
        ["a.", "Chart Category", f"{h.get('category','Civil Servants')} [{h.get('chart_name','')}]"],
        ["b.", "Date of Entry / Enrolment", str(h.get("doe"))],
        ["c.", "Date of Remustering / Regularization", str(h.get("d_remust", "N/A (Direct)"))]
    ]
    letters = ["d.", "e.", "f.", "g.", "h."]
    for idx, p in enumerate(h.get("promotions", [])):
        bio_data.append([letters[idx], f"Promotion to {p['rank']}", str(p["date"])])

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

# ==============================================================================
# 5. STREAMLIT APPLICATION INTERFACE
# ==============================================================================
def main():
    st.set_page_config(page_title="Pay Fixation Agent", layout="wide", page_icon="logo.png")

    col_logo, col_title = st.columns([1, 6])
    with col_logo:
        if os.path.exists("logo.png"):
            st.image("logo.png", width=95)
    with col_title:
        st.title("Pay Fixation Agent")
        st.caption("Developed by **SC Rehman Gul** for Civil Armed Forces (CAF) & Federal Civil Servants.")

    st.info("⚠️ **Audit Disclaimer:** All calculated basic pay fixations must be cross-checked against official service books and gazette notifications.")

    registry = load_registry()

    # --------------------------------------------------------------------------
    # SIDEBAR: CADRE SELECTION & REVISED CHART UPLOADER
    # --------------------------------------------------------------------------
    st.sidebar.header("📁 Pay Chart Management")
    target_category = st.sidebar.selectbox("Active Cadre / Category", ["Civil Servants", "CAF"], index=0)

    cat_meta = registry[target_category]["active"]
    st.sidebar.markdown(f"**Current Chart:** `{cat_meta['filename']}`")
    st.sidebar.markdown(f"**Approved On:** `{cat_meta['approval_date']}`")
    st.sidebar.markdown(f"**Revisions in Chart:** {len(cat_meta['effective_years'])} epochs ({cat_meta['effective_years'][0]} to {cat_meta['effective_years'][-1]})")

    st.sidebar.markdown("---")
    st.sidebar.subheader(f"Upload Revised {target_category} Chart")
    uploaded_chart = st.sidebar.file_uploader(
        f"Replace {target_category} Chart (.xlsx, .csv)",
        type=["xlsx", "xls", "csv"],
        key="new_chart_uploader"
    )

    if uploaded_chart is not None:
        parsed_res, err = parse_uploaded_chart(uploaded_chart)
        if err:
            st.sidebar.error(f"❌ {err}")
        else:
            st.sidebar.success(f"✅ Valid: {len(parsed_res['ranks'])} ranks, {len(parsed_res['effective_years'])} revision years detected.")
            with st.sidebar.expander("Preview Detected Scales"):
                st.write(parsed_res["ranks"])

            if st.sidebar.button(f"Confirm & Activate as New {target_category} Chart", type="primary"):
                registry[target_category]["backup"] = cat_meta
                new_data_file = f"{target_category.lower().replace(' ', '_')}_active.json"
                save_category_data(target_category, parsed_res["scales"], parsed_res["rect_rates"], new_data_file)

                registry[target_category]["active"] = {
                    "filename": uploaded_chart.name,
                    "approval_date": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "effective_years": parsed_res["effective_years"],
                    "ranks": parsed_res["ranks"],
                    "data_file": new_data_file
                }
                save_registry(registry)
                st.sidebar.success("🎉 New chart activated!")
                st.rerun()

    # --------------------------------------------------------------------------
    # MAIN FORM: PERSONNEL RECORD
    # --------------------------------------------------------------------------
    active_scales, active_rect_rates = load_category_data(target_category, cat_meta["data_file"])
    available_ranks = sorted(list(active_scales.keys()))
    if not available_ranks:
        available_ranks = [f"BPS-{i:02d}" for i in range(1, 20)] if target_category == "Civil Servants" else ["GD Sep", "GD Lnk", "GD Nk", "GD Hav", "GD N/Sub", "Sub", "SM"]

    col1, col2 = st.columns(2)
    with col1:
        st.subheader("👤 Employee Bio-Data")
        fc_no = st.text_input("Personnel / Service / FC No.", value="CS-10492" if target_category == "Civil Servants" else "102052")
        name = st.text_input("Full Name", value="Muhammad Ali" if target_category == "Civil Servants" else "Jamal")

        final_rank = st.selectbox("Current / Substantive Rank", available_ranks, index=min(10, len(available_ranks)-1))
        enrol_rank = st.selectbox("Appointment / Enrolled Rank", available_ranks, index=0)

        qualification = "Matric"
        if target_category == "CAF":
            qualification = st.radio("Recruit Educational Category", ["Matric", "Non-Matric"], horizontal=True)

        doe = st.date_input("Date of Appointment / Enrolment",
                            value=datetime.date(1994, 6, 1),
                            min_value=datetime.date(1972, 1, 1),
                            max_value=datetime.date(2035, 12, 31)).strftime("%Y-%m-%d")

        d_remust = None
        if target_category == "CAF" and enrol_rank != "NCs E":
            d_remust = st.date_input("Date of Remustering",
                                     value=datetime.date(1994, 12, 1),
                                     min_value=datetime.date(1972, 1, 1),
                                     max_value=datetime.date(2035, 12, 31)).strftime("%Y-%m-%d")

    with col2:
        st.subheader("📈 Promotion History")
        st.caption(f"Enter promotion events validated against active {target_category} chart:")
        num_promos = st.number_input("Number of Promotions", min_value=0, max_value=8, value=3, step=1)
        promotions = []
        default_promos = [
            (datetime.date(2001, 12, 1), available_ranks[min(4, len(available_ranks)-1)]),
            (datetime.date(2011, 7, 1), available_ranks[min(8, len(available_ranks)-1)]),
            (datetime.date(2017, 7, 1), available_ranks[min(13, len(available_ranks)-1)]),
            (datetime.date(2022, 7, 1), available_ranks[min(15, len(available_ranks)-1)])
        ]
        for i in range(num_promos):
            p_c1, p_c2 = st.columns([1, 1])
            def_dt, def_rk = default_promos[i] if i < len(default_promos) else (datetime.date(2020, 1, 1), available_ranks[0])
            with p_c1:
                p_dt = st.date_input(f"Promo {i+1} Date", value=def_dt,
                                     min_value=datetime.date(1972, 1, 1),
                                     max_value=datetime.date(2035, 12, 31),
                                     key=f"promo_dt_{i}").strftime("%Y-%m-%d")
            with p_c2:
                cur_idx = available_ranks.index(def_rk) if def_rk in available_ranks else 0
                p_rk = st.selectbox(f"Promo {i+1} Rank", available_ranks, index=cur_idx, key=f"promo_rk_{i}")
            promotions.append({"date": p_dt, "rank": p_rk})

    st.markdown("---")
    st.subheader("📝 Audit Directives & Special Instructions")
    special_instructions = st.text_input(
        "Enter case-specific orders or notes (e.g., 'Adv increment on qualification', 'Move-over order', 'Personal Pay protected'):",
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
