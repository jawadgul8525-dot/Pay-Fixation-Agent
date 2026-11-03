import streamlit as st
import datetime
import io
import os
import re
import json
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
# GLOBAL CONSTANTS & DATE BOUNDARIES (1990 to 2030 Inclusive)
# ==============================================================================
MIN_APP_DATE = datetime.date(1990, 1, 1)
MAX_APP_DATE = datetime.date(2030, 12, 31)

REGISTRY_SCHEMA_VERSION = "2026.11.v8_AUTO_CADRE_UPGRADATION_RESTORED"

ALL_BPS_SCALES = [f"BPS-{i:02d}" for i in range(1, 23)]
CAF_GD_RANKS = ["GD Sep", "GD Lnk", "GD Nk", "GD Hav", "GD N/Sub", "Sub", "SM", "H/Lt", "H/Capt"]
CAF_TECH_RANKS = ["Tech Sep", "Tech Lnk", "Tech Nk", "Tech Hav"]

CIVIL_COMMON_POSTS = [
    "Junior Clerk (LDC)",
    "Senior Clerk (UDC)",
    "Assistant",
    "Superintendent",
    "Stenotypist",
    "Assistant Private Secretary (APS)",
    "Private Secretary (PS)",
    "Junior Auditor",
    "Senior Auditor",
    "Naib Qasid",
    "Daftary",
    "Record Sorter",
    "Driver",
    "Custom / Other Designation"
]

DATA_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")
REGISTRY_FILE = os.path.join(DATA_DIR, "chart_registry.json")
os.makedirs(DATA_DIR, exist_ok=True)

# ==============================================================================
# 0. LOGO & DATE FORMATTING UTILITIES
# ==============================================================================
def get_logo_path():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    candidates = [
        os.path.join(base_dir, "logo.png"),
        os.path.join(base_dir, "assets", "logo.png"),
        os.path.join(os.getcwd(), "logo.png"),
        "logo.png"
    ]
    for p in candidates:
        if os.path.exists(p) and os.path.isfile(p):
            return os.path.abspath(p)
    for root, _, files in os.walk(base_dir):
        for f in files:
            if f.lower() == "logo.png":
                return os.path.abspath(os.path.join(root, f))
    return None

def to_display_date(d_val):
    if not d_val or str(d_val).strip() in ["", "None", "N/A"]:
        return "N/A"
    try:
        if isinstance(d_val, (datetime.date, datetime.datetime)):
            return d_val.strftime("%d-%m-%Y")
        clean_d = str(d_val).strip()[:10].replace('/', '-')
        parts = clean_d.split('-')
        if len(parts) == 3:
            if len(parts[0]) == 4:
                return f"{int(parts[2]):02d}-{int(parts[1]):02d}-{parts[0]}"
            else:
                return f"{int(parts[0]):02d}-{int(parts[1]):02d}-{parts[2]}"
    except Exception:
        pass
    return str(d_val)

# ==============================================================================
# 1. VERIFIED CIVIL SERVANTS PAY SCALES (BPS-01 to BPS-22: 1983 - 2026)
# ==============================================================================
DEFAULT_CIVIL_SCALES = {
    "BPS-01": {"1983-07-01":(440,10,640),"1987-07-01":(600,13,860),"1991-06-01":(920,26,1310),"1994-06-01":(1245,35,1770),"2001-12-01":(1870,55,3520),"2005-07-01":(2150,65,4100),"2007-07-01":(2475,75,4725),"2008-07-01":(2970,90,5670),"2011-07-01":(4800,150,9300),"2015-07-01":(6210,195,12060),"2016-07-01":(7640,240,14840),"2017-07-01":(9130,290,17830),"2022-07-01":(13550,430,26450),"2026-07-01":(16280,520,31880)},
    "BPS-02": {"1983-07-01":(460,12,700),"1987-07-01":(625,16,945),"1991-06-01":(945,32,1425),"1994-06-01":(1275,44,1935),"2001-12-01":(1915,65,3865),"2005-07-01":(2200,75,4450),"2007-07-01":(2530,85,5080),"2008-07-01":(3035,100,6035),"2011-07-01":(4900,170,10000),"2015-07-01":(6335,220,12935),"2016-07-01":(7790,275,16040),"2017-07-01":(9310,330,19210),"2022-07-01":(13820,490,28520),"2026-07-01":(16600,590,34300)},
    "BPS-03": {"1983-07-01":(480,14,760),"1987-07-01":(650,19,1030),"1991-06-01":(975,37,1530),"1994-06-01":(1320,50,2070),"2001-12-01":(1980,75,4230),"2005-07-01":(2275,85,4825),"2007-07-01":(2615,100,5615),"2008-07-01":(3140,120,6740),"2011-07-01":(5050,200,11050),"2015-07-01":(6535,260,14335),"2016-07-01":(8040,325,17790),"2017-07-01":(9610,390,21310),"2022-07-01":(14260,580,31660),"2026-07-01":(17130,700,38130)},
    "BPS-04": {"1983-07-01":(500,16,820),"1987-07-01":(675,22,1115),"1991-06-01":(1005,43,1650),"1994-06-01":(1360,58,2230),"2001-12-01":(2040,85,4590),"2005-07-01":(2345,100,5345),"2007-07-01":(2700,115,6150),"2008-07-01":(3240,140,7440),"2011-07-01":(5200,230,12100),"2015-07-01":(6730,300,15730),"2016-07-01":(8280,370,19380),"2017-07-01":(9900,440,23100),"2022-07-01":(14690,660,34490),"2026-07-01":(17650,800,41650)},
    "BPS-05": {"1983-07-01":(520,18,880),"1987-07-01":(700,25,1200),"1991-06-01":(1035,49,1770),"1994-06-01":(1400,66,2390),"2001-12-01":(2100,100,5100),"2005-07-01":(2415,115,5865),"2007-07-01":(2780,135,6830),"2008-07-01":(3340,160,8140),"2011-07-01":(5400,260,13200),"2015-07-01":(6985,340,17185),"2016-07-01":(8590,420,21190),"2017-07-01":(10260,500,25260),"2022-07-01":(15230,750,37730),"2026-07-01":(18300,910,45600)},
    "BPS-06": {"1983-07-01":(540,20,940),"1987-07-01":(725,28,1285),"1991-06-01":(1065,54,1875),"1994-06-01":(1440,73,2535),"2001-12-01":(2160,110,5460),"2005-07-01":(2485,125,6235),"2007-07-01":(2860,145,7210),"2008-07-01":(3430,175,8680),"2011-07-01":(5600,290,14300),"2015-07-01":(7235,375,18485),"2016-07-01":(8900,470,23000),"2017-07-01":(10620,560,27420),"2022-07-01":(15760,840,40960),"2026-07-01":(18930,1010,49230)},
    "BPS-07": {"1983-07-01":(560,23,1020),"1987-07-01":(750,31,1370),"1991-06-01":(1095,60,1995),"1994-06-01":(1480,81,2695),"2001-12-01":(2220,120,5820),"2005-07-01":(2555,140,6755),"2007-07-01":(2940,160,7740),"2008-07-01":(3530,190,9230),"2011-07-01":(5800,320,15400),"2015-07-01":(7490,415,19940),"2016-07-01":(9220,510,24520),"2017-07-01":(10990,610,29290),"2022-07-01":(16310,910,43610),"2026-07-01":(19590,1100,52590)},
    "BPS-08": {"1983-07-01":(590,26,1110),"1987-07-01":(790,34,1470),"1991-06-01":(1140,65,2115),"1994-06-01":(1540,88,2860),"2001-12-01":(2310,130,6210),"2005-07-01":(2655,150,7155),"2007-07-01":(3055,175,8305),"2008-07-01":(3665,210,9965),"2011-07-01":(6000,350,16500),"2015-07-01":(7750,455,21400),"2016-07-01":(9540,560,26340),"2017-07-01":(11380,670,31480),"2022-07-01":(16890,1000,46890),"2026-07-01":(20290,1210,56590)},
    "BPS-09": {"1983-07-01":(620,29,1200),"1987-07-01":(830,38,1590),"1991-06-01":(1185,72,2265),"1994-06-01":(1605,97,3060),"2001-12-01":(2410,145,6760),"2005-07-01":(2770,165,7720),"2007-07-01":(3185,190,8885),"2008-07-01":(3820,230,10720),"2011-07-01":(6200,380,17600),"2015-07-01":(8015,495,22865),"2016-07-01":(9860,610,28160),"2017-07-01":(11770,730,33670),"2022-07-01":(17470,1090,50170),"2026-07-01":(20990,1310,60290)},
    "BPS-10": {"1983-07-01":(660,32,1300),"1987-07-01":(870,42,1710),"1991-06-01":(1230,79,2415),"1994-06-01":(1660,107,3265),"2001-12-01":(2490,160,7290),"2005-07-01":(2865,185,8415),"2007-07-01":(3295,215,9745),"2008-07-01":(3955,260,11755),"2011-07-01":(6400,420,19000),"2015-07-01":(8275,544,24595),"2016-07-01":(10180,670,30280),"2017-07-01":(12160,800,36160),"2022-07-01":(18050,1190,53750),"2026-07-01":(21680,1430,64580)},
    "BPS-11": {"1983-07-01":(700,35,1400),"1987-07-01":(910,46,1830),"1991-06-01":(1275,86,2565),"1994-06-01":(1725,116,3465),"2001-12-01":(2590,175,7840),"2005-07-01":(2980,200,8980),"2007-07-01":(3430,230,10330),"2008-07-01":(4115,275,12365),"2011-07-01":(6600,460,20400),"2015-07-01":(8540,595,26390),"2016-07-01":(10510,740,32710),"2017-07-01":(12570,880,38970),"2022-07-01":(18650,1310,57950),"2026-07-01":(22410,1580,69810)},
    "BPS-12": {"1983-07-01":(750,40,1550),"1987-07-01":(970,52,2010),"1991-06-01":(1355,96,2795),"1994-06-01":(1830,130,3780),"2001-12-01":(2745,195,8595),"2005-07-01":(3155,225,9905),"2007-07-01":(3630,260,11430),"2008-07-01":(4355,310,13655),"2011-07-01":(7000,500,22000),"2015-07-01":(9055,650,28555),"2016-07-01":(11140,800,35140),"2017-07-01":(13320,960,42120),"2022-07-01":(19770,1430,62670),"2026-07-01":(23750,1720,75350)},
    "BPS-13": {"1983-07-01":(800,45,1700),"1987-07-01":(1035,58,2195),"1991-06-01":(1440,107,3045),"1994-06-01":(1950,144,4110),"2001-12-01":(2925,215,9375),"2005-07-01":(3365,245,10715),"2007-07-01":(3870,285,12420),"2008-07-01":(4645,340,14845),"2011-07-01":(7500,550,24000),"2015-07-01":(9700,715,31150),"2016-07-01":(11930,880,38330),"2017-07-01":(14260,1050,45760),"2022-07-01":(21160,1560,67960),"2026-07-01":(25420,1880,81820)},
    "BPS-14": {"1983-07-01":(850,50,1850),"1987-07-01":(1100,64,2380),"1991-06-01":(1530,119,3315),"1994-06-01":(2065,161,4480),"2001-12-01":(3100,240,10300),"2005-07-01":(3565,275,11815),"2007-07-01":(4100,315,13550),"2008-07-01":(4920,380,16320),"2011-07-01":(8000,610,26300),"2015-07-01":(10340,790,34040),"2016-07-01":(12720,980,42120),"2017-07-01":(15180,1170,50280),"2022-07-01":(22530,1740,74730),"2026-07-01":(27060,2090,89760)},
    "BPS-15": {"1983-07-01":(900,55,2000),"1987-07-01":(1165,71,2585),"1991-06-01":(1620,131,3585),"1994-06-01":(2190,177,4845),"2001-12-01":(3285,265,11235),"2005-07-01":(3780,305,12930),"2007-07-01":(4350,350,14850),"2008-07-01":(5220,420,17820),"2011-07-01":(8500,700,29500),"2015-07-01":(10985,905,38135),"2016-07-01":(13510,1120,47110),"2017-07-01":(16120,1330,56020),"2022-07-01":(23920,1980,83320),"2026-07-01":(28730,2380,100130)},
    "BPS-16": {"1983-07-01":(1050,80,2250),"1987-07-01":(1350,105,2925),"1991-06-01":(1875,146,4065),"1994-06-01":(2535,197,5490),"2001-12-01":(3805,295,12655),"2005-07-01":(4375,340,14575),"2007-07-01":(5050,390,16750),"2008-07-01":(6060,470,20160),"2011-07-01":(10000,800,34000),"2015-07-01":(12910,1035,43960),"2016-07-01":(15880,1280,54280),"2017-07-01":(18910,1520,64510),"2022-07-01":(28070,2260,95870),"2026-07-01":(33720,2720,115320)},
    "BPS-17": {"1983-07-01":(1600,120,3040),"1987-07-01":(2065,155,3925),"1991-06-01":(2870,215,5450),"1994-06-01":(3880,290,7360),"2001-12-01":(6210,465,15510),"2005-07-01":(7140,535,17840),"2007-07-01":(8210,615,20510),"2008-07-01":(9850,740,24650),"2011-07-01":(16000,1200,40000),"2015-07-01":(20680,1555,51780),"2016-07-01":(25440,1930,64040),"2017-07-01":(30370,2300,76370),"2022-07-01":(45070,3420,113470),"2026-07-01":(54140,4110,136340)},
    "BPS-18": {"1983-07-01":(2100,150,3600),"1987-07-01":(2710,195,4660),"1991-06-01":(3765,271,6475),"1994-06-01":(5085,366,8745),"2001-12-01":(8135,585,19835),"2005-07-01":(9355,675,22855),"2007-07-01":(10760,775,26260),"2008-07-01":(12910,930,31510),"2011-07-01":(20000,1500,50000),"2015-07-01":(25940,1950,64940),"2016-07-01":(31890,2400,79890),"2017-07-01":(38350,2870,95750),"2022-07-01":(56880,4260,142080),"2026-07-01":(68330,5120,170730)},
    "BPS-19": {"1983-07-01":(3200,160,4480),"1987-07-01":(4130,205,5770),"1991-06-01":(5740,285,8590),"1994-06-01":(7750,385,11600),"2001-12-01":(12400,615,24700),"2005-07-01":(14260,705,28360),"2007-07-01":(16400,810,32600),"2008-07-01":(19680,970,39080),"2011-07-01":(31000,1600,63000),"2015-07-01":(40155,2075,81655),"2016-07-01":(49370,2560,100570),"2017-07-01":(59210,3050,120210),"2022-07-01":(87840,4530,178440),"2026-07-01":(105510,5450,214510)},
    "BPS-20": {"1983-07-01":(3700,230,6000),"1987-07-01":(4800,300,7800),"1991-06-01":(6675,415,10825),"1994-06-01":(9010,560,14610),"2001-12-01":(14420,895,23370),"2005-07-01":(16580,1030,26880),"2007-07-01":(19070,1185,30920),"2008-07-01":(22880,1420,37080),"2011-07-01":(36000,2350,59500),"2015-07-01":(46830,3075,71430),"2016-07-01":(57600,3780,87840),"2017-07-01":(69090,4510,105190),"2022-07-01":(102490,6690,156010),"2026-07-01":(122990,8030,187220)},
    "BPS-21": {"1983-07-01":(4100,250,6600),"1987-07-01":(5350,330,8650),"1991-06-01":(7435,460,12035),"1994-06-01":(10035,625,16285),"2001-12-01":(16055,995,26005),"2005-07-01":(18465,1145,29915),"2007-07-01":(21235,1315,34385),"2008-07-01":(25480,1580,41280),"2011-07-01":(40000,2600,66000),"2015-07-01":(52150,3410,79430),"2016-07-01":(64140,4190,97620),"2017-07-01":(76880,5000,116880),"2022-07-01":(114040,7420,173400),"2026-07-01":(136850,8900,208050)},
    "BPS-22": {"1983-07-01":(4400,300,7400),"1987-07-01":(5750,390,9650),"1991-06-01":(7990,545,13440),"1994-06-01":(10790,735,18140),"2001-12-01":(17265,1175,29015),"2005-07-01":(19855,1350,33355),"2007-07-01":(22835,1550,38335),"2008-07-01":(27400,1860,46000),"2011-07-01":(43000,3050,73500),"2015-07-01":(55890,4005,87930),"2016-07-01":(68740,4920,108100),"2017-07-01":(82380,5870,129340),"2022-07-01":(122190,8710,191870),"2026-07-01":(146630,10450,230230)}
}

for i in range(1, 10):
    k_dbl = f"BPS-{i:02d}"
    k_sng = f"BPS-{i}"
    if k_dbl in DEFAULT_CIVIL_SCALES:
        DEFAULT_CIVIL_SCALES[k_sng] = DEFAULT_CIVIL_SCALES[k_dbl]

# ==============================================================================
# 2. VERIFIED CAF SCALES (INCLUDING NEW 2026 TECH HAV AMENDMENT: 25880-1810-80180)
# ==============================================================================
BOOTSTRAP_CAF_SCALES = {
    "NCs E": {
        "1994-06-01":(1245,35,1770),"2001-12-01":(1870,55,3520),"2005-07-01_REV":(2130,65,4100),"2005-07-01_UPG":(2130,65,4100),"2007-07-01":(2475,75,4725),"2008-07-01":(2970,90,5670),"2011-07-01":(4800,150,9300),"2015-07-01":(6210,195,12060),"2016-07-01":(7610,240,14840),"2017-07-01":(9125,285,17675),"2022-07-01":(13540,430,26440),"2024-06-01":(15620,490,30320),"2026-07-01":(18780,590,36480)
    },
    "GD Sep": {
        "1994-06-01":(1254,30,1800),"2001-12-01":(1875,75,3525),"2005-07-01_REV":(2155,85,4025),"2005-07-01_UPG":(2415,115,5865),"2007-07-01":(2780,135,6830),"2008-07-01":(3340,160,8140),"2011-07-01":(5400,270,13500),"2015-07-01":(6985,350,17485),"2016-07-01":(8590,430,21490),"2017-07-01":(10260,510,25560),"2022-07-01":(15230,760,38030),"2024-06-01":(18160,950,46660),"2026-07-01":(21830,1150,56330)
    },
    "Tech Sep": {
        "1994-06-01":(1262,61,1933),"2001-12-01":(1895,95,3985),"2005-07-01_REV":(2180,110,4600),"2005-07-01_UPG":(2415,115,5865),"2007-07-01":(2780,135,6830),"2008-07-01":(3340,160,8140),"2011-07-01":(5400,270,13500),"2015-07-01":(6985,350,17485),"2016-07-01":(8590,430,21490),"2017-07-01":(10260,510,25560),"2022-07-01":(15230,760,38030)
    },
    "GD Lnk": {
        "1994-06-01":(1262,61,1933),"2001-12-01":(1895,95,3985),"2005-07-01_REV":(2180,110,4600),"2005-07-01_UPG":(2485,125,6235),"2007-07-01":(2860,145,7210),"2008-07-01":(3430,175,8680),"2011-07-01":(5500,290,14200),"2015-07-01":(7115,375,18365),"2016-07-01":(8750,470,22850),"2017-07-01":(10465,560,27265),"2022-07-01":(15530,840,40730),"2024-06-01":(18780,1040,49980),"2026-07-01":(22580,1250,60080)
    },
    "Tech Lnk": {
        "1994-06-01":(1400,66,2390),"2001-12-01":(2400,100,5100),"2005-07-01_REV":(2415,115,5865),"2005-07-01_UPG":(2485,125,6235),"2007-07-01":(2860,145,7210),"2008-07-01":(3430,175,8680),"2011-07-01":(5500,290,14200),"2015-07-01":(7115,375,18365),"2016-07-01":(8750,470,22850),"2017-07-01":(10465,560,27265),"2022-07-01":(15530,840,40730)
    },
    "GD Nk": {
        "1994-06-01":(1400,66,2390),"2001-12-01":(2400,100,5100),"2005-07-01_REV":(2415,115,5865),"2005-07-01_UPG":(2555,140,6755),"2007-07-01":(2940,160,7740),"2008-07-01":(3530,190,9230),"2011-07-01":(5700,310,15000),"2015-07-01":(7370,405,19520),"2016-07-01":(9070,500,24070),"2017-07-01":(10835,595,28665),"2022-07-01":(16080,890,42780),"2024-06-01":(19450,1140,53650),"2026-07-01":(23380,1380,64780)
    },
    "Tech Nk": {
        "1994-06-01":(1440,73,2535),"2001-12-01":(2160,110,5460),"2005-07-01_REV":(2485,125,6235),"2005-07-01_UPG":(2555,140,6755),"2007-07-01":(2940,160,7740),"2008-07-01":(3530,190,9230),"2011-07-01":(5700,310,15000),"2015-07-01":(7370,405,19520),"2016-07-01":(9070,500,24070),"2017-07-01":(10835,595,28665),"2022-07-01":(16080,890,42780)
    },
    "GD Hav": {
        "1994-06-01":(1440,73,2535),"2001-12-01":(2160,110,5460),"2005-07-01_REV":(2485,125,6235),"2005-07-01_UPG":(2655,150,7155),"2007-07-01":(3055,175,8305),"2008-07-01":(3665,210,9965),"2011-07-01":(5900,350,16400),"2015-07-01":(7630,452,21190),"2016-07-01":(9390,560,26190),"2017-07-01":(11225,665,31175),"2022-07-01":(16660,990,46360),"2024-06-01":(20140,1240,57340),"2026-07-01":(24210,1500,69210)
    },
    "Tech Hav": {
        "1994-06-01":(1605,97,3060),"2001-12-01":(2410,145,6760),"2005-07-01_REV":(2770,165,7720),"2007-07-01":(3185,190,8885),"2008-07-01":(3825,230,10725),"2011-07-01":(6200,380,17600),"2015-07-01":(8015,495,22865),"2016-07-01":(9860,610,28160),"2017-07-01":(11775,725,26275),"2022-07-01":(17480,1080,39080),"2024-06-01":(21530,1500,66530),"2026-07-01":(25880,1810,80180)
    },
    "GD N/Sub": {
        "1994-06-01":(1725,116,3465),"2001-12-01":(2590,175,7840),"2005-07-01_REV":(2980,200,8980),"2007-07-01":(3430,230,10330),"2008-07-01":(4115,275,12365),"2011-07-01":(6600,120,20100),"2015-07-01":(8540,585,26090),"2016-07-01":(10310,720,32110),"2017-07-01":(12570,860,38370),"2022-07-01":(18650,1280,57050),"2024-06-01":(26640,1790,80340),"2026-07-01":(32030,2160,96830)
    },
    "Sub": {
        "1994-06-01":(2338,144,4110),"2001-12-01":(3360,215,8950),"2005-07-01_REV":(3865,245,10235),"2007-07-01":(4445,280,11725),"2008-07-01":(5335,335,14045),"2011-07-01":(8600,550,25100),"2015-07-01":(11120,710,32420),"2016-07-01":(13680,880,40080),"2017-07-01":(16350,1630,47850),"2022-07-01":(24260,1560,71060),"2024-06-01":(30520,2270,98620),"2026-07-01":(36690,2730,118590)
    },
    "SM": {
        "1994-06-01":(3800,197,5376),"2001-12-01":(3700,295,10420),"2005-07-01_REV":(6555,340,11995),"2007-07-01":(7540,390,13780),"2008-07-01":(9050,470,16570),"2011-07-01":(14600,770,37700),"2015-07-01":(18880,995,48730),"2016-07-01":(23220,1230,60120),"2017-07-01":(27745,1465,71695),"2022-07-01":(41170,2180,106570),"2024-06-01":(46610,2580,124010),"2026-07-01":(56030,3110,149330)
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

# ==============================================================================
# 3. REGISTRY STORAGE & PERSISTENCE
# ==============================================================================
def load_registry():
    if os.path.exists(REGISTRY_FILE):
        try:
            with open(REGISTRY_FILE, "r") as f:
                reg = json.load(f)
                if reg.get("_version") == REGISTRY_SCHEMA_VERSION:
                    return reg
        except Exception:
            pass

    reg = {
        "_version": REGISTRY_SCHEMA_VERSION,
        "CAF": {
            "active": {
                "filename": "Official_CAF_Pay_Scales_1994_2026_TechHav.xlsx",
                "approval_date": "2026-11-03 00:00:00",
                "effective_years": [1994, 2001, 2005, 2007, 2008, 2011, 2015, 2016, 2017, 2022, 2024, 2026],
                "ranks": list(BOOTSTRAP_CAF_SCALES.keys()),
                "data_file": "caf_active.json"
            },
            "backup": None
        },
        "Civil Servants": {
            "active": {
                "filename": "Official_Federal_Civil_Servants_BPS_1983_2026_Corrected.csv",
                "approval_date": "2026-11-03 00:00:00",
                "effective_years": [1983, 1987, 1991, 1994, 2001, 2005, 2007, 2008, 2011, 2015, 2016, 2017, 2022, 2026],
                "ranks": ALL_BPS_SCALES,
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
    base_scales = DEFAULT_CIVIL_SCALES if category == "Civil Servants" else BOOTSTRAP_CAF_SCALES
    base_rect = {} if category == "Civil Servants" else DEFAULT_RECT_RATES
    path = os.path.join(DATA_DIR, filename)

    if os.path.exists(path):
        try:
            with open(path, "r") as f:
                d = json.load(f)
                disk_scales = d.get("scales", {})
                disk_rect = d.get("rect_rates", {})
                if category == "CAF":
                    if disk_scales.get("Tech Hav", {}).get("2026-07-01", (0,0,0))[0] == 25880:
                        return disk_scales, disk_rect
                else:
                    if len(disk_scales) >= 19:
                        return disk_scales, disk_rect
        except Exception:
            pass

    save_category_data(category, base_scales, base_rect, filename)
    return base_scales, base_rect

# ==============================================================================
# 4. CHART PARSER ENGINE FOR UPLOADS
# ==============================================================================
def format_scale_str(sc):
    return f"{sc[0]}-{sc[1]}-{sc[2]}" if sc else ""

def normalize_rank_str(rk):
    rk = str(rk).strip()
    m = re.match(r'^bps-?0*(\d+)$', rk, re.IGNORECASE)
    if m:
        return f"BPS-{int(m.group(1)):02d}"
    return rk

def parse_scale_value(val_str):
    val_str = str(val_str).strip()
    if not val_str or val_str.lower() in ['nan', 'none', '-', '']:
        return None
    m = re.match(r'^(\d+)\s*[-/]\s*(\d+)\s*[-/]\s*(\d+)$', val_str)
    if m:
        return int(m.group(1)), int(m.group(2)), int(m.group(3))
    nums = re.findall(r'\b\d+\b', val_str)
    if len(nums) >= 3:
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
            return None, "File is completely empty."

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
            for col, dt in date_cols.items():
                sc = parse_scale_value(row[col])
                if sc:
                    if norm_rk not in scales:
                        scales[norm_rk] = {}
                    scales[norm_rk][dt] = sc

        if not scales:
            return None, "No recognizable pay scale structures found in file."

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
# 5. RESTORED & EXTENDED CADRE POST UP-GRADATION ENGINE (CIVIL SERVANTS ONLY)
# ==============================================================================
FEDERAL_CADRE_RULES = [
    {
        "pattern": r"(junior\s*clerk|ldc|lower\s*division\s*clerk|jc\b)",
        "designation": "Junior Clerk",
        "upgradations": [
            ("2007-07-01", "BPS-07", "Finance Div OM No. 6(4)R-1/2006 (01-07-2007)"),
            ("2016-07-01", "BPS-09", "Finance Div OM No. 9(7)R-1/2014-515/2016 (01-07-2016)"),
            ("2023-01-01", "BPS-11", "Finance Div OM No. 9(7)R-1/2014-62/2023 (01-01-2023)")
        ],
        "historical_initial": [
            ("2023-01-01", "BPS-11"),
            ("2016-07-01", "BPS-09"),
            ("2007-07-01", "BPS-07"),
            ("1983-01-01", "BPS-05")
        ]
    },
    {
        "pattern": r"(senior\s*clerk|udc|upper\s*division\s*clerk|sc\b)",
        "designation": "Senior Clerk",
        "upgradations": [
            ("2007-07-01", "BPS-09", "Finance Div OM No. 6(4)R-1/2006 (01-07-2007)"),
            ("2016-07-01", "BPS-11", "Finance Div OM No. 9(7)R-1/2014-515/2016 (01-07-2016)"),
            ("2023-01-01", "BPS-13", "Finance Div OM No. 9(7)R-1/2014-62/2023 (01-01-2023)")
        ],
        "historical_initial": [
            ("2023-01-01", "BPS-13"),
            ("2016-07-01", "BPS-11"),
            ("2007-07-01", "BPS-09"),
            ("1983-01-01", "BPS-07")
        ]
    },
    {
        "pattern": r"(assistant(?!\s*private)|head\s*clerk)",
        "designation": "Assistant",
        "upgradations": [
            ("2007-07-01", "BPS-14", "Finance Div OM No. 6(4)R-1/2006 (01-07-2007)"),
            ("2016-07-01", "BPS-16", "Finance Div OM No. 9(7)R-1/2014-515/2016 (01-07-2016)")
        ],
        "historical_initial": [
            ("2016-07-01", "BPS-16"),
            ("2007-07-01", "BPS-14"),
            ("1983-01-01", "BPS-11")
        ]
    },
    {
        "pattern": r"(superintendent)",
        "designation": "Superintendent",
        "upgradations": [
            ("2007-07-01", "BPS-16", "Finance Div OM No. 6(4)R-1/2006 (01-07-2007)"),
            ("2016-07-01", "BPS-17", "Finance Div OM No. 9(7)R-1/2014-515/2016 (01-07-2016)")
        ],
        "historical_initial": [
            ("2016-07-01", "BPS-17"),
            ("2007-07-01", "BPS-16"),
            ("1983-01-01", "BPS-14")
        ]
    },
    {
        "pattern": r"(stenotypist)",
        "designation": "Stenotypist",
        "upgradations": [
            ("2011-12-23", "BPS-14", "Finance Div OM No. 1(29)R-I/2010 (23-12-2011)")
        ],
        "historical_initial": [
            ("2011-12-23", "BPS-14"),
            ("1983-01-01", "BPS-12")
        ]
    },
    {
        "pattern": r"(stenographer|assistant\s*private\s*secretary|aps\b)",
        "designation": "Assistant Private Secretary",
        "upgradations": [
            ("2011-12-23", "BPS-16", "Finance Div OM No. 1(29)R-I/2010 (23-12-2011)")
        ],
        "historical_initial": [
            ("2011-12-23", "BPS-16"),
            ("1983-01-01", "BPS-15")
        ]
    },
    {
        "pattern": r"(junior\s*auditor)",
        "designation": "Junior Auditor",
        "upgradations": [
            ("2016-07-01", "BPS-11", "Finance Div OM No. 9(7)R-1/2014-515/2016 (01-07-2016)")
        ],
        "historical_initial": [
            ("2016-07-01", "BPS-11"),
            ("1983-01-01", "BPS-07")
        ]
    },
    {
        "pattern": r"(senior\s*auditor)",
        "designation": "Senior Auditor",
        "upgradations": [
            ("2016-07-01", "BPS-16", "Finance Div OM No. 9(7)R-1/2014-515/2016 (01-07-2016)")
        ],
        "historical_initial": [
            ("2016-07-01", "BPS-16"),
            ("1983-01-01", "BPS-11")
        ]
    },
    {
        "pattern": r"(daftary)",
        "designation": "Daftary",
        "upgradations": [
            ("2007-07-01", "BPS-02", "Finance Div OM 2007"),
            ("2016-07-01", "BPS-03", "Finance Div OM 2016")
        ],
        "historical_initial": [
            ("2016-07-01", "BPS-03"),
            ("2007-07-01", "BPS-02"),
            ("1983-01-01", "BPS-01")
        ]
    },
    {
        "pattern": r"(record\s*sorter)",
        "designation": "Record Sorter",
        "upgradations": [
            ("2007-07-01", "BPS-03", "Finance Div OM 2007"),
            ("2016-07-01", "BPS-04", "Finance Div OM 2016")
        ],
        "historical_initial": [
            ("2016-07-01", "BPS-04"),
            ("2007-07-01", "BPS-03"),
            ("1983-01-01", "BPS-02")
        ]
    }
]

def get_automatic_scale_for_civil_post(post_name, effective_date_str):
    if not post_name:
        return None
    p_lower = str(post_name).strip().lower()
    for rule in FEDERAL_CADRE_RULES:
        if re.search(rule["pattern"], p_lower):
            for cut_date, scale_code in rule["historical_initial"]:
                if effective_date_str >= cut_date:
                    return scale_code
    return None

def resolve_civil_servant_cadre_history(personnel, raw_upgradations):
    auto_upgrades = list(raw_upgradations)
    seen_events = {f"{u.get('date')}_{normalize_rank_str(u.get('rank'))}" for u in auto_upgrades}

    doe = personnel.get("doe")
    end_date = personnel.get("retirement_date") if personnel.get("retirement_date") else personnel.get("end_date", "2030-12-31")

    # Build sequential post tenure timeline
    tenures = []
    
    # 1. Enrolment post
    enrol_post = personnel.get("enrol_post", "")
    enrol_scale = personnel.get("enrol_rank", "BPS-01")
    tenures.append({"date": doe, "post": enrol_post, "scale": enrol_scale})

    # 2. Remustering post
    if personnel.get("d_remust"):
        tenures.append({
            "date": personnel.get("d_remust"),
            "post": personnel.get("remust_post", "GD Sepoy"),
            "scale": personnel.get("remust_rank", "BPS-05")
        })

    # 3. Lateral Transfers
    for t in personnel.get("transfers", []):
        if t.get("date"):
            tenures.append({
                "date": t.get("date"),
                "post": t.get("post_designation", "Transferred Post"),
                "scale": t.get("rank", "BPS-05")
            })

    # 4. Promotions
    for p in personnel.get("promotions", []):
        if p.get("date"):
            p_post = p.get("post_designation", "")
            # Automatically assign historical scale if user left default or generic
            auto_sc = get_automatic_scale_for_civil_post(p_post, p.get("date"))
            assigned_scale = auto_sc if auto_sc else p.get("rank", "BPS-07")
            p["rank"] = assigned_scale
            tenures.append({
                "date": p.get("date"),
                "post": p_post,
                "scale": assigned_scale
            })

    # Sort tenures chronologically
    tenures.sort(key=lambda x: str(x["date"]))

    # Scan tenures for automatic cadre upgradations
    for idx, tenure in enumerate(tenures):
        start_dt = str(tenure["date"])
        next_dt = str(tenures[idx+1]["date"]) if idx+1 < len(tenures) else str(end_date)
        post_str = str(tenure["post"]).strip().lower()

        for rule in FEDERAL_CADRE_RULES:
            if re.search(rule["pattern"], post_str):
                for upg_date, upg_scale, upg_auth in rule["upgradations"]:
                    # If upgradation occurred during tenure in this post
                    if start_dt <= upg_date <= next_dt:
                        event_key = f"{upg_date}_{upg_scale}"
                        if event_key not in seen_events:
                            auto_upgrades.append({
                                "date": upg_date,
                                "rank": upg_scale,
                                "designation": rule["designation"],
                                "authority": upg_auth,
                                "remarks": f"Upgraded as {rule['designation']} ({upg_scale})"
                            })
                            seen_events.add(event_key)

    return sorted(auto_upgrades, key=lambda x: str(x["date"]))

# ==============================================================================
# 6. CORE AUDITED CALCULATION ENGINE
# ==============================================================================
def get_scale_at_date(rank, target_date_str, scales_dict, rev_tag=None):
    norm_rk = normalize_rank_str(rank)
    if norm_rk not in scales_dict:
        clean = rank.replace("GD ", "")
        for k in scales_dict:
            if clean.lower() == k.lower() or rank.lower() == k.lower() or norm_rk.lower() == k.lower():
                norm_rk = k
                break
    if norm_rk not in scales_dict:
        return None
    r_scales = scales_dict[norm_rk]
    if rev_tag and rev_tag in r_scales:
        return r_scales[rev_tag]
    target_dt = datetime.datetime.strptime(target_date_str, "%Y-%m-%d").date()
    applicable = None
    for k, sc in sorted(r_scales.items(), key=lambda x: x[0]):
        dt_part = k.split("_")[0]
        if datetime.datetime.strptime(dt_part, "%Y-%m-%d").date() <= target_dt:
            applicable = sc
    return applicable

def run_comprehensive_pay_fixation(personnel, scales_dict, rect_rates, special_instructions=""):
    category = personnel.get("category", "CAF")
    doe = personnel["doe"]
    enrol_rank = normalize_rank_str(personnel.get("enrol_rank", "GD Sep"))
    enrol_post = personnel.get("enrol_post", "GD Rect")
    qualification = personnel.get("qualification", "Matric")
    d_remust = personnel.get("d_remust")
    remust_rank = normalize_rank_str(personnel.get("remust_rank", "GD Sep"))
    promotions = personnel.get("promotions", [])
    transfers = personnel.get("transfers", [])
    raw_upgradations = personnel.get("upgradations", [])
    retirement_date = personnel.get("retirement_date")
    end_date_str = retirement_date if retirement_date else personnel.get("end_date", "2030-12-31")

    # RESTORED AUTOMATIC CADRE UP-GRADATION DISCOVERY (CIVIL SERVANTS ONLY)
    if category == "Civil Servants":
        upgradations = resolve_civil_servant_cadre_history(personnel, raw_upgradations)
    else:
        upgradations = list(raw_upgradations)

    warnings = []
    events = []

    # 1. Enrolment / Appointment
    if category == "Civil Servants" or enrol_rank.startswith("BPS"):
        events.append({
            "date": doe, "type": "ENROLLMENT", "rank": enrol_rank,
            "remarks": "Appointed as GD Recruit Metric" if "recruit" in enrol_post.lower() else f"Appointed as {enrol_post}"
        })
    elif enrol_rank == "NCs E":
        events.append({
            "date": doe, "type": "ENROLLMENT", "rank": "NCs E",
            "remarks": "Enrolled as NCs E"
        })
    else:
        events.append({
            "date": doe, "type": "ENROLLMENT", "rank": "GD Rect",
            "remarks": "Enrolled as GD Rect"
        })

    # 2. Remustering
    if d_remust and d_remust <= end_date_str:
        rem_label = "Remustered as GD Sepoy" if category == "Civil Servants" else "Remustered as train sldr"
        events.append({
            "date": d_remust, "type": "REMUSTERING", "rank": remust_rank,
            "remarks": rem_label
        })

    # 3. Transfers (Civil Servants Only)
    for t in transfers:
        if t.get("date") and t["date"] <= end_date_str:
            t_rk = normalize_rank_str(t.get("rank", enrol_rank))
            t_post = t.get("post_designation", "Transferred Post")
            events.append({
                "date": t["date"], "type": "PAY_PROTECTION", "rank": t_rk,
                "remarks": f"Transferred as {t_post}"
            })

    # 4. Promotions
    for p in promotions:
        if p.get("date") and p["date"] <= end_date_str:
            p_rk = normalize_rank_str(p["rank"])
            p_post = p.get("post_designation", p_rk)
            rem = f"Pmtd {p_post} ({p_rk})" if category == "Civil Servants" else f"Pmtd {p_rk.replace('GD ', '')}"
            events.append({
                "date": p["date"], "type": "PROMOTION", "rank": p_rk,
                "remarks": rem
            })

    # 5. Up-gradations
    for u in upgradations:
        if u.get("date") and u["date"] <= end_date_str:
            u_rk = normalize_rank_str(u["rank"])
