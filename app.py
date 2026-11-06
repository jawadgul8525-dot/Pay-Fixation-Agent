import streamlit as st
import datetime
import io
import os
import re
import json
import base64
import pandas as pd
from PIL import Image
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

REGISTRY_SCHEMA_VERSION = "2026.11.v11_PARITY_UPGRADE_AND_BRANDED_PWA"

ALL_BPS_SCALES = [f"BPS-{i:02d}" for i in range(1, 23)]
NCSE_SCALES = ["NCs E (BPS-01)"] + [f"BPS-{i:02d}" for i in range(2, 12)]

CAF_ALL_RANKS = [
    "GD Sep", "Tech Sep",
    "GD Lnk", "Tech Lnk",
    "GD Nk", "Tech Nk",
    "GD Hav", "Tech Hav",
    "GD N/Sub", "Tech N/Sub",
    "Sub", "SM", "H/Lt", "H/Capt"
]

DATA_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")
REGISTRY_FILE = os.path.join(DATA_DIR, "chart_registry.json")
try:
    os.makedirs(DATA_DIR, exist_ok=True)
except Exception:
    pass

# ==============================================================================
# 0. ROBUST LOGO DISCOVERY & PWA BRANDING INJECTOR
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
    return None

def get_logo_base64():
    p = get_logo_path()
    if p and os.path.exists(p):
        try:
            with open(p, "rb") as f:
                return base64.b64encode(f.read()).decode("utf-8")
        except Exception:
            return ""
    return ""

def inject_pwa_branding():
    b64_img = get_logo_base64()
    if not b64_img:
        return
    
    data_uri = f"data:image/png;base64,{b64_img}"
    manifest_data = {
        "name": "Pay Fixation Agent",
        "short_name": "Pay Fixation",
        "description": "Government of Pakistan Basic Pay Fixation System (Civil Servants & CAF)",
        "start_url": ".",
        "display": "standalone",
        "background_color": "#0E1117",
        "theme_color": "#0E1117",
        "icons": [
            {
                "src": data_uri,
                "sizes": "192x192 512x512",
                "type": "image/png",
                "purpose": "any maskable"
            }
        ]
    }
    manifest_b64 = base64.b64encode(json.dumps(manifest_data).encode("utf-8")).decode("utf-8")
    manifest_uri = f"data:application/manifest+json;base64,{manifest_b64}"

    # Injects the custom favicon, apple-touch-icon, and PWA manifest directly into the DOM
    pwa_script = f"""
    <script>
        (function() {{
            const logoUri = "{data_uri}";
            const manifestUri = "{manifest_uri}";

            // 1. Browser Favicon
            let favicon = document.querySelector("link[rel*='icon']");
            if (!favicon) {{
                favicon = document.createElement("link");
                favicon.rel = "shortcut icon";
                document.head.appendChild(favicon);
            }}
            favicon.href = logoUri;

            // 2. Apple Touch Icon (iOS Home Screen / PWA Bookmark)
            let appleIcon = document.querySelector("link[rel='apple-touch-icon']");
            if (!appleIcon) {{
                appleIcon = document.createElement("link");
                appleIcon.rel = "apple-touch-icon";
                document.head.appendChild(appleIcon);
            }}
            appleIcon.href = logoUri;

            // 3. Web App Manifest (Android / Desktop Chrome PWA Installed Icon)
            let manifest = document.querySelector("link[rel='manifest']");
            if (!manifest) {{
                manifest = document.createElement("link");
                manifest.rel = "manifest";
                document.head.appendChild(manifest);
            }}
            manifest.href = manifestUri;
        }})();
    </script>
    """
    st.components.v1.html(pwa_script, height=0, width=0)

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
# 2. VERIFIED CAF SCALES (INCLUDING 2026 TECH HAV AMENDMENT: 25880-1810-80180)
# ==============================================================================
BOOTSTRAP_CAF_SCALES = {
    "NCs E": {
        "1994-06-01":(1245,35,1770),"2001-12-01":(1870,55,3520),"2005-07-01_REV":(2130,65,4100),"2005-07-01_UPG":(2130,65,4100),"2007-07-01":(2475,75,4725),"2008-07-01":(2970,90,5670),"2011-07-01":(4800,150,9300),"2015-07-01":(6210,195,12060),"2016-07-01":(7610,240,14840),"2017-07-01":(9125,285,17675),"2022-07-01":(13540,430,26440),"2024-06-01":(15620,490,30320),"2026-07-01":(18780,590,36480)
    },
    "GD Sep": {
        "1994-06-01":(1254,30,1800),"2001-12-01":(1875,75,3525),"2005-07-01_REV":(2155,85,4025),"2005-07-01_UPG":(2415,115,5865),"2007-07-01":(2780,135,6830),"2008-07-01":(3340,160,8140),"2011-07-01":(5400,270,13500),"2015-07-01":(6985,350,17485),"2016-07-01":(8590,430,21490),"2017-07-01":(10260,510,25560),"2022-07-01":(15230,760,38030),"2024-06-01":(18160,950,46660),"2026-07-01":(21830,1150,56330)
    },
    "Tech Sep": {
        "1994-06-01":(1262,61,1933),"2001-12-01":(1895,95,3985),"2005-07-01_REV":(2180,110,4600),"2005-07-01_UPG":(2415,115,5865),"2007-07-01":(2780,135,6830),"2008-07-01":(3340,160,8140),"2011-07-01":(5400,270,13500),"2015-07-01":(6985,350,17485),"2016-07-01":(8590,430,21490),"2017-07-01":(10260,510,25560),"2022-07-01":(15230,760,38030),"2024-06-01":(18160,950,46660),"2026-07-01":(21830,1150,56330)
    },
    "GD Lnk": {
        "1994-06-01":(1262,61,1933),"2001-12-01":(1895,95,3985),"2005-07-01_REV":(2180,110,4600),"2005-07-01_UPG":(2485,125,6235),"2007-07-01":(2860,145,7210),"2008-07-01":(3430,175,8680),"2011-07-01":(5500,290,14200),"2015-07-01":(7115,375,18365),"2016-07-01":(8750,470,22850),"2017-07-01":(10465,560,27265),"2022-07-01":(15530,840,40730),"2024-06-01":(18780,1040,49980),"2026-07-01":(22580,1250,60080)
    },
    "Tech Lnk": {
        "1994-06-01":(1400,66,2390),"2001-12-01":(2400,100,5100),"2005-07-01_REV":(2415,115,5865),"2005-07-01_UPG":(2485,125,6235),"2007-07-01":(2860,145,7210),"2008-07-01":(3430,175,8680),"2011-07-01":(5500,290,14200),"2015-07-01":(7115,375,18365),"2016-07-01":(8750,470,22850),"2017-07-01":(10465,560,27265),"2022-07-01":(15530,840,40730),"2024-06-01":(18780,1040,49980),"2026-07-01":(22580,1250,60080)
    },
    "GD Nk": {
        "1994-06-01":(1400,66,2390),"2001-12-01":(2400,100,5100),"2005-07-01_REV":(2415,115,5865),"2005-07-01_UPG":(2555,140,6755),"2007-07-01":(2940,160,7740),"2008-07-01":(3530,190,9230),"2011-07-01":(5700,310,15000),"2015-07-01":(7370,405,19520),"2016-07-01":(9070,500,24070),"2017-07-01":(10835,595,28665),"2022-07-01":(16080,890,42780),"2024-06-01":(19450,1140,53650),"2026-07-01":(23380,1380,64780)
    },
    "Tech Nk": {
        "1994-06-01":(1440,73,2535),"2001-12-01":(2160,110,5460),"2005-07-01_REV":(2485,125,6235),"2005-07-01_UPG":(2555,140,6755),"2007-07-01":(2940,160,7740),"2008-07-01":(3530,190,9230),"2011-07-01":(5700,310,15000),"2015-07-01":(7370,405,19520),"2016-07-01":(9070,500,24070),"2017-07-01":(10835,595,28665),"2022-07-01":(16080,890,42780),"2024-06-01":(19450,1140,53650),"2026-07-01":(23380,1380,64780)
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
    "Tech N/Sub": {
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
    try:
        with open(REGISTRY_FILE, "w") as f:
            json.dump(reg, f, indent=2)
    except Exception:
        pass

def save_category_data(category, scales, rect_rates, filename):
    try:
        path = os.path.join(DATA_DIR, filename)
        with open(path, "w") as f:
            json.dump({"scales": scales, "rect_rates": rect_rates}, f, indent=2)
    except Exception:
        pass

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
    if rk == "NCs E (BPS-01)":
        return "NCs E"
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
# 5. RESTORED AUTOMATIC CADRE POST UP-GRADATION ENGINE (CIVIL SERVANTS ONLY)
# ==============================================================================
FEDERAL_CADRE_RULES = [
    {
        "keywords": ["junior clerk", "ldc", "lower division clerk", "jc"],
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
        "keywords": ["senior clerk", "udc", "upper division clerk", "sc"],
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
        "keywords": ["assistant", "head clerk"],
        "exclude": ["private"],
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
        "keywords": ["superintendent"],
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
        "keywords": ["stenotypist"],
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
        "keywords": ["stenographer", "assistant private secretary", "aps"],
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
        "keywords": ["junior auditor"],
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
        "keywords": ["senior auditor"],
        "designation": "Senior Auditor",
        "upgradations": [
            ("2016-07-01", "BPS-16", "Finance Div OM No. 9(7)R-1/2014-515/2016 (01-07-2016)")
        ],
        "historical_initial": [
            ("2016-07-01", "BPS-16"),
            ("1983-01-01", "BPS-11")
        ]
    }
]

def matches_cadre_rule(post_name, rule):
    p_lower = str(post_name).strip().lower()
    if not p_lower:
        return False
    if "exclude" in rule:
        for ex in rule["exclude"]:
            if ex in p_lower:
                return False
    for kw in rule["keywords"]:
        if kw in p_lower:
            return True
    return False

def get_automatic_scale_for_civil_post(post_name, effective_date_str):
    if not post_name:
        return None
    for rule in FEDERAL_CADRE_RULES:
        if matches_cadre_rule(post_name, rule):
            for cut_date, scale_code in rule["historical_initial"]:
                if effective_date_str >= cut_date:
                    return scale_code
    return None

def resolve_civil_servant_cadre_history(personnel, raw_upgradations):
    auto_upgrades = list(raw_upgradations)
    seen_events = {f"{u.get('date')}_{normalize_rank_str(u.get('rank'))}" for u in auto_upgrades}

    doe = str(personnel.get("doe", "1990-01-01"))
    end_date = str(personnel.get("retirement_date")) if personnel.get("retirement_date") else str(personnel.get("end_date", "2030-12-31"))

    tenures = []
    enrol_post = personnel.get("enrol_post", "")
    enrol_scale = personnel.get("enrol_rank", "BPS-01")
    tenures.append({"date": doe, "post": enrol_post, "scale": enrol_scale})

    if personnel.get("d_remust"):
        tenures.append({
            "date": str(personnel.get("d_remust")),
            "post": personnel.get("remust_post", "GD Sepoy"),
            "scale": personnel.get("remust_rank", "BPS-05")
        })

    for t in personnel.get("transfers", []):
        if t.get("date"):
            tenures.append({
                "date": str(t.get("date")),
                "post": t.get("post_designation", "Transferred Post"),
                "scale": t.get("rank", "BPS-05")
            })

    for p in personnel.get("promotions", []):
        if p.get("date"):
            p_post = p.get("post_designation", "")
            auto_sc = get_automatic_scale_for_civil_post(p_post, str(p.get("date")))
            assigned_scale = auto_sc if auto_sc else p.get("rank", "BPS-07")
            tenures.append({
                "date": str(p.get("date")),
                "post": p_post,
                "scale": assigned_scale
            })

    tenures.sort(key=lambda x: str(x["date"]))

    for idx, tenure in enumerate(tenures):
        start_dt = str(tenure["date"])
        next_dt = str(tenures[idx+1]["date"]) if idx+1 < len(tenures) else end_date
        post_str = tenure.get("post", "")

        for rule in FEDERAL_CADRE_RULES:
            if matches_cadre_rule(post_str, rule):
                for upg_date, upg_scale, upg_auth in rule["upgradations"]:
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
    if norm_rk.startswith("BPS-") and norm_rk not in scales_dict:
        # Check Civil Scales directly for NCsE BPS-02 through BPS-11
        if norm_rk in DEFAULT_CIVIL_SCALES:
            r_scales = DEFAULT_CIVIL_SCALES[norm_rk]
            target_dt = datetime.datetime.strptime(target_date_str, "%Y-%m-%d").date()
            applicable = None
            for k, sc in sorted(r_scales.items(), key=lambda x: x[0]):
                dt_part = k.split("_")[0]
                if datetime.datetime.strptime(dt_part, "%Y-%m-%d").date() <= target_dt:
                    applicable = sc
            return applicable

    if norm_rk not in scales_dict:
        clean = rank.replace("GD ", "").replace("Tech ", "")
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
            "remarks": f"Enrolled as NCs E ({personnel.get('ncse_init_scale', 'NCs E')})"
        })
    else:
        events.append({
            "date": doe, "type": "ENROLLMENT", "rank": "GD Rect",
            "remarks": f"Enrolled as GD Rect"
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

    # 4. Promotions / Technical Appointments
    for p in promotions:
        if p.get("date") and p["date"] <= end_date_str:
            p_rk = normalize_rank_str(p["rank"])
            p_post = p.get("post_designation", p_rk)
            rem = f"Pmtd {p_post} ({p_rk})" if category == "Civil Servants" else f"Pmtd {p_rk.replace('GD ', '')}"
            events.append({
                "date": p["date"], "type": "PROMOTION", "rank": p_rk,
                "remarks": rem
            })

    # 5. Up-gradations (NCsE Scale Movements / Cadre Up-gradations)
    for u in upgradations:
        if u.get("date") and u["date"] <= end_date_str:
            u_rk = normalize_rank_str(u["rank"])
            events.append({
                "date": u["date"], "type": "UPGRADATION", "rank": u_rk,
                "remarks": u.get("remarks", f"Upgraded to {u_rk}")
            })

    # 6. Revisions
    active_revisions = set()
    for rk, dt_dict in scales_dict.items():
        for d_key in dt_dict.keys():
            active_revisions.add(d_key)

    for r_key in sorted(list(active_revisions)):
        r_dt = r_key.split("_")[0]
        r_type = "UPGRADATION" if "UPG" in r_key else "REVISION"
        rem_str = "Upgradation" if "UPG" in r_key else "Revision"
        if doe <= r_dt <= end_date_str:
            events.append({
                "date": r_dt, "type": r_type, "rev_key": r_key,
                "remarks": rem_str
            })

    # 7. Annual Increments
    start_year = datetime.datetime.strptime(doe, "%Y-%m-%d").year
    end_year = datetime.datetime.strptime(end_date_str, "%Y-%m-%d").year
    for y in range(start_year, end_year + 1):
        inc_dt = f"{y}-12-01"
        if doe <= inc_dt <= end_date_str:
            events.append({
                "date": inc_dt, "type": "ANNUAL_INCREMENT",
                "remarks": "A/Inc"
            })

    # 8. Retirement Event
    if retirement_date and retirement_date <= end_date_str:
        ret_dt = datetime.datetime.strptime(retirement_date, "%Y-%m-%d").date()
        ret_mon = ret_dt.month
        entitled_ret_inc = (6 <= ret_mon <= 11)
        events.append({
            "date": retirement_date, "type": "RETIREMENT",
            "entitled_increment": entitled_ret_inc,
            "remarks": "Retirement"
        })

    # Sequence priorities: Revisions run before Upgradations on identical dates
    priority = {
        "ENROLLMENT": 0, "REMUSTERING": 1, "PAY_PROTECTION": 2, "REVISION": 3,
        "UPGRADATION": 4, "PROMOTION": 5, "ANNUAL_INCREMENT": 6, "RETIREMENT": 7
    }
    events.sort(key=lambda x: (x["date"], priority.get(x["type"], 9)))

    ledger = []
    current_pay = 0
    current_rank = enrol_rank
    current_scale = None
    substantive_scale = None
    last_change_dt = None
    personal_pay = 0

    for ev in events:
        ev_dt = datetime.datetime.strptime(ev["date"], "%Y-%m-%d").date()
        ev_t = ev["type"]

        if ev_t == "ENROLLMENT":
            if category == "Civil Servants" or enrol_rank.startswith("BPS") or enrol_rank == "NCs E":
                init_sc_key = personnel.get("ncse_init_scale", enrol_rank)
                current_scale = get_scale_at_date(init_sc_key, ev["date"], scales_dict)
                substantive_scale = current_scale
                if "recruit" in enrol_post.lower() or "rect" in enrol_post.lower():
                    rates = rect_rates.get(qualification, [])
                    rect_pay = next((r[1] for r in reversed(rates) if r[0] <= ev["date"]), None)
                    current_pay = rect_pay if rect_pay else (current_scale[0] if current_scale else 2128)
                    scale_lbl = str(current_pay)
                else:
                    current_pay = current_scale[0] if current_scale else 0
                    scale_lbl = format_scale_str(current_scale)
                last_change_dt = ev_dt
                ledger.append({
                    "date": ev_dt.strftime("%d-%m-%Y"), "pay": current_pay,
                    "scale": scale_lbl, "remarks": ev.get("remarks", "Appointed")
                })
            else:
                rates = rect_rates.get(qualification, [])
                rect_pay = next((r[1] for r in reversed(rates) if r[0] <= ev["date"]), None)
                current_scale = get_scale_at_date(enrol_rank, ev["date"], scales_dict)
                substantive_scale = current_scale
                current_pay = rect_pay if rect_pay else (current_scale[0] if current_scale else 2128)
                last_change_dt = ev_dt
                ledger.append({
                    "date": ev_dt.strftime("%d-%m-%Y"), "pay": current_pay,
                    "scale": "", "remarks": ev.get("remarks", "Enrolled as GD Rect")
                })

        elif ev_t == "REMUSTERING":
            current_rank = ev["rank"]
            current_scale = get_scale_at_date(current_rank, ev["date"], scales_dict)
            substantive_scale = current_scale
            current_pay = current_scale[0] if current_scale else "Review Req"
            last_change_dt = ev_dt
            ledger.append({
                "date": ev_dt.strftime("%d-%m-%Y"), "pay": current_pay,
                "scale": format_scale_str(current_scale), "remarks": ev.get("remarks", "Remustered as train sldr")
            })

        elif ev_t == "PAY_PROTECTION":
            target_rank = ev["rank"]
            new_sc = get_scale_at_date(target_rank, ev["date"], scales_dict)
            current_rank = target_rank
            if new_sc:
                current_scale = new_sc
                substantive_scale = new_sc
            last_change_dt = ev_dt
            ledger.append({
                "date": ev_dt.strftime("%d-%m-%Y"), "pay": current_pay,
                "scale": format_scale_str(current_scale) if current_scale else '"',
                "remarks": ev.get("remarks", f"Transferred as {target_rank}")
            })

        elif ev_t == "ANNUAL_INCREMENT":
            if not isinstance(current_pay, (int, float)):
                ledger.append({"date": ev_dt.strftime("%d-%m-%Y"), "pay": current_pay, "scale": '"', "remarks": "HOLD — Scale Unresolved"})
                continue
            days_served = (ev_dt - last_change_dt).days if last_change_dt else 365
            if days_served < 180:
                ledger.append({
                    "date": ev_dt.strftime("%d-%m-%Y"), "pay": current_pay, "scale": '"',
                    "remarks": "Not entitled"
                })
            else:
                active_inc = substantive_scale[1] if substantive_scale else (current_scale[1] if current_scale else 0)
                if current_scale and current_pay + active_inc > current_scale[2]:
                    personal_pay += active_inc
                    ledger.append({
                        "date": ev_dt.strftime("%d-%m-%Y"), "pay": current_pay, "scale": '"',
                        "remarks": f"At Ceiling (Rs. {personal_pay} PP)"
                    })
                else:
                    current_pay += active_inc
                    ledger.append({
                        "date": ev_dt.strftime("%d-%m-%Y"), "pay": current_pay, "scale": '"',
                        "remarks": "A/Inc"
                    })

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
                ledger.append({
                    "date": ev_dt.strftime("%d-%m-%Y"), "pay": current_pay,
                    "scale": format_scale_str(new_sc) if new_sc else "Missing Scale",
                    "remarks": "Revision"
                })

        # ======================================================================
        # CORRECTION 1: PARITY FOR UP-GRADATION BENEFIT (CIVIL & CAF FROM 28-01-2002)
        # ======================================================================
        elif ev_t == "UPGRADATION":
            target_rank = ev.get("rank", current_rank)
            new_sc = get_scale_at_date(target_rank, ev["date"], scales_dict, rev_tag=ev.get("rev_key"))
            if new_sc and isinstance(current_pay, (int, float)):
                p_step = new_sc[0]
                while p_step < current_pay:
                    p_step += new_sc[1]

                # Finance Division OM dated 31-05-2013 and Court Judgements:
                # Every qualifying post upgradation is treated equivalent to promotion.
                # Both Civil Servants and CAF personnel receive one additional increment
                # from 28-01-2002 onward (including retrospective cases such as the 2007 CAF upgradation).
                if ev["date"] >= "2002-01-28":
                    current_pay = p_step + new_sc[1]
                else:
                    current_pay = p_step

                current_scale = new_sc
                substantive_scale = new_sc
                current_rank = target_rank
                last_change_dt = ev_dt
                remarks_text = ev.get("remarks") if ev.get("remarks") else f"Upgraded to {current_rank}"
                ledger.append({
                    "date": ev_dt.strftime("%d-%m-%Y"), "pay": current_pay,
                    "scale": format_scale_str(current_scale), "remarks": remarks_text
                })
            else:
                ledger.append({
                    "date": ev_dt.strftime("%d-%m-%Y"), "pay": current_pay,
                    "scale": format_scale_str(new_sc) if new_sc else "Missing Scale",
                    "remarks": "Upgradation"
                })

        # PROMOTION / TECHNICAL PROMOTION FIXATION (FR 22(a)(i))
        elif ev_t == "PROMOTION":
            target_rank = ev.get("rank", current_rank)
            if target_rank in ["H/Lt", "H/Capt"]:
                sm_scale = get_scale_at_date("SM", ev["date"], scales_dict)
                honorary_inc = sm_scale[1] if sm_scale else 0
                if isinstance(current_pay, (int, float)):
                    current_pay += honorary_inc
                last_change_dt = ev_dt
                ledger.append({
                    "date": ev_dt.strftime("%d-%m-%Y"), "pay": current_pay,
                    "scale": format_scale_str(substantive_scale) if substantive_scale else '"',
                    "remarks": f"1 x adv inc on prom as {target_rank}"
                })
            else:
                current_rank = target_rank
                new_sc = get_scale_at_date(current_rank, ev["date"], scales_dict)
                if not new_sc:
                    warnings.append(f"HOLD: Scale missing for {current_rank} on {ev['date']}.")
                    current_pay = "Review Req"
                else:
                    if isinstance(current_pay, (int, float)):
                        p_step = new_sc[0]
                        while p_step < current_pay:
                            p_step += new_sc[1]
                        current_pay = p_step + new_sc[1] # Step-up + 1 Promotional Increment
                    current_scale = new_sc
                    substantive_scale = new_sc
                last_change_dt = ev_dt
                remarks_str = ev.get("remarks") if ev.get("remarks") else f"Pmtd {current_rank.replace('GD ', '')}"
                ledger.append({
                    "date": ev_dt.strftime("%d-%m-%Y"), "pay": current_pay,
                    "scale": format_scale_str(current_scale), "remarks": remarks_str
                })

        elif ev_t == "RETIREMENT":
            final_pension_pay = current_pay
            if ev.get("entitled_increment") and isinstance(current_pay, (int, float)):
                ret_inc = current_scale[1] if current_scale else 0
                final_pension_pay += ret_inc
                ledger.append({
                    "date": ev_dt.strftime("%d-%m-%Y"), "pay": final_pension_pay,
                    "scale": format_scale_str(current_scale),
                    "remarks": "Usual Increment on retirement"
                })
            else:
                ledger.append({
                    "date": ev_dt.strftime("%d-%m-%Y"), "pay": current_pay,
                    "scale": format_scale_str(current_scale),
                    "remarks": "Retirement"
                })

    return {
        "header": personnel, "ledger": ledger,
        "special_instructions": special_instructions, "warnings": warnings,
        "personal_pay": personal_pay, "discovered_upgradations": upgradations
    }

# ==============================================================================
# 7. "ILYAS SB" REPORT GENERATORS (EXCEL, WORD, PDF)
# ==============================================================================
def get_official_heading(h):
    fc_no = str(h.get('fc_no', '')).strip()
    rank = str(h.get('rank', '')).strip()
    name = str(h.get('name', '')).strip().upper()
    return f"PAY FIXATION IRO OF FC / PERSONAL {fc_no}, Rank/Designation {rank} and Name {name}"

def build_particulars_rows(h):
    rows = []
    is_caf = (h.get("category") == "CAF")
    lbl_enrol = "D.O E" if is_caf else f"D.O.Appointment as {h.get('enrol_post', 'GD Recruit')}"
    rows.append(("a", lbl_enrol, to_display_date(h.get("doe"))))

    marker_idx = 1
    letters = ["b.", "c.", "d.", "e.", "f.", "g.", "h.", "i.", "j.", "k.", "l.", "m.", "n.", "o."]

    if h.get("d_remust"):
        m = letters[marker_idx] if marker_idx < len(letters) else f"{marker_idx+1}."
        lbl_rem = "D.O.Remustration as GD Sepoy" if not is_caf else "D.O.Remustration"
        rows.append((m, lbl_rem, to_display_date(h.get("d_remust"))))
        marker_idx += 1

    for t in h.get("transfers", []):
        t_p = t.get("post_designation", "Transferred Post")
        t_dt = to_display_date(t.get("date", ""))
        m = letters[marker_idx] if marker_idx < len(letters) else f"{marker_idx+1}."
        rows.append((m, f"D.O.Transfer as {t_p}", t_dt))
        marker_idx += 1

    timeline_events = []
    for p in h.get("promotions", []):
        p_post = p.get("post_designation", str(p.get("rank", "")).replace("GD ", ""))
        p_rk = str(p.get("rank", "")).replace("GD ", "")
        lbl = f"D.O.Promotion as {p_post} ({p_rk})" if p_post and p_post != p_rk else f"D.O.Promotion as {p_rk}"
        timeline_events.append((p.get("date"), lbl))

    for u in h.get("upgradations", []):
        u_post = u.get("designation", u.get("rank", "Upgraded Post"))
        u_rk = str(u.get("rank", "")).replace("GD ", "")
        timeline_events.append((u.get("date"), f"D.O.Upgradation as {u_post} ({u_rk})"))

    timeline_events.sort(key=lambda x: str(x[0]))
    for ev_dt, ev_lbl in timeline_events:
        m = letters[marker_idx] if marker_idx < len(letters) else f"{marker_idx+1}."
        rows.append((m, ev_lbl, to_display_date(ev_dt)))
        marker_idx += 1

    if h.get("retirement_date"):
        m = letters[marker_idx] if marker_idx < len(letters) else f"{marker_idx+1}."
        rows.append((m, "D.O.Retirement", to_display_date(h.get("retirement_date"))))
        marker_idx += 1

    return rows

def generate_excel_bytes(data):
    wb = Workbook()
    ws = wb.active
    ws.title = "PAY FIXATION"
    ws.views.sheetView[0].showGridLines = True

    h = data["header"]
    heading_text = get_official_heading(h)

    ws.merge_cells("A1:E1")
    cell_h = ws["A1"]
    cell_h.value = heading_text
    cell_h.font = Font(name="Calibri", size=11, bold=True)
    cell_h.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 24

    particulars = build_particulars_rows(h)
    curr_r = 3
    for marker, milestone, dt_val in particulars:
        c_m = ws.cell(row=curr_r, column=1, value=marker)
        c_m.font = Font(name="Calibri", size=10, bold=True)
        c_m.alignment = Alignment(horizontal="center", vertical="center")

        c_lbl = ws.cell(row=curr_r, column=2, value=milestone)
        c_lbl.font = Font(name="Calibri", size=10)
        c_lbl.alignment = Alignment(horizontal="left", vertical="center")

        ws.cell(row=curr_r, column=3, value="")

        c_dt = ws.cell(row=curr_r, column=4, value=dt_val)
        c_dt.font = Font(name="Calibri", size=10, bold=True)
        c_dt.alignment = Alignment(horizontal="center", vertical="center")
        curr_r += 1

    if data.get("special_instructions"):
        curr_r += 1
        ws.cell(row=curr_r, column=1, value="Note:").font = Font(bold=True)
        ws.merge_cells(start_row=curr_r, start_column=2, end_row=curr_r, end_column=5)
        ws.cell(row=curr_r, column=2, value=data["special_instructions"]).font = Font(italic=True, color="000080")

    curr_r += 2
    cols = ["Ser", "Date", "Pay Fixed", "Pay Scale", "Remarks"]
    for ci, col in enumerate(cols, start=1):
        cell = ws.cell(row=curr_r, column=ci, value=col)
        cell.font = Font(name="Calibri", size=10, bold=True)
        cell.fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[curr_r].height = 20

    border_thin = Border(
        left=Side(style='thin', color='A6A6A6'), right=Side(style='thin', color='A6A6A6'),
        top=Side(style='thin', color='A6A6A6'), bottom=Side(style='thin', color='A6A6A6')
    )

    curr_r += 1
    for idx, row in enumerate(data["ledger"], start=1):
        c1 = ws.cell(row=curr_r, column=1, value=idx)
        c2 = ws.cell(row=curr_r, column=2, value=row["date"])
        c3 = ws.cell(row=curr_r, column=3, value=row["pay"])
        c4 = ws.cell(row=curr_r, column=4, value=row["scale"])
        c5 = ws.cell(row=curr_r, column=5, value=row["remarks"])

        c1.alignment = Alignment(horizontal="center", vertical="center")
        c2.alignment = Alignment(horizontal="center", vertical="center")
        c3.alignment = Alignment(horizontal="center", vertical="center")
        c4.alignment = Alignment(horizontal="center", vertical="center")
        c5.alignment = Alignment(horizontal="left", vertical="center")

        for c in [c1, c2, c3, c4, c5]:
            c.font = Font(name="Calibri", size=10)
            c.border = border_thin
        curr_r += 1

    ws.column_dimensions['A'].width = 8
    ws.column_dimensions['B'].width = 14
    ws.column_dimensions['C'].width = 14
    ws.column_dimensions['D'].width = 24
    ws.column_dimensions['E'].width = 38

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

def generate_word_bytes(data):
    doc = Document()
    h = data["header"]
    heading_text = get_official_heading(h)

    logo_path = get_logo_path()
    if logo_path:
        try:
            p_logo = doc.add_paragraph()
            p_logo.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_logo.add_run().add_picture(logo_path, width=Inches(1.0))
        except Exception:
            pass

    p_title = doc.add_paragraph()
    run = p_title.add_run(heading_text)
    run.font.name = "Calibri"
    run.font.bold = True
    run.font.size = Pt(11)
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    particulars = build_particulars_rows(h)
    p_table = doc.add_table(rows=0, cols=3)
    p_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for marker, milestone, dt_val in particulars:
        r_cells = p_table.add_row().cells
        r_cells[0].text = marker
        r_cells[1].text = milestone
        r_cells[2].text = dt_val

        r_cells[0].paragraphs[0].runs[0].font.name = "Calibri"
        r_cells[0].paragraphs[0].runs[0].font.size = Pt(9.5)
        r_cells[0].paragraphs[0].runs[0].font.bold = True

        r_cells[1].paragraphs[0].runs[0].font.name = "Calibri"
        r_cells[1].paragraphs[0].runs[0].font.size = Pt(9.5)

        r_cells[2].paragraphs[0].runs[0].font.name = "Calibri"
        r_cells[2].paragraphs[0].runs[0].font.size = Pt(9.5)
        r_cells[2].paragraphs[0].runs[0].font.bold = True

        r_cells[0].width = Inches(0.5)
        r_cells[1].width = Inches(3.2)
        r_cells[2].width = Inches(1.8)

    if data.get("special_instructions"):
        p_note = doc.add_paragraph()
        r_note = p_note.add_run(f"\nNote: {data['special_instructions']}")
        r_note.font.italic = True
        r_note.font.size = Pt(9)
        r_note.font.color.rgb = RGBColor(0, 32, 96)

    doc.add_paragraph().text = ""

    table = doc.add_table(rows=1, cols=5)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = 'Table Grid'
    hdr_cells = table.rows[0].cells
    cols = ["Ser", "Date", "Pay Fixed", "Pay Scale", "Remarks"]
    for i, title in enumerate(cols):
        hdr_cells[i].text = title
        hdr_cells[i].paragraphs[0].runs[0].font.name = "Calibri"
        hdr_cells[i].paragraphs[0].runs[0].font.bold = True
        hdr_cells[i].paragraphs[0].runs[0].font.size = Pt(9.5)
        shd = parse_xml(r'<w:shd {} w:fill="D9E1F2"/>'.format(nsdecls('w')))
        hdr_cells[i]._tc.get_or_add_tcPr().append(shd)

    for idx, row in enumerate(data["ledger"], start=1):
        row_cells = table.add_row().cells
        row_cells[0].text = str(idx)
        row_cells[1].text = str(row["date"])
        row_cells[2].text = str(row["pay"])
        row_cells[3].text = str(row["scale"])
        row_cells[4].text = str(row["remarks"])
        for c in row_cells:
            c.paragraphs[0].runs[0].font.name = "Calibri"
            c.paragraphs[0].runs[0].font.size = Pt(9.0)

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()

def generate_pdf_bytes(data):
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=letter, leftMargin=36, rightMargin=36, topMargin=36, bottomMargin=36)
    styles = getSampleStyleSheet()
    elements = []

    logo_path = get_logo_path()
    if logo_path:
        try:
            elements.append(RLImage(logo_path, width=65, height=65))
            elements.append(Spacer(1, 4))
        except Exception:
            pass

    h = data["header"]
    heading_text = f"<b>{get_official_heading(h)}</b>"
    elements.append(Paragraph(heading_text, styles['Title']))
    elements.append(Spacer(1, 8))

    particulars = build_particulars_rows(h)
    p_data = [[m, lbl, dt_val] for m, lbl, dt_val in particulars]
    t_part = Table(p_data, colWidths=[25, 260, 140])
    t_part.setStyle(TableStyle([
        ('FONTNAME', (0,0), (-1,-1), 'Helvetica'),
        ('FONTSIZE', (0,0), (-1,-1), 8.5),
        ('FONTNAME', (0,0), (0,-1), 'Helvetica-Bold'),
        ('FONTNAME', (2,0), (2,-1), 'Helvetica-Bold'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('BOTTOMPADDING', (0,0), (-1,-1), 2),
    ]))
    elements.append(t_part)

    if data.get("special_instructions"):
        elements.append(Spacer(1, 4))
        elements.append(Paragraph(f"<i>Note: {data['special_instructions']}</i>", styles['Normal']))

    elements.append(Spacer(1, 10))

    tbl_data = [["Ser", "Date", "Pay Fixed", "Pay Scale", "Remarks"]]
    for idx, row in enumerate(data["ledger"], start=1):
        tbl_data.append([str(idx), str(row["date"]), str(row["pay"]), str(row["scale"]), str(row["remarks"])])

    t_main = Table(tbl_data, colWidths=[35, 75, 75, 135, 200])
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
# 8. STREAMLIT APPLICATION INTERFACE
# ==============================================================================
def main():
    # --------------------------------------------------------------------------
    # CORRECTION 2: OFFICIAL APP ICON FOR INSTALLED PWA & BROWSER FAVICON
    # --------------------------------------------------------------------------
    logo_file = get_logo_path()
    page_icon_obj = "🎖️"
    if logo_file:
        try:
            page_icon_obj = Image.open(logo_file)
        except Exception:
            page_icon_obj = "🎖️"

    st.set_page_config(page_title="Pay Fixation Agent", layout="wide", page_icon=page_icon_obj)

    # Injects Web App Manifest & Apple Touch Icon so installed/downloaded PWA displays official logo
    inject_pwa_branding()

    h_col1, h_col2 = st.columns([1, 6])
    with h_col1:
        if logo_file:
            try:
                st.image(logo_file, width=95)
            except Exception:
                st.markdown("## 🎖️")
        else:
            st.markdown("## 🎖️")
    with h_col2:
        st.title("Pay Fixation Agent")
        st.caption("Developed by **SC Rehman Gul** | Government of Pakistan Basic Pay Fixation System (Civil Servants & CAF)")

    registry = load_registry()

    st.markdown("---")
    app_mode = st.radio(
        "Select Pay Fixation Cadre:",
        ["Civil Armed Forces (CAF / Frontier Corps)", "Federal Civil Servants (BPS-01 to BPS-22)"],
        horizontal=True
    )

    is_caf = (app_mode == "Civil Armed Forces (CAF / Frontier Corps)")
    target_category = "CAF" if is_caf else "Civil Servants"
    cat_meta = registry[target_category]["active"]

    # Chart Vault & Live Uploader
    with st.expander(f"📁 Pay Scale Chart Vault & Live Uploader — [{target_category}]", expanded=False):
        c_v1, c_v2 = st.columns([2, 3])
        with c_v1:
            st.markdown(f"**Active Standard:** `{cat_meta['filename']}`")
            st.markdown(f"**Revision Epochs:** `{cat_meta['effective_years']}`")
            st.markdown(f"**Scales Covered:** {len(cat_meta['ranks'])}")
        with c_v2:
            up_chart = st.file_uploader(
                f"Upload New / Revised {target_category} Chart (.xlsx, .csv)",
                type=["xlsx", "xls", "csv"],
                key="main_chart_uploader"
            )
            if up_chart is not None:
                parsed_res, err = parse_uploaded_chart(up_chart)
                if err:
                    st.error(f"❌ {err}")
                else:
                    st.success(f"✅ Readable: {len(parsed_res['ranks'])} scales detected.")
                    if st.button(f"Activate as New {target_category} Chart", type="primary", key="btn_act_main"):
                        registry[target_category]["backup"] = cat_meta
                        new_data_file = f"{target_category.lower().replace(' ', '_')}_active.json"
                        save_category_data(target_category, parsed_res["scales"], parsed_res["rect_rates"], new_data_file)
                        registry[target_category]["active"] = {
                            "filename": up_chart.name,
                            "approval_date": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                            "effective_years": parsed_res["effective_years"],
                            "ranks": parsed_res["ranks"],
                            "data_file": new_data_file
                        }
                        save_registry(registry)
                        st.success("🎉 New standard successfully activated!")
                        st.rerun()

    active_scales, active_rect_rates = load_category_data(target_category, cat_meta["data_file"])

    # ==========================================================================
    # MODULE 1: CAF PAY FIXATION (MIXED GD-TECH & NCsE UP TO BPS-11 PRESERVED)
    # ==========================================================================
    if is_caf:
        st.subheader("🎖️ CAF Personnel Service Record")
        c1, c2, c3 = st.columns(3)
        with c1:
            fc_no = st.text_input("FC Number", value="102052")
            name = st.text_input("Full Name", value="Jamal")
        with c2:
            cadre_type = st.radio("Cadre Type", ["Regular Combatant / Technical", "NCs E (Non-Combatant)"], horizontal=True)
            if cadre_type != "NCs E (Non-Combatant)":
                current_rank = st.selectbox("Current / Final Substantive Rank", CAF_ALL_RANKS, index=CAF_ALL_RANKS.index("GD N/Sub"))
            else:
                ncse_init = st.selectbox("Initial Scale at Enrolment", NCSE_SCALES, index=0)
                current_rank = st.selectbox("Current / Substantive Scale", NCSE_SCALES, index=min(3, len(NCSE_SCALES)-1))
        with c3:
            if cadre_type != "NCs E (Non-Combatant)":
                qualification = st.radio("Educational Category at Enrolment", ["Matric", "Non-Matric"], index=0, horizontal=True)
            else:
                qualification = "Matric"
                st.caption("NCs E: Direct BPS enrolment progression.")

        st.markdown("---")
        st.subheader("📅 Service Dates & Progression")
        d_c1, d_c2 = st.columns(2)
        with d_c1:
            doe = st.date_input("Date of Enrolment (D.O.E)",
                                value=datetime.date(2007, 3, 15),
                                min_value=MIN_APP_DATE,
                                max_value=MAX_APP_DATE).strftime("%Y-%m-%d")
            d_remust = None
            if cadre_type != "NCs E (Non-Combatant)":
                d_remust = st.date_input("Date of Remustering (as Trained Soldier / Sepoy)",
                                         value=datetime.date(2007, 9, 10),
                                         min_value=MIN_APP_DATE,
                                         max_value=MAX_APP_DATE).strftime("%Y-%m-%d")

        with d_c2:
            retire_choice = st.checkbox("Has the individual retired?", value=False)
            retirement_date = None
            if retire_choice:
                retirement_date = st.date_input("Date of Retirement",
                                                value=datetime.date(2026, 7, 1),
                                                min_value=MIN_APP_DATE,
                                                max_value=MAX_APP_DATE).strftime("%Y-%m-%d")

        promotions = []
        upgradations = []

        if cadre_type != "NCs E (Non-Combatant)":
            st.markdown("##### 📈 Promotions & Technical Appointments Timeline")
            st.caption("Record each promotion event. You can select either GD or Technical ranks (e.g., GD Sep ➔ GD Lnk ➔ GD Nk ➔ Tech Hav ➔ Tech N/Sub):")
            
            p_num = st.number_input("Number of Promotions / Appointments to Record", min_value=0, max_value=8, value=4, step=1)
            
            default_caf_mix = [
                (datetime.date(2012, 3, 15), "GD Lnk"),
                (datetime.date(2014, 2, 20), "GD Nk"),
                (datetime.date(2017, 12, 2), "GD Hav"),
                (datetime.date(2026, 1, 8), "GD N/Sub")
            ]

            for i in range(p_num):
                col_p1, col_p2 = st.columns([1, 1.2])
                def_d, def_r = default_caf_mix[i] if i < len(default_caf_mix) else (datetime.date(2020, 1, 1), "GD Lnk")
                with col_p1:
                    p_dt = st.date_input(f"Promotion / Appointment {i+1} Date", value=def_d,
                                         min_value=MIN_APP_DATE, max_value=MAX_APP_DATE,
                                         key=f"caf_mix_d_{i}").strftime("%Y-%m-%d")
                with col_p2:
                    cur_r_idx = CAF_ALL_RANKS.index(def_r) if def_r in CAF_ALL_RANKS else 0
                    p_rk = st.selectbox(f"Rank {i+1} (GD or Technical)", CAF_ALL_RANKS, index=cur_r_idx, key=f"caf_mix_r_{i}")
                
                promotions.append({"date": p_dt, "rank": p_rk})

        else:
            st.markdown("##### 📈 NCs E Movement to Higher Scale (BPS-01 to BPS-11)")
            st.caption("Record dates of transition to subsequent higher pay scales:")
            num_nc_moves = st.number_input("Number of Higher Scale Movements to Record", min_value=0, max_value=10, value=2, step=1)
            
            default_nc_moves = [
                (datetime.date(2012, 12, 1), "BPS-02"),
                (datetime.date(2017, 12, 1), "BPS-03"),
                (datetime.date(2021, 12, 1), "BPS-04"),
                (datetime.date(2024, 12, 1), "BPS-05")
            ]
            for m_i in range(num_nc_moves):
                m_c1, m_c2 = st.columns(2)
                def_d, def_s = default_nc_moves[m_i] if m_i < len(default_nc_moves) else (datetime.date(2020, 1, 1), NCSE_SCALES[min(m_i+1, len(NCSE_SCALES)-1)])
                with m_c1:
                    m_d = st.date_input(f"Movement {m_i+1} Date", value=def_d,
                                        min_value=MIN_APP_DATE, max_value=MAX_APP_DATE, key=f"nc_d_{m_i}").strftime("%Y-%m-%d")
                with m_c2:
                    s_idx = NCSE_SCALES.index(def_s) if def_s in NCSE_SCALES else min(m_i+1, len(NCSE_SCALES)-1)
                    m_s = st.selectbox(f"Scale {m_i+1}", NCSE_SCALES, index=s_idx, key=f"nc_s_{m_i}")
                upgradations.append({"date": m_d, "rank": normalize_rank_str(m_s), "remarks": f"Moved to {m_s}"})

        st.markdown("---")
        user_notes = st.text_input(
            "Special Instructions / Directives (Optional):",
            placeholder="Type notes or official order reference here...",
            key="caf_notes"
        )

        st.markdown("---")
        if st.button("🚀 Calculate CAF Pay Fixation Statement", type="primary"):
            if cadre_type == "NCs E (Non-Combatant)":
                mapped_enrol_rank = "NCs E"
                ncse_initial_code = normalize_rank_str(ncse_init)
            else:
                first_rank = promotions[0]["rank"] if promotions else current_rank
                mapped_enrol_rank = "Tech Sep" if "Tech" in first_rank else "GD Sep"
                ncse_initial_code = None

            personnel_payload = {
                "name": name,
                "fc_no": fc_no,
                "rank": current_rank,
                "enrol_rank": mapped_enrol_rank,
                "enrol_post": "GD Rect",
                "qualification": qualification,
                "doe": doe,
                "d_remust": d_remust,
                "remust_rank": mapped_enrol_rank,
                "ncse_init_scale": ncse_initial_code,
                "transfers": [],
                "promotions": promotions,
                "upgradations": upgradations,
                "retirement_date": retirement_date,
                "end_date": "2030-12-31",
                "category": "CAF",
                "chart_name": cat_meta["filename"]
            }

            result = run_comprehensive_pay_fixation(personnel_payload, active_scales, active_rect_rates, user_notes)

            if result.get("warnings"):
                for w in result["warnings"]:
                    st.warning(f"⚠️ {w}")

            st.success("✅ CAF Pay fixation calculated successfully in official Ilyas Sb format!")

            official_title = get_official_heading(personnel_payload)
            st.markdown(f"#### {official_title}")

            particulars_data = build_particulars_rows(personnel_payload)
            df_part = pd.DataFrame(particulars_data, columns=["Marker", "Milestone", "Date"])
            st.table(df_part)

            if user_notes.strip():
                st.markdown(f"**Note:** *{user_notes}*")

            df_ledger = pd.DataFrame(result["ledger"])
            df_ledger.index = df_ledger.index + 1
            st.dataframe(df_ledger, use_container_width=True)

            # Downloads
            st.markdown("### 📥 Download Formal Statements (Ilyas Sb Format)")
            dcol1, dcol2, dcol3 = st.columns(3)
            dcol1.download_button(label="📊 Download Excel (.xlsx)", data=generate_excel_bytes(result), file_name=f"Pay_Fixation_{fc_no}_{name}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            dcol2.download_button(label="📄 Download Word (.docx)", data=generate_word_bytes(result), file_name=f"Pay_Fixation_{fc_no}_{name}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            dcol3.download_button(label="📑 Download PDF (.pdf)", data=generate_pdf_bytes(result), file_name=f"Pay_Fixation_{fc_no}_{name}.pdf", mime="application/pdf")

    # ==========================================================================
    # MODULE 2: COMPLETE CIVIL SERVANTS WORKFLOW (PRESERVED)
    # ==========================================================================
    else:
        st.subheader("👤 Federal Civil Servant Dossier")
        c1, c2, c3 = st.columns(3)
        with c1:
            fc_no = st.text_input("Personnel / Service ID / CNIC", value="CS-10294")
            name = st.text_input("Officer / Employee Full Name", value="Muhammad Iqbal")
        with c2:
            enrol_post = st.text_input("Designation at Appointment", value="GD Recruit Metric")
            enrol_rank = st.selectbox("Scale at Appointment", ALL_BPS_SCALES, index=3)
        with c3:
            final_rank = st.selectbox("Current / Substantive Scale", ALL_BPS_SCALES, index=12)
            qualification = "Matric"

        st.markdown("---")
        st.subheader("📅 Service Chronology & Key Milestones (1990 to 2030)")
        d_col1, d_col2 = st.columns(2)
        with d_col1:
            doe = st.date_input("Date of Appointment / Enrolment",
                                value=datetime.date(2006, 9, 14),
                                min_value=MIN_APP_DATE, max_value=MAX_APP_DATE).strftime("%Y-%m-%d")

            remust_choice = st.checkbox("Was the employee remustered / trained?", value=True)
            d_remust = None
            remust_post = "GD Sepoy"
            remust_rank = "BPS-05"
            if remust_choice:
                r_c1, r_c2, r_c3 = st.columns(3)
                with r_c1:
                    d_remust = st.date_input("Date of Remustering", value=datetime.date(2007, 4, 20),
                                             min_value=MIN_APP_DATE, max_value=MAX_APP_DATE).strftime("%Y-%m-%d")
                with r_c2:
                    remust_post = st.text_input("Remustered Post", value="GD Sepoy")
                with r_c3:
                    remust_rank = st.selectbox("Remustered Scale", ALL_BPS_SCALES, index=4)

        with d_col2:
            retire_choice = st.checkbox("Has the employee retired?", value=False)
            retirement_date = None
            if retire_choice:
                retirement_date = st.date_input("Date of Superannuation / Retirement",
                                                value=datetime.date(2024, 8, 14),
                                                min_value=MIN_APP_DATE, max_value=MAX_APP_DATE).strftime("%Y-%m-%d")

        # Lateral Transfers
        st.markdown("---")
        st.subheader("🔄 Lateral Transfers & Cadre Shifts (FR 22-B)")
        trans_choice = st.checkbox("Did the employee undergo a lateral transfer (e.g. Soldier Clerk)?", value=True)
        transfers = []
        if trans_choice:
            t_c1, t_c2, t_c3 = st.columns(3)
            with t_c1:
                t_dt = st.date_input("Transfer Date", value=datetime.date(2009, 2, 1),
                                     min_value=MIN_APP_DATE, max_value=MAX_APP_DATE, key="t_dt_civ").strftime("%Y-%m-%d")
            with t_c2:
                t_post = st.text_input("Transferred Post", value="Soldier Clerk", key="t_p_civ")
            with t_c3:
                t_rk = st.selectbox("Transferred Scale", ALL_BPS_SCALES, index=4, key="t_r_civ")
            transfers.append({"date": t_dt, "post_designation": t_post, "rank": t_rk})

        # Promotions Timeline with Automatic Cadre Upgradation Intelligence
        st.markdown("---")
        st.subheader("📈 Promotions Timeline")
        st.caption("ℹ️ **Automatic Post Upgradation Active:** Enter promotion dates and posts. The application automatically determines the applicable BPS for that year and injects all subsequent official cadre up-gradations (e.g. Junior Clerk to BPS-09 in 2016, Senior Clerk to BPS-13 in 2023) with OM 31-05-2013 premature increments.")

        p_num = st.number_input("Number of Promotions to Record", min_value=0, max_value=8, value=2, step=1)
        promotions = []
        default_promos_data = [
            (datetime.date(2010, 2, 19), "Junior Clerk", "BPS-07"),
            (datetime.date(2017, 3, 1), "Senior Clerk", "BPS-11"),
            (datetime.date(2020, 1, 1), "Assistant", "BPS-16"),
            (datetime.date(2024, 1, 1), "Superintendent", "BPS-17")
        ]

        for i in range(p_num):
            p_c1, p_c2, p_c3 = st.columns([1, 1.3, 1])
            def_dt, def_name, def_sc = default_promos_data[i] if i < len(default_promos_data) else (datetime.date(2020, 1, 1), "Junior Clerk", "BPS-07")
            with p_c1:
                p_date = st.date_input(f"Promotion {i+1} Date", value=def_dt,
                                       min_value=MIN_APP_DATE, max_value=MAX_APP_DATE,
                                       key=f"civ_promo_d_{i}").strftime("%Y-%m-%d")
            with p_c2:
                p_post = st.text_input(f"Post {i+1} Designation", value=def_name, key=f"civ_promo_p_{i}")
            with p_c3:
                auto_sc = get_automatic_scale_for_civil_post(p_post, p_date)
                use_sc = auto_sc if auto_sc else def_sc
                sc_idx = ALL_BPS_SCALES.index(use_sc) if use_sc in ALL_BPS_SCALES else 6
                p_scale = st.selectbox(f"Scale {i+1} (Auto / Select)", ALL_BPS_SCALES, index=sc_idx, key=f"civ_promo_r_{i}")
            promotions.append({"date": p_date, "post_designation": p_post, "rank": p_scale})

        # Manual Custom Upgradations (for special non-cadre posts)
        st.markdown("---")
        st.subheader("⭐ Custom Departmental Up-gradations (Optional Overrides)")
        manual_upgradations = []
        with st.expander("➕ Add Non-Standard / Departmental Up-gradation (If Not Covered by Federal OMs)"):
            u_chk = st.checkbox("Record Custom Departmental Up-gradation", value=False)
            if u_chk:
                u_c1, u_c2, u_c3 = st.columns(3)
                with u_c1:
                    u_d = st.date_input("Upgradation Date", value=datetime.date(2016, 7, 1),
                                         min_value=MIN_APP_DATE, max_value=MAX_APP_DATE, key="civ_u_d").strftime("%Y-%m-%d")
                with u_c2:
                    u_p = st.text_input("Upgraded Post", value="Junior Clerk", key="civ_u_p")
                with u_c3:
                    u_s = st.selectbox("Upgraded Scale", ALL_BPS_SCALES, index=8, key="civ_u_s")
                manual_upgradations.append({"date": u_d, "designation": u_p, "rank": u_s})

        st.markdown("---")
        user_notes = st.text_input(
            "Directives & Notification References:",
            value="Automatic Cadre Upgradation applied under Finance Division OMs (2007, 2016, 2023) read with OM No. F.11(4)R-2/2011-1153/2013 dated 31-05-2013 (Premature Increment on Upgradation).",
            key="civ_notes"
        )

        st.markdown("---")
        if st.button("🚀 Calculate Civil Servants Pay Fixation Statement", type="primary"):
            personnel_payload = {
                "name": name,
                "fc_no": fc_no,
                "rank": final_rank,
                "enrol_rank": enrol_rank,
                "enrol_post": enrol_post,
                "qualification": qualification,
                "doe": doe,
                "d_remust": d_remust,
                "remust_post": remust_post,
                "remust_rank": remust_rank,
                "transfers": transfers,
                "promotions": promotions,
                "upgradations": manual_upgradations,
                "retirement_date": retirement_date,
                "end_date": "2030-12-31",
                "category": "Civil Servants",
                "chart_name": cat_meta["filename"]
            }

            result = run_comprehensive_pay_fixation(personnel_payload, active_scales, active_rect_rates, user_notes)
            personnel_payload["upgradations"] = result.get("discovered_upgradations", [])

            if result.get("warnings"):
                for w in result["warnings"]:
                    st.warning(f"⚠️ {w}")

            st.success("✅ Civil Servants Pay fixation calculated with automatic post up-gradations and premature increments!")

            official_title = get_official_heading(personnel_payload)
            st.markdown(f"#### {official_title}")

            particulars_data = build_particulars_rows(personnel_payload)
            df_part = pd.DataFrame(particulars_data, columns=["Marker", "Milestone", "Date"])
            st.table(df_part)

            if user_notes.strip():
                st.markdown(f"**Note:** *{user_notes}*")

            df_ledger = pd.DataFrame(result["ledger"])
            df_ledger.index = df_ledger.index + 1
            st.dataframe(df_ledger, use_container_width=True)

            # Downloads
            st.markdown("### 📥 Download Formal Statements (Ilyas Sb Layout)")
            dcol1, dcol2, dcol3 = st.columns(3)
            dcol1.download_button(label="📊 Download Excel (.xlsx)", data=generate_excel_bytes(result), file_name=f"Pay_Fixation_{fc_no}_{name}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            dcol2.download_button(label="📄 Download Word (.docx)", data=generate_word_bytes(result), file_name=f"Pay_Fixation_{fc_no}_{name}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            dcol3.download_button(label="📑 Download PDF (.pdf)", data=generate_pdf_bytes(result), file_name=f"Pay_Fixation_{fc_no}_{name}.pdf", mime="application/pdf")

    # App Sharing Footer
    st.markdown("---")
    sh_col1, sh_col2, sh_col3 = st.columns([1, 3, 2])
    with sh_col1:
        if logo_file:
            try:
                st.image(logo_file, width=80)
            except Exception:
                pass
    with sh_col2:
        st.markdown("**Pay Fixation Agent — Direct Access & Share**")
        app_url = st.text_input("App / Repository Direct Link:", value="https://github.com", key="share_url")
        st.markdown(f"🔗 **Direct Link:** [{app_url}]({app_url})")
    with sh_col3:
        qr_url = f"https://api.qrserver.com/v1/create-qr-code/?size=120x120&data={app_url}"
        st.image(qr_url, caption="Scan QR to open app", width=120)

# ==============================================================================
# 9. FAIL-SAFE APPLICATION ENTRYPOINT
# ==============================================================================
if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        import traceback
        st.error(f"Application Runtime Notice: {str(e)}")
        st.code(traceback.format_exc())
