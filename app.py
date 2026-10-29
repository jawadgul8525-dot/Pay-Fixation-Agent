import streamlit as st
import datetime
import io
import os
import re
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
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

# ==============================================================================
# 1. VERIFIED PAY SCALES (1994 – 2026) INCLUDING SUB, SM, & NCs E
# ==============================================================================
DEFAULT_SCALES = {
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

RECT_RATES = {
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

REVISIONS_SCHEDULE = [
    ("2005-07-01_REV", "2005-07-01", "Revision"),
    ("2005-07-01_UPG", "2005-07-01", "Upgradation"),
    ("2007-07-01", "2007-07-01", "Revision"),
    ("2008-07-01", "2008-07-01", "Revision"),
    ("2011-07-01", "2011-07-01", "Revision"),
    ("2015-07-01", "2015-07-01", "Revision"),
    ("2016-07-01", "2016-07-01", "Revision"),
    ("2017-07-01", "2017-07-01", "Revision"),
    ("2022-07-01", "2022-07-01", "Revision"),
    ("2024-06-01", "2024-06-01", "Revision"),
    ("2026-07-01", "2026-07-01", "Revision")
]

# ==============================================================================
# 2. HELPER FUNCTIONS & UPLOAD PARSER
# ==============================================================================
def format_scale_str(sc):
    return f"{sc[0]}-{sc[1]}-{sc[2]}" if sc else ""

def get_rect_pay(qualification, enrol_date_str):
    rates = RECT_RATES.get(qualification, RECT_RATES["Matric"])
    applicable_pay = None
    for rev_dt, rate in rates:
        if rev_dt <= enrol_date_str:
            applicable_pay = rate
    if not applicable_pay:
        applicable_pay = rates[0][1]
    return applicable_pay

def get_scale_at_date(rank, target_date_str, scales_dict, rev_tag=None):
    clean_rank = rank.replace("GD ", "")
    lookup_rank = rank if rank in scales_dict else clean_rank
    if lookup_rank not in scales_dict:
        for k in scales_dict:
            if k in rank or rank in k:
                lookup_rank = k
                break
    if lookup_rank not in scales_dict:
        return None
    r_scales = scales_dict[lookup_rank]
    if rev_tag and rev_tag in r_scales:
        return r_scales[rev_tag]
    target_dt = datetime.datetime.strptime(target_date_str, "%Y-%m-%d").date()
    applicable = None
    for k, sc in r_scales.items():
        dt_part = k.split("_")[0]
        if datetime.datetime.strptime(dt_part, "%Y-%m-%d").date() <= target_dt:
            applicable = sc
    return applicable

def parse_uploaded_pay_chart(uploaded_file):
    """Safely parse uploaded Excel/CSV pay chart into runtime dictionary."""
    try:
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)
        
        # Look for rank column and scale string cells
        rank_col = None
        for col in df.columns:
            if "rank" in str(col).lower():
                rank_col = col
                break
        if rank_col is None:
            rank_col = df.columns[0]
            
        custom_scales = {}
        date_pattern = re.compile(r'(\d{1,2}[-/]\d{1,2}[-/]\d{2,4})')
        scale_pattern = re.compile(r'(\d+)\s*[-/]\s*(\d+)\s*[-/]\s*(\d+)')

        for _, row in df.iterrows():
            raw_rank = str(row[rank_col]).strip()
            if not raw_rank or raw_rank.lower() == 'nan':
                continue
            custom_scales[raw_rank] = {}
            for col_name in df.columns:
                val = str(row[col_name]).strip()
                m_scale = scale_pattern.search(val)
                if m_scale:
                    min_p, inc_p, max_p = map(int, m_scale.groups())
                    # Check column name or text for date
                    m_date = date_pattern.search(str(col_name))
                    if m_date:
                        raw_d = m_date.group(1).replace('/', '-')
                        parts = raw_d.split('-')
                        if len(parts[0]) == 4:
                            iso_d = f"{parts[0]}-{int(parts[1]):02d}-{int(parts[2]):02d}"
                        else:
                            iso_d = f"{int(parts[2]):04d}-{int(parts[1]):02d}-{int(parts[0]):02d}"
                        custom_scales[raw_rank][iso_d] = (min_p, inc_p, max_p)
                        
        if not custom_scales:
            return None, "No recognizable pay scales (Min-Inc-Max) found in the file."
        return custom_scales, None
    except Exception as e:
        return None, f"Failed to parse file: {str(e)}"

# ==============================================================================
# 3. COMPLETE CALCULATION ENGINE
# ==============================================================================
def run_pay_fixation_engine(personnel, scales_dict, special_instructions=""):
    doe = personnel["doe"]
    enrol_rank = personnel.get("enrol_rank", "GD Sep")
    qualification = personnel.get("qualification", "Matric")
    d_remust = personnel.get("d_remust")
    promotions = personnel.get("promotions", [])
    end_date_str = personnel.get("end_date", "2026-07-01")
    warnings = []

    events = []
    is_ncs_e = (enrol_rank == "NCs E")

    if is_ncs_e:
        events.append({
            "date": doe, "type": "ENROLLMENT", "rank": "NCs E",
            "remarks": "Enrolled directly as NCs E"
        })
    else:
        events.append({
            "date": doe, "type": "ENROLLMENT", "rank": "GD Rect",
            "remarks": f"Enrolled as GD Rect ({qualification})"
        })
        if d_remust:
            events.append({
                "date": d_remust, "type": "REMUSTERING", "rank": "GD Sep",
                "remarks": "Remustered as train sldr"
            })

    for p in promotions:
        p_rank = p["rank"]
        rem = f"Pmtd {p_rank.replace('GD ', '')}"
        events.append({"date": p["date"], "type": "PROMOTION", "rank": p_rank, "remarks": rem})

    for r_key, r_dt, r_type in REVISIONS_SCHEDULE:
        if doe <= r_dt <= end_date_str:
            events.append({"date": r_dt, "type": r_type.upper(), "rev_key": r_key, "remarks": r_type})

    start_year = datetime.datetime.strptime(doe, "%Y-%m-%d").year
    end_year = datetime.datetime.strptime(end_date_str, "%Y-%m-%d").year
    for y in range(start_year, end_year + 1):
        inc_dt = f"{y}-12-01"
        if doe <= inc_dt <= end_date_str:
            events.append({"date": inc_dt, "type": "ANNUAL_INCREMENT", "remarks": "A/Inc"})

    priority = {
        "ENROLLMENT": 0, "REMUSTERING": 1, "REVISION": 2,
        "UPGRADATION": 3, "PROMOTION": 4, "ANNUAL_INCREMENT": 5
    }
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
                    warnings.append(f"NCs E scale not found for date {ev['date']}. Review required.")
                    current_pay = 0
                else:
                    current_pay = current_scale[0]
                last_change_dt = ev_dt
                ledger.append({
                    "date": ev_dt.strftime("%d-%m-%Y"), "pay": current_pay,
                    "scale": format_scale_str(current_scale), "remarks": ev["remarks"]
                })
            else:
                current_pay = get_rect_pay(qualification, ev["date"])
                current_scale = get_scale_at_date("GD Sep", ev["date"], scales_dict)
                substantive_scale = current_scale
                last_change_dt = ev_dt
                ledger.append({
                    "date": ev_dt.strftime("%d-%m-%Y"), "pay": current_pay,
                    "scale": "Fixed", "remarks": ev["remarks"]
                })

        elif ev_t == "REMUSTERING":
            current_rank = "GD Sep"
            current_scale = get_scale_at_date("GD Sep", ev["date"], scales_dict)
            substantive_scale = current_scale
            current_pay = current_scale[0]
            last_change_dt = ev_dt
            ledger.append({
                "date": ev_dt.strftime("%d-%m-%Y"), "pay": current_pay,
                "scale": format_scale_str(current_scale), "remarks": ev["remarks"]
            })

        elif ev_t == "ANNUAL_INCREMENT":
            days_served = (ev_dt - last_change_dt).days if last_change_dt else 365
            if days_served < 180:
                ledger.append({
                    "date": ev_dt.strftime("%d-%m-%Y"), "pay": current_pay,
                    "scale": '"', "remarks": "Not entitled (< 6 mos in scale)"
                })
            else:
                active_inc = substantive_scale[1] if substantive_scale else current_scale[1]
                current_pay += active_inc
                ledger.append({
                    "date": ev_dt.strftime("%d-%m-%Y"), "pay": current_pay,
                    "scale": '"', "remarks": "A/Inc"
                })

        elif ev_t == "REVISION":
            active_rank = current_rank
            if active_rank in ["H/Lt", "H/Capt"]:
                active_rank = "Sub" if "Sub" in str(substantive_scale) else "SM"
            new_sc = get_scale_at_date(active_rank, ev["date"], scales_dict, rev_tag=ev.get("rev_key"))
            if new_sc and current_scale:
                stages = (current_pay - current_scale[0]) // current_scale[1]
                current_pay = new_sc[0] + (stages * new_sc[1])
                current_scale = new_sc
                substantive_scale = new_sc
                ledger.append({
                    "date": ev_dt.strftime("%d-%m-%Y"), "pay": current_pay,
                    "scale": format_scale_str(current_scale), "remarks": "Revision"
                })

        elif ev_t == "UPGRADATION":
            active_rank = current_rank
            if active_rank in ["H/Lt", "H/Capt"]:
                active_rank = "SM"
            new_sc = get_scale_at_date(active_rank, ev["date"], scales_dict, rev_tag=ev.get("rev_key"))
            if new_sc:
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
                # Honorary promotions preserve substantive scale and add 1 SM annual increment
                sm_scale = get_scale_at_date("SM", ev["date"], scales_dict)
                if not sm_scale:
                    warnings.append(f"SM Scale unavailable on {ev['date']} for {new_rank} increment. Flagged for manual review.")
                    honorary_inc = 0
                else:
                    honorary_inc = sm_scale[1]
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
                    warnings.append(f"Scale missing for rank {current_rank} on {ev['date']}. Review required.")
                else:
                    p_step = new_sc[0]
                    while p_step < current_pay:
                        p_step += new_sc[1]
                    current_pay = p_step + new_sc[1]  # Equal/Over + 1 promotional increment
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
# 4. REPORT GENERATORS (EXCEL, WORD, PDF)
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
        ("a", "D.O E", h.get("doe")),
        ("b", "D.O.Remustration", h.get("d_remust", "N/A (Direct NCs E)"))
    ]
    letters = ["c", "d", "e", "f", "g", "h"]
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
        ws.cell(row=r, column=2, value=f"User Special Instructions: {data['special_instructions']} (Verification Required)").font = Font(italic=True, color="000080")

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
        r[1].width = Inches(2.5)
        r[2].width = Inches(2.0)

    add_bio("a.", "D.O E", h.get("doe"))
    add_bio("b.", "D.O.Remustration", h.get("d_remust", "N/A"))
    letters = ["c.", "d.", "e.", "f.", "g.", "h."]
    for idx, p in enumerate(h.get("promotions", [])):
        add_bio(letters[idx], f"D.O.Promotion as {p['rank'].replace('GD ', '')}", p["date"])

    if data.get("special_instructions"):
        p_note = doc.add_paragraph()
        r_note = p_note.add_run(f"\nSpecial Instructions (Audit Note): {data['special_instructions']} (Must be verified against official orders)")
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
        elements.append(Spacer(1, 5))

    h = data["header"]
    title_text = f"<b>PAY FIXATION IRO FC NO. {h.get('fc_no','')} {h.get('rank','')} {h.get('name','').upper()}</b>"
    elements.append(Paragraph(title_text, styles['Title']))
    elements.append(Spacer(1, 10))

    bio_data = [
        ["a.", "D.O E", str(h.get("doe"))],
        ["b.", "D.O.Remustration", str(h.get("d_remust", "N/A"))]
    ]
    letters = ["c.", "d.", "e.", "f.", "g.", "h."]
    for idx, p in enumerate(h.get("promotions", [])):
        bio_data.append([letters[idx], f"D.O.Promotion as {p['rank'].replace('GD ', '')}", str(p["date"])])

    t_bio = Table(bio_data, colWidths=[30, 200, 120])
    t_bio.setStyle(TableStyle([
        ('FONTNAME', (0,0), (-1,-1), 'Helvetica'),
        ('FONTSIZE', (0,0), (-1,-1), 9),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
    ]))
    elements.append(t_bio)

    if data.get("special_instructions"):
        elements.append(Spacer(1, 8))
        elements.append(Paragraph(f"<i>Special Instructions: {data['special_instructions']} (Verification Required)</i>", styles['Normal']))

    elements.append(Spacer(1, 12))

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
    st.set_page_config(
    page_title="Pay Fixation Agent",
    layout="wide",
    page_icon="logo.png"
)
col_logo, col_title = st.columns([1, 6])

with col_logo:
    st.image("logo.png", width=95)

with col_title:
    st.title("Pay Fixation Agent")
       
        st.caption("Developed by **SC Rehman Gul** to assist Civil Armed Forces (CAF) personnel in preparing and auditing basic pay fixation.")

    st.info("⚠️ **Official Audit Disclaimer:** All calculated results must be verified against official departmental notifications, government gazettes, and original service books.")

    # Sidebar: Rules & Dynamic Upload Manager
    st.sidebar.header("📁 Pay Chart Management")
    uploaded_chart = st.sidebar.file_uploader("Upload New Pay Chart (Excel / CSV)", type=["xlsx", "csv"])
    runtime_scales = DEFAULT_SCALES.copy()

    if uploaded_chart is not None:
        parsed_scales, err = parse_uploaded_pay_chart(uploaded_chart)
        if err:
            st.sidebar.error(f"❌ {err}")
            st.sidebar.warning("Falling back to verified default pay charts.")
        else:
            runtime_scales.update(parsed_scales)
            st.sidebar.success(f"✅ Loaded {len(parsed_scales)} rank scales from file!")

    # Input Form
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("👤 Personnel Details")
        fc_no = st.text_input("FC Number", value="102052")
        name = st.text_input("Full Name", value="Jamal")
        rank_list = ["NCs E", "GD Sep", "GD Lnk", "GD Nk", "GD Hav", "GD N/Sub", "Sub", "SM", "H/Lt", "H/Capt"]
        final_rank = st.selectbox("Current / Final Rank", rank_list, index=5)
        enrol_rank = st.selectbox("Enrolled As", ["GD Rect", "NCs E"], index=0)
        qualification = "Matric"
        if enrol_rank == "GD Rect":
            qualification = st.radio("Recruit Educational Category", ["Matric", "Non-Matric"], horizontal=True)

        doe = st.date_input("Date of Enrollment (D.O.E)",
                            value=datetime.date(2007, 3, 15),
                            min_value=datetime.date(1990, 1, 1),
                            max_value=datetime.date(2035, 12, 31)).strftime("%Y-%m-%d")

        d_remust = None
        if enrol_rank != "NCs E":
            d_remust = st.date_input("Date of Remustering",
                                     value=datetime.date(2007, 9, 10),
                                     min_value=datetime.date(1990, 1, 1),
                                     max_value=datetime.date(2035, 12, 31)).strftime("%Y-%m-%d")

    with col2:
        st.subheader("📈 Promotion Timeline")
        st.caption("Add promotion events across your entire service history (1990 to date):")
        num_promotions = st.number_input("Number of Promotions to Add", min_value=0, max_value=8, value=4, step=1)
        promotions = []
        default_promos = [
            (datetime.date(2012, 3, 15), "GD Lnk"),
            (datetime.date(2014, 2, 20), "GD Nk"),
            (datetime.date(2017, 12, 2), "GD Hav"),
            (datetime.date(2026, 1, 8), "GD N/Sub"),
            (datetime.date(2028, 1, 1), "Sub"),
            (datetime.date(2030, 1, 1), "SM"),
            (datetime.date(2031, 1, 1), "H/Lt"),
            (datetime.date(2032, 1, 1), "H/Capt")
        ]
        for i in range(num_promotions):
            p_col1, p_col2 = st.columns([1, 1])
            def_dt, def_rk = default_promos[i] if i < len(default_promos) else (datetime.date(2020, 1, 1), "Sub")
            with p_col1:
                p_dt = st.date_input(f"Promo {i+1} Date", value=def_dt,
                                     min_value=datetime.date(1990, 1, 1),
                                     max_value=datetime.date(2035, 12, 31),
                                     key=f"p_dt_{i}").strftime("%Y-%m-%d")
            with p_col2:
                p_rk = st.selectbox(f"Promo {i+1} Rank", rank_list[1:],
                                    index=rank_list[1:].index(def_rk) if def_rk in rank_list[1:] else 0,
                                    key=f"p_rk_{i}")
            promotions.append({"date": p_dt, "rank": p_rk})

    st.markdown("---")
    st.subheader("📝 Special Instructions & Directives")
    special_instructions = st.text_input(
        "Enter special instructions (e.g., 'Adv increment for qualification', 'Withheld increment in 2018', 'Protected pay order'):",
        placeholder="Type note or order reference here..."
    )
    if special_instructions:
        st.warning("⚠️ **Notice:** User instructions are recorded and highlighted in all generated reports. Non-standard calculations must be cross-checked manually against verified official orders.")

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
            "end_date": "2026-07-01"
        }

        result = run_pay_fixation_engine(personnel_payload, runtime_scales, special_instructions)

        if result.get("warnings"):
            for w in result["warnings"]:
                st.warning(f"⚠️ {w}")

        st.success("✅ Pay fixation calculated and audited successfully!")
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

if __name__ == "__main__":
    main()
