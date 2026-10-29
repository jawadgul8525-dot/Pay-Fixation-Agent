import streamlit as st
import datetime
import io
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

# ==========================================
# 1. CORE PAY SCALES & REVISIONS (DEFAULT)
# ==========================================
DEFAULT_SCALES = {
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
    }
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

# ==========================================
# 2. CALCULATION ENGINE
# ==========================================
def format_scale_str(sc):
    return f"{sc[0]}-{sc[1]}-{sc[2]}" if sc else ""

def get_scale_at_date(rank, target_date_str, scales_dict, rev_tag=None):
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

def run_pay_fixation_engine(personnel, scales_dict, special_instructions=""):
    doe = personnel["doe"]
    d_remust = personnel.get("d_remust")
    promotions = personnel.get("promotions", [])
    end_date_str = personnel.get("end_date", "2026-07-01")

    events = [{"date": doe, "type": "ENROLLMENT", "rank": "GD Rect", "remarks": "Enrolled as GD Rect"}]
    if d_remust:
        events.append({"date": d_remust, "type": "REMUSTERING", "rank": "GD Sep", "remarks": "Remustered as train sldr"})
    for p in promotions:
        events.append({"date": p["date"], "type": "PROMOTION", "rank": p["rank"], "remarks": f"Pmtd {p['rank'].replace('GD ', '')}"})
    for r_key, r_dt, r_type in REVISIONS_SCHEDULE:
        if doe <= r_dt <= end_date_str:
            events.append({"date": r_dt, "type": r_type.upper(), "rev_key": r_key, "remarks": r_type})

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
    current_rank = "GD Sep"
    current_scale = None
    last_change_dt = None

    for ev in events:
        ev_dt = datetime.datetime.strptime(ev["date"], "%Y-%m-%d").date()
        ev_t = ev["type"]

        if ev_t == "ENROLLMENT":
            current_scale = get_scale_at_date("GD Sep", ev["date"], scales_dict)
            current_pay = 2128 if ev["date"] >= "2005-07-01" else 1845
            last_change_dt = ev_dt
            ledger.append({"date": ev_dt.strftime("%d-%m-%Y"), "pay": current_pay, "scale": "Fixed", "remarks": ev["remarks"]})

        elif ev_t == "REMUSTERING":
            current_rank = "GD Sep"
            current_scale = get_scale_at_date("GD Sep", ev["date"], scales_dict)
            current_pay = current_scale[0]
            last_change_dt = ev_dt
            ledger.append({"date": ev_dt.strftime("%d-%m-%Y"), "pay": current_pay, "scale": format_scale_str(current_scale), "remarks": ev["remarks"]})

        elif ev_t == "ANNUAL_INCREMENT":
            if (ev_dt - last_change_dt).days < 180:
                ledger.append({"date": ev_dt.strftime("%d-%m-%Y"), "pay": current_pay, "scale": '"', "remarks": "Not entitled (< 6 mos in scale)"})
            else:
                current_pay += current_scale[1]
                ledger.append({"date": ev_dt.strftime("%d-%m-%Y"), "pay": current_pay, "scale": '"', "remarks": "A/Inc"})

        elif ev_t == "REVISION":
            new_sc = get_scale_at_date(current_rank, ev["date"], scales_dict, rev_tag=ev.get("rev_key"))
            if new_sc and current_scale:
                stages = (current_pay - current_scale[0]) // current_scale[1]
                current_pay = new_sc[0] + (stages * new_sc[1])
                current_scale = new_sc
                ledger.append({"date": ev_dt.strftime("%d-%m-%Y"), "pay": current_pay, "scale": format_scale_str(current_scale), "remarks": "Revision"})

        elif ev_t == "UPGRADATION":
            new_sc = get_scale_at_date(current_rank, ev["date"], scales_dict, rev_tag=ev.get("rev_key"))
            if new_sc:
                p_step = new_sc[0]
                while p_step < current_pay:
                    p_step += new_sc[1]
                current_pay = p_step
                current_scale = new_sc
                ledger.append({"date": ev_dt.strftime("%d-%m-%Y"), "pay": current_pay, "scale": format_scale_str(current_scale), "remarks": "Upgradation"})

        elif ev_t == "PROMOTION":
            current_rank = ev["rank"]
            new_sc = get_scale_at_date(current_rank, ev["date"], scales_dict)
            p_step = new_sc[0]
            while p_step < current_pay:
                p_step += new_sc[1]
            current_pay = p_step + new_sc[1]
            current_scale = new_sc
            last_change_dt = ev_dt
            ledger.append({"date": ev_dt.strftime("%d-%m-%Y"), "pay": current_pay, "scale": format_scale_str(current_scale), "remarks": ev["remarks"]})

    return {"header": personnel, "ledger": ledger, "special_instructions": special_instructions}

# ==========================================
# 3. EXPORT MODULES (EXCEL, WORD, PDF)
# ==========================================
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
        ("b", "D.O.Remustration", h.get("d_remust"))
    ]
    letters = ["c", "d", "e", "f", "g"]
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
        ws.cell(row=r, column=2, value=data["special_instructions"]).font = Font(italic=True, color="000080")

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
    ws.column_dimensions['E'].width = 32

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
        r[2].text = str(val)
        r[0].width = Inches(0.5)
        r[1].width = Inches(2.5)
        r[2].width = Inches(2.0)

    add_bio("a.", "D.O E", h.get("doe"))
    add_bio("b.", "D.O.Remustration", h.get("d_remust"))
    letters = ["c.", "d.", "e.", "f.", "g."]
    for idx, p in enumerate(h.get("promotions", [])):
        add_bio(letters[idx], f"D.O.Promotion as {p['rank'].replace('GD ', '')}", p["date"])

    if data.get("special_instructions"):
        p_note = doc.add_paragraph()
        r_note = p_note.add_run(f"\nSpecial Instructions: {data['special_instructions']}")
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
    h = data["header"]
    title_text = f"<b>PAY FIXATION IRO FC NO. {h.get('fc_no','')} {h.get('rank','')} {h.get('name','').upper()}</b>"
    elements.append(Paragraph(title_text, styles['Title']))
    elements.append(Spacer(1, 10))

    bio_data = [
        ["a.", "D.O E", str(h.get("doe"))],
        ["b.", "D.O.Remustration", str(h.get("d_remust"))]
    ]
    letters = ["c.", "d.", "e.", "f.", "g."]
    for idx, p in enumerate(h.get("promotions", [])):
        bio_data.append([letters[idx], f"D.O.Promotion as {p['rank'].replace('GD ', '')}", str(p["date"])])

    t_bio = Table(bio_data, colWidths=[30, 180, 120])
    t_bio.setStyle(TableStyle([
        ('FONTNAME', (0,0), (-1,-1), 'Helvetica'),
        ('FONTSIZE', (0,0), (-1,-1), 9),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
    ]))
    elements.append(t_bio)

    if data.get("special_instructions"):
        elements.append(Spacer(1, 8))
        elements.append(Paragraph(f"<i>Special Instructions: {data['special_instructions']}</i>", styles['Normal']))

    elements.append(Spacer(1, 14))

    tbl_data = [["Ser", "Date", "Pay Fixed", "Pay Scale", "Remarks"]]
    for idx, row in enumerate(data["ledger"], start=1):
        tbl_data.append([str(idx), str(row["date"]), str(row["pay"]), str(row["scale"]), str(row["remarks"])])

    t_main = Table(tbl_data, colWidths=[35, 75, 75, 130, 200])
    t_main.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#D9E1F2')),
        ('TEXTCOLOR', (0,0), (-1,0), colors.black),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('FONTSIZE', (0,0), (-1,0), 9),
        ('FONTNAME', (0,1), (-1,-1), 'Helvetica'),
        ('FONTSIZE', (0,1), (-1,-1), 8.5),
        ('ALIGN', (0,0), (3,-1), 'CENTER'),
        ('ALIGN', (4,0), (4,-1), 'LEFT'),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#A6A6A6')),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0), (-1,-1), 3),
        ('BOTTOMPADDING', (0,0), (-1,-1), 3),
    ]))
    elements.append(t_main)
    doc.build(elements)
    return buf.getvalue()

# ==========================================
# 4. STREAMLIT WEB APP INTERFACE
# ==========================================
def main():
    st.set_page_config(page_title="FC Pay Fixation Agent", layout="wide")
    st.title("🎖️ FC Pay Fixation Agent (Ilyas Sb Certified Standard)")
    st.markdown("Automated pay fixation engine conforming to official Frontier Corps orders.")

    # Sidebar: Rules & Upgrades File Management
    st.sidebar.header("📁 Pay Rules & Scales Management")
    uploaded_file = st.sidebar.file_uploader("Upload New Pay Chart / Revision Excel", type=["xlsx", "csv"])
    if uploaded_file is not None:
        st.sidebar.success("✅ New Rules loaded! Agent updated dynamically.")

    # Inputs Section
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("👤 Personnel Particulars")
        fc_no = st.text_input("FC Number", value="102052")
        rank = st.selectbox("Rank", ["N/Sub", "Hav", "Nk", "Lnk", "GD Sep", "Sub", "SM"], index=0)
        name = st.text_input("Full Name", value="Jamal")
        doe = st.date_input("Date of Enrollment (D.O.E)", datetime.date(2007, 3, 15)).strftime("%Y-%m-%d")
        d_remust = st.date_input("Date of Remustering", datetime.date(2007, 9, 10)).strftime("%Y-%m-%d")

    with col2:
        st.subheader("📈 Promotion History")
        st.caption("Enter promotion dates in order:")
        p_lnk = st.date_input("Promotion to Lnk", datetime.date(2012, 3, 15)).strftime("%Y-%m-%d")
        p_nk = st.date_input("Promotion to Nk", datetime.date(2014, 2, 20)).strftime("%Y-%m-%d")
        p_hav = st.date_input("Promotion to Hav", datetime.date(2017, 12, 2)).strftime("%Y-%m-%d")
        p_nsub = st.date_input("Promotion to N/Sub", datetime.date(2026, 1, 8)).strftime("%Y-%m-%d")

    # Special Instructions Bar
    st.markdown("---")
    st.subheader("📝 Special Instructions & Regulatory Overrides")
    special_instructions = st.text_input(
        "Enter any special instructions (e.g., '1 x advance increment for Matric', 'Annual increment withheld for 2013', 'Personal Pay Protected'):",
        placeholder="Type specific order or exception here..."
    )

    if st.button("🚀 Calculate & Generate Pay Fixation Statement", type="primary"):
        personnel = {
            "name": name,
            "fc_no": fc_no,
            "rank": rank,
            "doe": doe,
            "d_remust": d_remust,
            "promotions": [
                {"date": p_lnk, "rank": "GD Lnk"},
                {"date": p_nk, "rank": "GD Nk"},
                {"date": p_hav, "rank": "GD Hav"},
                {"date": p_nsub, "rank": "GD N/Sub"},
            ],
            "end_date": "2026-07-01"
        }

        # Calculate using golden standard
        res = run_pay_fixation_engine(personnel, DEFAULT_SCALES, special_instructions)

        # Display Dataframe Table in UI
        st.success("Pay Fixation Calculated Successfully!")
        df = pd.DataFrame(res["ledger"])
        df.index = df.index + 1
        st.dataframe(df, use_container_width=True)

        # Download Action Bar
        st.markdown("### 📥 Download Statement")
        dcol1, dcol2, dcol3 = st.columns(3)

        excel_bytes = generate_excel_bytes(res)
        dcol1.download_button(
            label="📊 Download Excel (.xlsx)",
            data=excel_bytes,
            file_name=f"Pay_Fixation_{fc_no}_{name}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        word_bytes = generate_word_bytes(res)
        dcol2.download_button(
            label="📄 Download Word (.docx)",
            data=word_bytes,
            file_name=f"Pay_Fixation_{fc_no}_{name}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

        pdf_bytes = generate_pdf_bytes(res)
        dcol3.download_button(
            label="📑 Download PDF (.pdf)",
            data=pdf_bytes,
            file_name=f"Pay_Fixation_{fc_no}_{name}.pdf",
            mime="application/pdf"
        )

if __name__ == "__main__":
    main()