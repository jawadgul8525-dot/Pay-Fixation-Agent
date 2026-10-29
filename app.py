from __future__ import annotations

import datetime as dt
import io
import re
from pathlib import Path
from typing import Any

import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.shared import Inches, Pt, RGBColor
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle

APP_TITLE = "Pay Fixation Agent"
APP_DESCRIPTION = (
    "Developed by SC Rehman Gul to help HQ Frontier Corps KP (N) troops prepare and review "
    "their basic-pay and pay-fixation calculations independently. Results are a "
    "working aid and must be checked against applicable official orders and service records."
)

# These are the values supplied in the original pasted application. They have NOT
# been independently verified against official orders. Missing scales are not guessed.
DEFAULT_SCALES: dict[str, dict[str, tuple[int, int, int]]] = {
    "GD Sep": {
        "1994-06-01": (1254, 30, 1800), "2001-12-01": (1875, 75, 3525),
        "2005-07-01_REV": (2155, 85, 4025), "2005-07-01_UPG": (2415, 115, 5865),
        "2007-07-01": (2780, 135, 6830), "2008-07-01": (3340, 160, 8140),
        "2011-07-01": (5400, 270, 13500), "2015-07-01": (6985, 350, 17485),
        "2016-07-01": (8590, 430, 21490), "2017-07-01": (10260, 510, 25560),
        "2022-07-01": (15230, 760, 38030), "2024-06-01": (18160, 950, 46660),
        "2026-07-01": (21830, 1150, 56330),
    },
    "GD Lnk": {
        "1994-06-01": (1262, 61, 1933), "2001-12-01": (1895, 95, 3985),
        "2005-07-01_REV": (2180, 110, 4600), "2005-07-01_UPG": (2485, 125, 6235),
        "2007-07-01": (2860, 145, 7210), "2008-07-01": (3430, 175, 8680),
        "2011-07-01": (5500, 290, 14200), "2015-07-01": (7115, 375, 18365),
        "2016-07-01": (8750, 470, 22850), "2017-07-01": (10465, 560, 27265),
        "2022-07-01": (15530, 840, 40730), "2024-06-01": (18780, 1040, 49980),
        "2026-07-01": (22580, 1250, 60080),
    },
    "GD Nk": {
        "1994-06-01": (1400, 66, 2390), "2001-12-01": (2400, 100, 5100),
        "2005-07-01_REV": (2415, 115, 5865), "2005-07-01_UPG": (2555, 140, 6755),
        "2007-07-01": (2940, 160, 7740), "2008-07-01": (3530, 190, 9230),
        "2011-07-01": (5700, 310, 15000), "2015-07-01": (7370, 405, 19520),
        "2016-07-01": (9070, 500, 24070), "2017-07-01": (10835, 595, 28665),
        "2022-07-01": (16080, 890, 42780), "2024-06-01": (19450, 1140, 53650),
        "2026-07-01": (23380, 1380, 64780),
    },
    "GD Hav": {
        "1994-06-01": (1440, 73, 2535), "2001-12-01": (2160, 110, 5460),
        "2005-07-01_REV": (2485, 125, 6235), "2005-07-01_UPG": (2655, 150, 7155),
        "2007-07-01": (3055, 175, 8305), "2008-07-01": (3665, 210, 9965),
        "2011-07-01": (5900, 350, 16400), "2015-07-01": (7630, 452, 21190),
        "2016-07-01": (9390, 560, 26190), "2017-07-01": (11225, 665, 31175),
        "2022-07-01": (16660, 990, 46360), "2024-06-01": (20140, 1240, 57340),
        "2026-07-01": (24210, 1500, 69210),
    },
    "GD N/Sub": {
        "1994-06-01": (1725, 116, 3465), "2001-12-01": (2590, 175, 7840),
        "2005-07-01_REV": (2980, 200, 8980), "2007-07-01": (3430, 230, 10330),
        "2008-07-01": (4115, 275, 12365), "2011-07-01": (6600, 120, 20100),
        "2015-07-01": (8540, 585, 26090), "2016-07-01": (10310, 720, 32110),
        "2017-07-01": (12570, 860, 38370), "2022-07-01": (18650, 1280, 57050),
        "2024-06-01": (26640, 1790, 80340), "2026-07-01": (32030, 2160, 96830),
    },
    "Sub": {
        "1994-06-01": (2338, 144, 4110), "2001-12-01": (3360, 215, 8950),
        "2005-07-01_REV": (3865, 245, 10235), "2007-07-01": (4445, 280, 11725),
        "2008-07-01": (5335, 335, 14045), "2011-07-01": (8600, 550, 25100),
        "2015-07-01": (11120, 710, 32420), "2016-07-01": (13680, 880, 40080),
        "2017-07-01": (16350, 1630, 47850), "2022-07-01": (24260, 1560, 71060),
        "2024-06-01": (30520, 2270, 98620), "2026-07-01": (36690, 2730, 118590),
    },
}

# Display labels map to scale keys. The SM scale is deliberately not fabricated.
RANK_OPTIONS = [
    "GD Sep", "GD Lnk", "GD Nk", "GD Hav", "GD N/Sub", "Sub", "SM", "H/Lt", "H/Capt"
]
HONORARY_RANKS = {"H/Lt", "H/Capt"}
RANK_ALIASES = {
    "Lnk": "GD Lnk", "GD Lnk": "GD Lnk", "Nk": "GD Nk", "GD Nk": "GD Nk",
    "Hav": "GD Hav", "GD Hav": "GD Hav", "N/Sub": "GD N/Sub", "GD N/Sub": "GD N/Sub",
    "Subedar": "Sub", "Sub": "Sub", "Subedar Major": "SM", "SM": "SM",
}

REVISIONS_SCHEDULE = [
    ("2005-07-01_REV", "2005-07-01", "REVISION"),
    ("2005-07-01_UPG", "2005-07-01", "UPGRADATION"),
    ("2007-07-01", "2007-07-01", "REVISION"),
    ("2008-07-01", "2008-07-01", "REVISION"),
    ("2011-07-01", "2011-07-01", "REVISION"),
    ("2015-07-01", "2015-07-01", "REVISION"),
    ("2016-07-01", "2016-07-01", "REVISION"),
    ("2017-07-01", "2017-07-01", "REVISION"),
    ("2022-07-01", "2022-07-01", "REVISION"),
    ("2024-06-01", "2024-06-01", "REVISION"),
    ("2026-07-01", "2026-07-01", "REVISION"),
]


def parse_date(value: Any) -> dt.date | None:
    """Parse an exact date. Year-only values are rejected to avoid inventing a day/month."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    if isinstance(value, dt.datetime):
        return value.date()
    if isinstance(value, dt.date):
        return value
    text = str(value).strip()
    if not text:
        return None
    if re.fullmatch(r"\d{4}", text):
        raise ValueError(f"'{text}' contains only a year. Please select or enter the complete day, month and year.")
    formats = ["%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y", "%d %b %Y", "%d %B %Y", "%d-%b-%Y", "%d.%m.%Y"]
    for fmt in formats:
        try:
            return dt.datetime.strptime(text, fmt).date()
        except ValueError:
            pass
    raise ValueError(f"Could not read date '{text}'. Use YYYY-MM-DD, DD-MM-YYYY, or '15 Jan 2010'.")


def scale_key_date(key: str) -> dt.date:
    return dt.date.fromisoformat(key.split("_")[0])


def format_scale_str(sc: tuple[int, int, int] | None) -> str:
    return f"{sc[0]}-{sc[1]}-{sc[2]}" if sc else "SCALE NOT AVAILABLE"


def get_scale_at_date(rank: str, target_date: dt.date | str, scales_dict: dict, rev_tag: str | None = None):
    rank_key = RANK_ALIASES.get(rank, rank)
    if rank_key not in scales_dict:
        return None
    r_scales = scales_dict[rank_key]
    if rev_tag and rev_tag in r_scales:
        return r_scales[rev_tag]
    target_dt = dt.date.fromisoformat(target_date) if isinstance(target_date, str) else target_date
    eligible = []
    for key, scale in r_scales.items():
        try:
            key_dt = scale_key_date(key)
        except ValueError:
            continue
        if key_dt <= target_dt:
            # At a shared effective date, use UPG after REV where both are present.
            suffix_order = 2 if key.endswith("_UPG") else (1 if key.endswith("_REV") else 0)
            eligible.append((key_dt, suffix_order, scale))
    return sorted(eligible, key=lambda item: (item[0], item[1]))[-1][2] if eligible else None


def merge_uploaded_scales(uploaded_file, current_scales: dict) -> tuple[dict, str]:
    """Load a normalized long-format pay chart. Never claim arbitrary layouts were loaded."""
    if uploaded_file.name.lower().endswith(".csv"):
        frame = pd.read_csv(uploaded_file)
    else:
        frame = pd.read_excel(uploaded_file)
    frame.columns = [str(c).strip().lower().replace("_", " ") for c in frame.columns]
    aliases = {
        "rank": ["rank", "scale rank", "designation"],
        "effective date": ["effective date", "date", "from date", "w.e.f."],
        "initial pay": ["initial pay", "initial", "min pay", "minimum pay"],
        "increment": ["increment", "annual increment", "annual inc"],
        "maximum pay": ["maximum pay", "maximum", "max pay"],
        "event": ["event", "type", "change type"],
    }
    rename = {}
    for standard, options in aliases.items():
        for col in frame.columns:
            if col in options:
                rename[col] = standard
                break
    frame = frame.rename(columns=rename)
    required = {"rank", "effective date", "initial pay", "increment", "maximum pay"}
    if not required.issubset(frame.columns):
        missing = ", ".join(sorted(required - set(frame.columns)))
        raise ValueError(
            "This file layout was not recognized. Required columns: Rank, Effective Date, Initial Pay, Increment, Maximum Pay. "
            f"Missing: {missing}. Use the downloadable template; your file has not been applied."
        )
    updated = {rank: dict(scales) for rank, scales in current_scales.items()}
    rows_loaded = 0
    for _, row in frame.iterrows():
        if pd.isna(row["rank"]) or pd.isna(row["effective date"]):
            continue
        rank = RANK_ALIASES.get(str(row["rank"]).strip(), str(row["rank"]).strip())
        effective = parse_date(row["effective date"])
        initial, increment, maximum = int(row["initial pay"]), int(row["increment"]), int(row["maximum pay"])
        if initial <= 0 or increment <= 0 or maximum < initial:
            raise ValueError(f"Invalid scale values in row {int(_)+2}: check initial pay, increment and maximum pay.")
        event = str(row.get("event", "Revision")).strip().upper() if "event" in frame.columns else "REVISION"
        suffix = "_UPG" if "UPG" in event or "UPGRAD" in event else ("_REV" if effective == dt.date(2005, 7, 1) else "")
        key = effective.isoformat() + suffix
        updated.setdefault(rank, {})[key] = (initial, increment, maximum)
        rows_loaded += 1
    if not rows_loaded:
        raise ValueError("No valid scale rows were found. The uploaded file was not applied.")
    return updated, f"Loaded {rows_loaded} scale row(s) from {uploaded_file.name}. Check the values and effective dates before official use."


def rules_template_bytes() -> bytes:
    frame = pd.DataFrame([
        {"Rank": "GD Sep", "Effective Date": "2024-06-01", "Initial Pay": 18160,
         "Increment": 950, "Maximum Pay": 46660, "Event": "Revision"},
        {"Rank": "GD Sep", "Effective Date": "2005-07-01", "Initial Pay": 2415,
         "Increment": 115, "Maximum Pay": 5865, "Event": "Upgradation"},
    ])
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        frame.to_excel(writer, index=False, sheet_name="Pay Scales")
    return buffer.getvalue()


def run_pay_fixation_engine(personnel: dict, scales_dict: dict, special_instructions: str = "") -> dict:
    issues: list[str] = [
        "The built-in scale figures and calculation rules were supplied in the original app and have not been independently verified against official orders.",
        "The enrollment starting-pay values (1845 before 01-07-2005 and 2128 from that date) are provisional hard-coded values.",
        "Annual increment eligibility is implemented as a provisional 180-day check before the 1 December increment date.",
        "Special instructions are recorded in the report for manual review; natural-language instructions do not automatically alter calculations.",
    ]
    events: list[dict] = []
    doe = personnel["doe"]
    end_date = personnel["end_date"]
    if end_date < doe:
        raise ValueError("End date cannot be earlier than the enrollment date.")
    events.append({"date": doe, "type": "ENROLLMENT", "rank": "GD Sep", "remarks": "Enrolled; starting pay uses provisional configured value"})
    remust = personnel.get("d_remust")
    if remust:
        if remust < doe:
            raise ValueError("Remustering date cannot be earlier than enrollment date.")
        if remust <= end_date:
            events.append({"date": remust, "type": "REMUSTERING", "rank": "GD Sep", "remarks": "Remustered as GD Sep / trained soldier"})
    for promotion in personnel.get("promotions", []):
        pdate, prank = promotion["date"], promotion["rank"]
        if pdate < doe:
            raise ValueError(f"Promotion to {prank} cannot be earlier than enrollment.")
        if pdate > end_date:
            issues.append(f"Promotion to {prank} on {pdate:%d-%m-%Y} is after the selected end date and was not calculated.")
            continue
        if prank in HONORARY_RANKS:
            events.append({"date": pdate, "type": "HONORARY_PROMOTION", "rank": prank,
                           "remarks": f"Honorary promotion to {prank}; no scale change; SM-scale increment rule requires verification"})
        else:
            events.append({"date": pdate, "type": "PROMOTION", "rank": prank, "remarks": f"Promotion to {prank}"})
    for rev_key, rev_date_text, rev_type in REVISIONS_SCHEDULE:
        rev_date = dt.date.fromisoformat(rev_date_text)
        if doe <= rev_date <= end_date:
            events.append({"date": rev_date, "type": rev_type, "rev_key": rev_key, "remarks": rev_type.title()})
    for year in range(doe.year, end_date.year + 1):
        inc_date = dt.date(year, 12, 1)
        if doe <= inc_date <= end_date:
            events.append({"date": inc_date, "type": "ANNUAL_INCREMENT", "remarks": "Annual increment (provisional rule)"})
    priority = {"ENROLLMENT": 0, "REMUSTERING": 1, "REVISION": 2, "UPGRADATION": 3,
                "PROMOTION": 4, "HONORARY_PROMOTION": 5, "ANNUAL_INCREMENT": 6}
    events.sort(key=lambda event: (event["date"], priority.get(event["type"], 99)))

    ledger: list[dict] = []
    current_pay: int | None = None
    current_rank = "GD Sep"
    current_scale = None
    last_change_date: dt.date | None = None

    for event in events:
        event_date, event_type = event["date"], event["type"]
        remarks = event["remarks"]
        if event_type == "ENROLLMENT":
            current_scale = get_scale_at_date("GD Sep", event_date, scales_dict)
            current_pay = 2128 if event_date >= dt.date(2005, 7, 1) else 1845
            last_change_date = event_date
            ledger.append({"date": event_date.strftime("%d-%m-%Y"), "pay": current_pay,
                           "scale": format_scale_str(current_scale), "remarks": remarks})
        elif event_type == "REMUSTERING":
            current_rank = "GD Sep"
            current_scale = get_scale_at_date(current_rank, event_date, scales_dict)
            if current_scale is None:
                ledger.append({"date": event_date.strftime("%d-%m-%Y"), "pay": current_pay,
                               "scale": "SCALE NOT AVAILABLE", "remarks": "HOLD — GD Sep scale unavailable"})
                issues.append(f"No GD Sep scale available for remustering date {event_date:%d-%m-%Y}.")
                continue
            current_pay = current_scale[0]
            last_change_date = event_date
            ledger.append({"date": event_date.strftime("%d-%m-%Y"), "pay": current_pay,
                           "scale": format_scale_str(current_scale), "remarks": remarks})
        elif event_type == "ANNUAL_INCREMENT":
            if current_pay is None or current_scale is None:
                ledger.append({"date": event_date.strftime("%d-%m-%Y"), "pay": current_pay or "—",
                               "scale": format_scale_str(current_scale), "remarks": "HOLD — current scale unavailable"})
                issues.append(f"Annual increment on {event_date:%d-%m-%Y} could not be calculated because the current scale is unavailable.")
                continue
            if last_change_date and (event_date - last_change_date).days < 180:
                ledger.append({"date": event_date.strftime("%d-%m-%Y"), "pay": current_pay,
                               "scale": '"', "remarks": "Not applied: less than 180 days since last rank/pay fixation change (provisional)"})
            else:
                new_pay = current_pay + current_scale[1]
                if new_pay > current_scale[2]:
                    ledger.append({"date": event_date.strftime("%d-%m-%Y"), "pay": current_pay,
                                   "scale": '"', "remarks": "HOLD — increment would exceed configured scale maximum"})
                    issues.append(f"Increment on {event_date:%d-%m-%Y} would exceed the configured maximum of the current scale.")
                else:
                    current_pay = new_pay
                    ledger.append({"date": event_date.strftime("%d-%m-%Y"), "pay": current_pay,
                                   "scale": '"', "remarks": "Annual increment applied (provisional rule)"})
        elif event_type in {"REVISION", "UPGRADATION"}:
            new_scale = get_scale_at_date(current_rank, event_date, scales_dict, rev_tag=event.get("rev_key"))
            if new_scale is None or current_scale is None or current_pay is None:
                ledger.append({"date": event_date.strftime("%d-%m-%Y"), "pay": current_pay or "—",
                               "scale": "SCALE NOT AVAILABLE", "remarks": f"HOLD — {event_type.title()} scale unavailable for {current_rank}"})
                issues.append(f"{event_type.title()} on {event_date:%d-%m-%Y} could not be calculated for {current_rank}; scale data is missing.")
                continue
            if event_type == "REVISION":
                stages = max(0, (current_pay - current_scale[0]) // current_scale[1])
                new_pay = new_scale[0] + stages * new_scale[1]
                if new_pay > new_scale[2]:
                    new_pay = new_scale[2]
                    issues.append(f"Revision on {event_date:%d-%m-%Y} reached the configured scale maximum; verify official fixation method.")
                current_pay = new_pay
            else:
                new_pay = new_scale[0]
                while new_pay < current_pay and new_pay + new_scale[1] <= new_scale[2]:
                    new_pay += new_scale[1]
                if new_pay < current_pay:
                    ledger.append({"date": event_date.strftime("%d-%m-%Y"), "pay": current_pay,
                                   "scale": format_scale_str(new_scale), "remarks": "HOLD — current pay exceeds new scale maximum"})
                    issues.append(f"Upgradation on {event_date:%d-%m-%Y}: current pay exceeds the configured maximum.")
                    continue
                current_pay = new_pay
            current_scale = new_scale
            ledger.append({"date": event_date.strftime("%d-%m-%Y"), "pay": current_pay,
                           "scale": format_scale_str(current_scale), "remarks": event_type.title()})
        elif event_type == "PROMOTION":
            current_rank = event["rank"]
            new_scale = get_scale_at_date(current_rank, event_date, scales_dict)
            if new_scale is None:
                ledger.append({"date": event_date.strftime("%d-%m-%Y"), "pay": current_pay or "—",
                               "scale": "SCALE NOT AVAILABLE", "remarks": f"HOLD — scale for promotion to {current_rank} not configured"})
                issues.append(f"Promotion to {current_rank} on {event_date:%d-%m-%Y} is on HOLD because its scale is not configured. Do not infer an SM scale.")
                continue
            if current_pay is None:
                current_pay = new_scale[0]
            new_pay = new_scale[0]
            while new_pay < current_pay and new_pay + new_scale[1] <= new_scale[2]:
                new_pay += new_scale[1]
            # Promotion formula from the original code is provisional and requires official verification.
            if new_pay + new_scale[1] <= new_scale[2]:
                new_pay += new_scale[1]
            else:
                issues.append(f"Promotion to {current_rank} on {event_date:%d-%m-%Y} may exceed the scale maximum; verify the promotion formula.")
            current_pay = new_pay
            current_scale = new_scale
            last_change_date = event_date
            ledger.append({"date": event_date.strftime("%d-%m-%Y"), "pay": current_pay,
                           "scale": format_scale_str(current_scale), "remarks": remarks + " (provisional fixation formula)"})
        elif event_type == "HONORARY_PROMOTION":
            # User-specified behavior: no scale change, add one SM-scale increment.
            sm_scale = get_scale_at_date("SM", event_date, scales_dict)
            if sm_scale is None or current_pay is None or current_scale is None:
                ledger.append({"date": event_date.strftime("%d-%m-%Y"), "pay": current_pay or "—",
                               "scale": format_scale_str(current_scale),
                               "remarks": f"HOLD — {event['rank']}: SM scale is not configured; no increment calculated"})
                issues.append(f"Honorary promotion to {event['rank']} on {event_date:%d-%m-%Y} is on HOLD because the SM scale is missing. Upload/verify the official SM scale.")
                continue
            proposed = current_pay + sm_scale[1]
            if proposed > current_scale[2]:
                ledger.append({"date": event_date.strftime("%d-%m-%Y"), "pay": current_pay,
                               "scale": format_scale_str(current_scale),
                               "remarks": f"HOLD — {event['rank']}: SM increment would exceed current scale maximum"})
                issues.append(f"Honorary promotion to {event['rank']} on {event_date:%d-%m-%Y} would exceed the current scale maximum; verify treatment.")
                continue
            if current_rank != "SM":
                ledger.append({"date": event_date.strftime("%d-%m-%Y"), "pay": current_pay,
                               "scale": format_scale_str(current_scale),
                               "remarks": f"HOLD — {event['rank']}: honorary promotion rule is configured for personnel already on SM rank/scale"})
                issues.append(f"Honorary promotion to {event['rank']} on {event_date:%d-%m-%Y} was not applied because the current recorded rank is {current_rank}, not SM. Verify the service history.")
                continue
            current_pay = proposed
            last_change_date = event_date
            # Keep the underlying rank as SM so later revisions continue to use the SM scale.
            ledger.append({"date": event_date.strftime("%d-%m-%Y"), "pay": current_pay,
                           "scale": format_scale_str(current_scale),
                           "remarks": f"{event['rank']} — one SM-scale increment; no scale change (provisional rule)"})

    if special_instructions.strip():
        issues.append("Special instructions were recorded in the reports but were NOT automatically applied to pay calculations. Review them manually.")
    return {"header": personnel, "ledger": ledger, "special_instructions": special_instructions.strip(),
            "issues": list(dict.fromkeys(issues)), "status": "HOLD — REVIEW REQUIRED" if issues else "CALCULATED"}


def generate_excel_bytes(data: dict) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "PAY FIXATION"
    ws.merge_cells("A1:E1")
    h = data["header"]
    ws["A1"] = f"PAY FIXATION IRO FC NO. {h.get('fc_no','')} {h.get('rank','')} {h.get('name','').upper()}"
    ws["A1"].font = Font(name="Calibri", size=12, bold=True)
    ws["A1"].alignment = Alignment(horizontal="center")
    ws.merge_cells("A2:E2")
    ws["A2"] = f"Status: {data.get('status', 'REVIEW REQUIRED')}"
    ws["A2"].font = Font(bold=True, color="9C0006")
    bio = [("a", "Date of Enrollment", h.get("doe")), ("b", "Date of Remustering", h.get("d_remust") or "Not provided")]
    for index, p in enumerate(h.get("promotions", []), start=3):
        bio.append((chr(96 + index), f"Promotion: {p['rank']}", p["date"].strftime("%d-%m-%Y")))
    row_num = 4
    for mark, label, value in bio:
        ws.cell(row=row_num, column=1, value=mark)
        ws.cell(row=row_num, column=2, value=label)
        ws.cell(row=row_num, column=3, value=str(value))
        row_num += 1
    if data.get("special_instructions"):
        row_num += 1
        ws.cell(row=row_num, column=1, value="Special instructions (manual review):")
        ws.merge_cells(start_row=row_num, start_column=2, end_row=row_num, end_column=5)
        ws.cell(row=row_num, column=2, value=data["special_instructions"])
        row_num += 1
    if data.get("issues"):
        row_num += 1
        ws.cell(row=row_num, column=1, value="Validation / Review Required")
        ws.cell(row=row_num, column=1).font = Font(bold=True, color="9C0006")
        row_num += 1
        for issue in data["issues"]:
            ws.merge_cells(start_row=row_num, start_column=1, end_row=row_num, end_column=5)
            ws.cell(row=row_num, column=1, value=f"• {issue}")
            row_num += 1
    row_num += 1
    headers = ["Ser", "Date", "Pay Fixed", "Pay Scale", "Remarks"]
    border = Border(*( [Side(style="thin", color="A6A6A6")] * 4 ))
    for col, title in enumerate(headers, 1):
        cell = ws.cell(row=row_num, column=col, value=title)
        cell.font = Font(bold=True)
        cell.fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
        cell.alignment = Alignment(horizontal="center")
        cell.border = border
    row_num += 1
    for serial, item in enumerate(data["ledger"], 1):
        values = [serial, item["date"], item["pay"], item["scale"], item["remarks"]]
        for col, value in enumerate(values, 1):
            cell = ws.cell(row=row_num, column=col, value=value)
            cell.border = border
            cell.alignment = Alignment(vertical="top", wrap_text=True, horizontal="center" if col < 5 else "left")
        row_num += 1
    for col, width in {"A": 8, "B": 14, "C": 14, "D": 24, "E": 58}.items():
        ws.column_dimensions[col].width = width
    ws.freeze_panes = "A" + str(row_num - len(data["ledger"]))
    ws.auto_filter.ref = f"A{row_num-len(data['ledger'])}:E{max(row_num-1, row_num-len(data['ledger']))}"
    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue()


def generate_word_bytes(data: dict) -> bytes:
    doc = Document()
    h = data["header"]
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.add_run(f"PAY FIXATION IRO FC NO. {h.get('fc_no','')} {h.get('rank','')} {h.get('name','').upper()}")
    run.bold = True
    run.font.size = Pt(12)
    doc.add_paragraph(f"Status: {data.get('status', 'REVIEW REQUIRED')}")
    table = doc.add_table(rows=0, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    particulars = [("Date of Enrollment", h.get("doe")), ("Date of Remustering", h.get("d_remust") or "Not provided")]
    particulars += [(f"Promotion: {p['rank']}", p['date'].strftime("%d-%m-%Y")) for p in h.get("promotions", [])]
    for label, value in particulars:
        cells = table.add_row().cells
        cells[0].text, cells[1].text = str(label), str(value)
    if data.get("special_instructions"):
        doc.add_paragraph("Special instructions (recorded for manual review): " + data["special_instructions"])
    if data.get("issues"):
        doc.add_heading("Validation / Review Required", level=2)
        for issue in data["issues"]:
            doc.add_paragraph(issue, style="List Bullet")
    doc.add_heading("Pay Fixation Ledger", level=2)
    main = doc.add_table(rows=1, cols=5)
    main.style = "Table Grid"
    for i, label in enumerate(["Ser", "Date", "Pay Fixed", "Pay Scale", "Remarks"]):
        main.rows[0].cells[i].text = label
        shading = parse_xml(r'<w:shd {} w:fill="D9E1F2"/>'.format(nsdecls("w")))
        main.rows[0].cells[i]._tc.get_or_add_tcPr().append(shading)
    for i, item in enumerate(data["ledger"], 1):
        cells = main.add_row().cells
        for j, value in enumerate([i, item["date"], item["pay"], item["scale"], item["remarks"]]):
            cells[j].text = str(value)
            for paragraph in cells[j].paragraphs:
                for run in paragraph.runs:
                    run.font.size = Pt(8.5)
    buffer = io.BytesIO()
    doc.save(buffer)
    return buffer.getvalue()


def generate_pdf_bytes(data: dict) -> bytes:
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=landscape(A4), leftMargin=24, rightMargin=24, topMargin=28, bottomMargin=28)
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name="SmallCell", parent=styles["BodyText"], fontSize=7.5, leading=9))
    elements = []
    h = data["header"]
    elements.append(Paragraph(f"<b>PAY FIXATION IRO FC NO. {h.get('fc_no','')} {h.get('rank','')} {h.get('name','').upper()}</b>", styles["Title"]))
    elements.append(Paragraph(f"Status: {data.get('status', 'REVIEW REQUIRED')}", styles["Normal"]))
    elements.append(Spacer(1, 8))
    bio = [["Particular", "Value"], ["Date of Enrollment", str(h.get("doe"))], ["Date of Remustering", str(h.get("d_remust") or "Not provided")]]
    bio.extend([[f"Promotion: {p['rank']}", p['date'].strftime("%d-%m-%Y")] for p in h.get("promotions", [])])
    elements.append(Table(bio, colWidths=[150, 130], style=TableStyle([
        ("GRID", (0,0), (-1,-1), .4, colors.grey), ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#D9E1F2")),
        ("FONTSIZE", (0,0), (-1,-1), 8), ("VALIGN", (0,0), (-1,-1), "TOP")
    ])))
    if data.get("special_instructions"):
        elements += [Spacer(1, 6), Paragraph("<b>Special instructions (manual review):</b> " + data["special_instructions"].replace("&", "&amp;"), styles["Normal"])]
    if data.get("issues"):
        elements += [Spacer(1, 6), Paragraph("<b>Validation / Review Required</b>", styles["Heading2"])]
        for issue in data["issues"]:
            elements.append(Paragraph("• " + issue.replace("&", "&amp;"), styles["SmallCell"]))
    elements += [Spacer(1, 10), Paragraph("<b>Pay Fixation Ledger</b>", styles["Heading2"])]
    table_data = [["Ser", "Date", "Pay Fixed", "Pay Scale", "Remarks"]]
    for i, item in enumerate(data["ledger"], 1):
        table_data.append([str(i), item["date"], str(item["pay"]), item["scale"], item["remarks"]])
    table = Table(table_data, colWidths=[30, 65, 65, 110, 440], repeatRows=1)
    table.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#D9E1F2")),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"), ("FONTSIZE", (0,0), (-1,0), 8),
        ("FONTSIZE", (0,1), (-1,-1), 7.2), ("GRID", (0,0), (-1,-1), .35, colors.HexColor("#A6A6A6")),
        ("VALIGN", (0,0), (-1,-1), "TOP"), ("LEFTPADDING", (0,0), (-1,-1), 4),
        ("RIGHTPADDING", (0,0), (-1,-1), 4), ("TOPPADDING", (0,0), (-1,-1), 3),
        ("BOTTOMPADDING", (0,0), (-1,-1), 3),
    ]))
    elements.append(table)
    doc.build(elements)
    return buffer.getvalue()


def main() -> None:
    st.set_page_config(page_title=APP_TITLE, page_icon="🎖️", layout="wide")
    st.title(APP_TITLE)
    st.markdown(APP_DESCRIPTION)
    st.caption("For review and assistance only — verify all results against official orders before official use.")

    local_logo = Path(__file__).with_name("logo.png")
    if local_logo.exists():
        st.image(str(local_logo), width=130)

    scales_dict = {rank: dict(scales) for rank, scales in DEFAULT_SCALES.items()}
    st.sidebar.header("Pay Rules & App Logo")
    st.sidebar.caption("Uploaded pay charts are parsed only if they use the supported column format. Unsupported files are not applied.")
    uploaded_rules = st.sidebar.file_uploader("Upload pay-scale chart (.xlsx or .csv)", type=["xlsx", "csv"], key="rules_upload")
    if uploaded_rules is not None:
        try:
            scales_dict, upload_message = merge_uploaded_scales(uploaded_rules, scales_dict)
            st.sidebar.success(upload_message)
        except Exception as exc:
            st.sidebar.error(str(exc))
    st.sidebar.download_button("Download pay-scale upload template", rules_template_bytes(),
                               file_name="Pay_Scale_Upload_Template.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    uploaded_logo = st.sidebar.file_uploader("Upload your logo for this session (PNG/JPG)", type=["png", "jpg", "jpeg"], key="logo_upload")
    if uploaded_logo is not None:
        st.image(uploaded_logo, width=130, caption="Logo preview (session only; add logo.png to GitHub to bundle it permanently)")

    with st.form("pay_fixation_form"):
        left, right = st.columns(2)
        with left:
            st.subheader("Personnel Particulars")
            fc_no = st.text_input("FC Number", value="")
            name = st.text_input("Full Name", value="")
            starting_rank = st.selectbox("Current / reported rank", RANK_OPTIONS, index=0)
            doe_text = st.text_input("Date of Enrollment (complete date)", value="15 Mar 2007", help="Examples: 2007-03-15, 15-03-2007, 15 Jan 2007. Year-only values are not used because they do not identify an exact effective date.")
            remust_text = st.text_input("Date of Remustering (optional)", value="10 Sep 2007", help="Leave blank if not applicable. Historical dates are accepted.")
            end_text = st.text_input("Calculate through date", value="01 Jul 2026", help="Enter an exact date. Historical dates are accepted.")
        with right:
            st.subheader("Promotion History")
            st.caption("Enter only promotions that apply. Leave the date blank for ranks not attained. Old dates are accepted; e.g. 15 Jan 2010 or 15-01-2000.")
            promotion_inputs = []
            promotion_ranks = ["GD Lnk", "GD Nk", "GD Hav", "GD N/Sub", "Sub", "SM", "H/Lt", "H/Capt"]
            for prank in promotion_ranks:
                date_value = st.text_input(f"Promotion to {prank} — date (optional)", value="", key=f"promo_{prank}",
                                           placeholder="e.g. 15 Jan 2010")
                if date_value.strip():
                    promotion_inputs.append((prank, date_value.strip()))
        st.subheader("Special Instructions & Regulatory Overrides")
        special_instructions = st.text_area(
            "Enter relevant instructions / order references",
            placeholder="Example: Annual increment withheld for 2013; cite the relevant order if available.",
            help="Instructions are included in all reports and flagged for manual review. They do not automatically change pay calculations."
        )
        calculate = st.form_submit_button("Calculate & Generate Pay Fixation Statement", type="primary")

    if calculate:
        errors = []
        try:
            doe = parse_date(doe_text)
            if doe is None:
                errors.append("Date of enrollment is required.")
        except ValueError as exc:
            errors.append(str(exc))
            doe = None
        try:
            end_date = parse_date(end_text)
            if end_date is None:
                errors.append("Calculate-through date is required.")
        except ValueError as exc:
            errors.append(str(exc))
            end_date = None
        try:
            remust = parse_date(remust_text) if remust_text.strip() else None
        except ValueError as exc:
            errors.append(str(exc))
            remust = None
        promotions = []
        for prank, text_value in promotion_inputs:
            try:
                pdate = parse_date(text_value)
                if pdate is not None:
                    promotions.append({"date": pdate, "rank": prank})
            except ValueError as exc:
                errors.append(f"Promotion to {prank}: {exc}")
        if doe and end_date and end_date < doe:
            errors.append("Calculate-through date cannot be earlier than enrollment.")
        if doe:
            for promotion in promotions:
                if promotion["date"] < doe:
                    errors.append(f"Promotion to {promotion['rank']} is earlier than enrollment.")
        if errors:
            for error in errors:
                st.error(error)
            st.stop()
        promotions.sort(key=lambda p: p["date"])
        personnel = {
            "name": name, "fc_no": fc_no, "rank": starting_rank, "doe": doe,
            "d_remust": remust, "promotions": promotions, "end_date": end_date,
        }
        try:
            result = run_pay_fixation_engine(personnel, scales_dict, special_instructions)
        except Exception as exc:
            st.error(f"Calculation stopped: {exc}")
            st.stop()
        st.warning("This report is a draft calculation. Review all listed issues and confirm the applicable official rules before using it for pay or pension purposes.")
        st.subheader(f"Calculation status: {result['status']}")
        st.dataframe(pd.DataFrame(result["ledger"]), use_container_width=True, hide_index=True)
        if result.get("issues"):
            with st.expander("Validation notes / items requiring review", expanded=True):
                for issue in result["issues"]:
                    st.markdown(f"- {issue}")
        if special_instructions.strip():
            st.info("Special instructions have been included in the reports for manual review; they were not automatically applied to the calculations.")
        st.markdown("### Download Statement")
        c1, c2, c3 = st.columns(3)
        safe_fc = re.sub(r"[^A-Za-z0-9_-]+", "_", fc_no or "FC")
        safe_name = re.sub(r"[^A-Za-z0-9_-]+", "_", name or "Personnel")
        c1.download_button("Download Excel (.xlsx)", generate_excel_bytes(result),
                           file_name=f"Pay_Fixation_{safe_fc}_{safe_name}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        c2.download_button("Download Word (.docx)", generate_word_bytes(result),
                           file_name=f"Pay_Fixation_{safe_fc}_{safe_name}.docx",
                           mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        c3.download_button("Download PDF (.pdf)", generate_pdf_bytes(result),
                           file_name=f"Pay_Fixation_{safe_fc}_{safe_name}.pdf", mime="application/pdf")


if __name__ == "__main__":
    main()
