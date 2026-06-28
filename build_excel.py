#!/usr/bin/env python3
"""Rebuild HARIOM LOGISTICS Excel with dropdowns, owner ledger, and auto-billing."""

from __future__ import annotations

import json
from copy import copy
from datetime import datetime
from pathlib import Path

import openpyxl
import pandas as pd
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo

SOURCE = Path("/workspace/vehicles freight details.xlsx")
OUTPUT = Path("/workspace/HARIOM LOGISTICS - Automated.xlsx")

TRIP_HEADER_ROW = 3
TRIP_DATA_START = 4
TRIP_DATA_END = 502
PAY_DATA_START = 4
PAY_DATA_END = 203
MASTER_ROW_MAX = 500
TRUCK_TABLE_END = 150
TRIP_RANGE = f"'Trip Log'!$A${TRIP_DATA_START}:$T${TRIP_DATA_END}"

# Styles
HEADER_FILL = PatternFill("solid", fgColor="1F4E79")
HEADER_FONT = Font(bold=True, color="FFFFFF", size=11)
TITLE_FONT = Font(bold=True, size=14, color="1F4E79")
SUBTITLE_FONT = Font(bold=True, size=11, color="1F4E79")
SECTION_FILL = PatternFill("solid", fgColor="D9E2F3")
MONEY_FMT = "#,##0"
DATE_FMT = "DD-MMM-YYYY"
THIN_BORDER = Border(
    left=Side(style="thin", color="B4B4B4"),
    right=Side(style="thin", color="B4B4B4"),
    top=Side(style="thin", color="B4B4B4"),
    bottom=Side(style="thin", color="B4B4B4"),
)
ALERT_FILL = PatternFill(patternType="solid", fgColor="FFC7CE", bgColor="FFC7CE")
WARN_FILL = PatternFill(patternType="solid", fgColor="FFEB9C", bgColor="FFEB9C")
STRIPE_FILL = PatternFill(patternType="solid", fgColor="F2F7FB", bgColor="F2F7FB")


def add_trip_highlight_rules(ws, first_row: int, last_row: int) -> None:
    """Color-code trip rows. Uses ISBLANK (not date checks) so rules work in all Excel versions."""
    range_ref = f"A{first_row}:T{last_row}"
    anchor = first_row

    # Priority 1 — red: trip started but Bill Ref missing
    red = FormulaRule(
        formula=[f'AND($A{anchor}<>"",$G{anchor}<>"",ISBLANK($Q{anchor}))'],
        fill=ALERT_FILL,
        stopIfTrue=True,
    )
    red.priority = 1
    ws.conditional_formatting.add(range_ref, red)

    # Priority 2 — yellow: trip started but PAHUNCH missing
    yellow = FormulaRule(
        formula=[f'AND($A{anchor}<>"",$G{anchor}<>"",ISBLANK($P{anchor}))'],
        fill=WARN_FILL,
        stopIfTrue=True,
    )
    yellow.priority = 2
    ws.conditional_formatting.add(range_ref, yellow)

    # Priority 3 — subtle zebra stripes on data rows
    stripe = FormulaRule(
        formula=[f'AND(MOD(ROW(),2)=0,$A{anchor}<>"")'],
        fill=STRIPE_FILL,
        stopIfTrue=False,
    )
    stripe.priority = 3
    ws.conditional_formatting.add(range_ref, stripe)


def add_owner_ledger_highlight_rules(ws, first_row: int, last_row: int) -> None:
    range_ref = f"A{first_row}:J{last_row}"
    due_rule = FormulaRule(formula=[f'$H{first_row}>0'], fill=WARN_FILL, stopIfTrue=False)
    due_rule.priority = 1
    ws.conditional_formatting.add(range_ref, due_rule)


def add_freight_bill_highlight_rules(ws, first_row: int, last_row: int) -> None:
    range_ref = f"A{first_row}:L{last_row}"
    anchor = first_row
    epod_rule = FormulaRule(formula=[f'$G{anchor}="NO"'], fill=WARN_FILL, stopIfTrue=False)
    epod_rule.priority = 1
    ws.conditional_formatting.add(range_ref, epod_rule)

    pending_rule = FormulaRule(formula=[f'$L{anchor}="PENDING"'], fill=ALERT_FILL, stopIfTrue=False)
    pending_rule.priority = 2
    ws.conditional_formatting.add(range_ref, pending_rule)


def style_header_row(ws, row: int, max_col: int) -> None:
    for col in range(1, max_col + 1):
        cell = ws.cell(row=row, column=col)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = THIN_BORDER


def style_title(ws, cell_ref: str, text: str) -> None:
    ws[cell_ref] = text
    ws[cell_ref].font = TITLE_FONT


def add_list_validation(ws, cell_range: str, defined_name: str) -> None:
    """Add a dropdown list backed by a workbook defined name (most reliable in Excel)."""
    formula = defined_name if defined_name.startswith("=") else f"={defined_name}"
    dv = DataValidation(type="list", formula1=formula, allow_blank=True)
    dv.showDropDown = False  # Excel quirk: False = show the dropdown arrow
    dv.showInputMessage = True
    dv.showErrorMessage = True
    dv.error = "Please choose a value from the dropdown list."
    dv.errorTitle = "Invalid entry"
    dv.prompt = "Click the arrow and select from the list"
    dv.promptTitle = "Dropdown"
    ws.add_data_validation(dv)
    dv.add(cell_range)


def lookup_truck(truck_cell: str, result_col: str, fallback: str) -> str:
    """INDEX/MATCH lookup — works in Excel 2007+ (unlike XLOOKUP)."""
    return (
        f'=IF({truck_cell}="","",IFERROR('
        f"INDEX('Master Data'!${result_col}$5:${result_col}${TRUCK_TABLE_END},"
        f"MATCH({truck_cell},'Master Data'!$E$5:$E${TRUCK_TABLE_END},0)),"
        f"{fallback}))"
    )


def trip_formulas(row: int) -> dict[int, str]:
    """Return formula columns for one trip row."""
    return {
        6: lookup_truck(f"E{row}", "F", '""'),
        9: f'=IF(OR(G{row}="",H{row}=""),"",G{row}*H{row})',
        12: f'=IF(I{row}="","",I{row}-IF(J{row}="",0,J{row})-IF(K{row}="",0,K{row}))',
        13: lookup_truck(f"E{row}", "G", "0"),
        14: lookup_truck(f"E{row}", "H", "0"),
        15: f'=IF(L{row}="","",L{row}-M{row}-N{row})',
    }


def create_defined_names(wb: openpyxl.Workbook) -> dict[str, str]:
    """Register named ranges used by dropdowns. Dynamic OFFSET ranges grow as lists expand."""
    names = {
        "HL_Parties": (
            f"OFFSET('Master Data'!$A$5,0,0,"
            f"COUNTA('Master Data'!$A$5:$A${MASTER_ROW_MAX}),1)"
        ),
        "HL_Destinations": (
            f"OFFSET('Master Data'!$C$5,0,0,"
            f"COUNTA('Master Data'!$C$5:$C${MASTER_ROW_MAX}),1)"
        ),
        "HL_Trucks": (
            f"OFFSET('Master Data'!$E$5,0,0,"
            f"COUNTA('Master Data'!$E$5:$E${MASTER_ROW_MAX}),1)"
        ),
        "HL_Owners": (
            f"OFFSET('Master Data'!$J$5,0,0,"
            f"COUNTA('Master Data'!$J$5:$J${MASTER_ROW_MAX}),1)"
        ),
        "HL_Pahunch": "'Master Data'!$L$5:$L$5",
        "HL_EPOD": "'Master Data'!$M$5:$M$6",
        "HL_BillStatus": "'Master Data'!$N$5:$N$7",
        "HL_BillRefs": (
            f"OFFSET('Master Data'!$P$5,0,0,"
            f"COUNTA('Master Data'!$P$5:$P${MASTER_ROW_MAX}),1)"
        ),
    }
    for name, ref in names.items():
        wb.defined_names.add(DefinedName(name=name, attr_text=ref))
    return names


def normalize_bill_ref(value) -> str | None:
    if pd.isna(value):
        return None
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    if isinstance(value, int):
        return str(value)
    text = str(value).strip()
    if text.endswith(".0"):
        text = text[:-2]
    return text or None


def bill_ref_mapping() -> dict[str, str]:
    """Best-match links between formal client bills and trip batch refs."""
    return {
        "HO/JUN/01": "1 EPOD",
        "HO/JUN/02": "2 EPOD",
        "HO/JUN/03": "3",
        "HO/JUN/04": "4 EPOD",
        "HO/JUN/05": "5",
        "HO/JUN/06": "6",
        "HO/JUN/07": "7 EPOD",
        "HO/JUN/08": "8 EPOD",
        "HO/JUN/09": "9 EPOD",
        "HO/JUN/10": "10",
        "HO/JUN/11": "11",
    }


def build_trucks_master(vd: pd.DataFrame) -> pd.DataFrame:
    """Build truck master including every truck in the trip log."""
    rows: list[dict] = []
    for truck, sub in vd.groupby("Truck No", dropna=True):
        truck_str = str(truck).strip()
        if not truck_str:
            continue
        owners = sub["OWNER NAME"].dropna()
        owner = owners.mode().iloc[0] if len(owners) else ""
        comm = sub["commition"].dropna()
        muns = sub["Munsiyana"].dropna()
        rows.append(
            {
                "Truck No": truck_str,
                "OWNER NAME": str(owner).strip() if pd.notna(owner) else "",
                "commition": float(comm.median()) if len(comm) else 0,
                "Munsiyana": float(muns.median()) if len(muns) else 0,
            }
        )
    return pd.DataFrame(rows).sort_values(["OWNER NAME", "Truck No"], na_position="last")


def load_source_data() -> dict:
    vd = pd.read_excel(SOURCE, sheet_name="vehicle details", header=1)
    vd.columns = [str(c).strip() for c in vd.columns]
    vd = vd.dropna(how="all")
    vd = vd[vd["LR No"].notna()].copy()
    vd["Date"] = pd.to_datetime(vd["Date"], errors="coerce")

    pay = pd.read_excel(SOURCE, sheet_name="PAYMENT", header=1)
    pay.columns = [str(c).strip() for c in pay.columns]
    pay = pay.dropna(how="all").copy()

    bills = pd.read_excel(SOURCE, sheet_name="Sheet1", header=2)
    bills.columns = [str(c).strip() for c in bills.columns]
    bills = bills[bills["BILL No."].notna()].copy()

    trucks_df = build_trucks_master(vd)

    owners = sorted(
        {
            str(o).strip()
            for o in pd.concat([vd["OWNER NAME"], trucks_df["OWNER NAME"]])
            if pd.notna(o) and str(o).strip()
        }
    )

    return {
        "trips": vd,
        "payments": pay,
        "bills": bills,
        "parties": sorted(vd["Sold To Party"].dropna().unique().tolist()),
        "destinations": sorted(vd["Destination"].dropna().unique().tolist()),
        "trucks": trucks_df,
        "owners": owners,
        "bill_refs": sorted(
            {r for x in vd["BILL No"].dropna().unique().tolist() if (r := normalize_bill_ref(x))},
            key=lambda x: (len(x), x),
        ),
    }


def build_master_data(wb: openpyxl.Workbook, data: dict) -> openpyxl.worksheet.worksheet.Worksheet:
    ws = wb.active
    ws.title = "Master Data"
    ws.sheet_view.showGridLines = True

    style_title(ws, "A1", "HARIOM LOGISTICS — MASTER DATA")
    ws.merge_cells("A1:H1")

    # Parties
    ws["A3"] = "PARTIES (Sold To)"
    ws["A3"].font = SUBTITLE_FONT
    ws["A3"].fill = SECTION_FILL
    ws["A4"] = "Party Name"
    style_header_row(ws, 4, 1)
    for i, party in enumerate(data["parties"], start=5):
        ws.cell(row=i, column=1, value=party)
    party_end = 4 + len(data["parties"])

    # Destinations
    ws["C3"] = "DESTINATIONS"
    ws["C3"].font = SUBTITLE_FONT
    ws["C3"].fill = SECTION_FILL
    ws["C4"] = "Destination"
    style_header_row(ws, 4, 3)
    for i, dest in enumerate(data["destinations"], start=5):
        ws.cell(row=i, column=3, value=dest)
    dest_end = 4 + len(data["destinations"])

    # Trucks
    ws["E3"] = "TRUCKS (add owner & commission for new trucks)"
    ws["E3"].font = SUBTITLE_FONT
    ws["E3"].fill = SECTION_FILL
    headers = ["Truck No", "Owner", "Commission", "Munsiyana"]
    for col, h in enumerate(headers, start=5):
        ws.cell(row=4, column=col, value=h)
    style_header_row(ws, 4, 8)
    truck_start = 5
    for i, row in enumerate(data["trucks"].itertuples(index=False), start=truck_start):
        ws.cell(row=i, column=5, value=row[0])
        ws.cell(row=i, column=6, value=row[1])
        ws.cell(row=i, column=7, value=float(row[2]) if pd.notna(row[2]) else 0)
        ws.cell(row=i, column=8, value=float(row[3]) if pd.notna(row[3]) else 0)
        ws.cell(row=i, column=7).number_format = MONEY_FMT
        ws.cell(row=i, column=8).number_format = MONEY_FMT
    truck_end = truck_start + len(data["trucks"]) - 1

    # Owners
    ws["J3"] = "OWNERS"
    ws["J3"].font = SUBTITLE_FONT
    ws["J3"].fill = SECTION_FILL
    ws["J4"] = "Owner Name"
    style_header_row(ws, 4, 10)
    for i, owner in enumerate(data["owners"], start=5):
        ws.cell(row=i, column=10, value=owner)
    owner_end = 4 + len(data["owners"])

    # Status lists
    ws["L3"] = "STATUS LISTS"
    ws["L3"].font = SUBTITLE_FONT
    ws["L3"].fill = SECTION_FILL
    ws["L4"], ws["M4"], ws["N4"] = "PAHUNCH", "EPOD STATUS", "BILL STATUS"
    style_header_row(ws, 4, 14)
    pahunch_vals = ["YES"]
    epod_vals = ["YES", "NO"]
    bill_status_vals = ["PENDING", "PASSED", "PARTIAL"]
    for i, v in enumerate(pahunch_vals, start=5):
        ws.cell(row=i, column=12, value=v)
    for i, v in enumerate(epod_vals, start=5):
        ws.cell(row=i, column=13, value=v)
    for i, v in enumerate(bill_status_vals, start=5):
        ws.cell(row=i, column=14, value=v)

    # Trip bill refs
    ws["P3"] = "TRIP BILL REFS"
    ws["P3"].font = SUBTITLE_FONT
    ws["P3"].fill = SECTION_FILL
    ws["P4"] = "Bill Ref"
    style_header_row(ws, 4, 16)
    for i, ref in enumerate(data["bill_refs"], start=5):
        ws.cell(row=i, column=16, value=ref)
    bill_ref_end = 4 + len(data["bill_refs"])

    # Named range helpers in row 2 (hidden metadata for formulas)
    ws["A2"] = f"Parties!$A$5:$A${party_end}"
    ws["C2"] = f"Destinations!$C$5:$C${dest_end}"
    ws["E2"] = f"Trucks!$E$5:$E${truck_end}"
    ws["J2"] = f"Owners!$J$5:$J${owner_end}"
    ws["L2"] = "$L$5:$L$5"
    ws["P2"] = f"BillRefs!$P$5:$P${bill_ref_end}"

    # Column widths
    widths = {"A": 36, "B": 3, "C": 28, "D": 3, "E": 14, "F": 14, "G": 12, "H": 12, "J": 16, "L": 12, "M": 12, "N": 12, "P": 14}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w

    ws.row_dimensions[2].hidden = True
    return ws, {
        "party_range": "HL_Parties",
        "dest_range": "HL_Destinations",
        "truck_range": "HL_Trucks",
        "owner_range": "HL_Owners",
        "pahunch_range": "HL_Pahunch",
        "epod_range": "HL_EPOD",
        "bill_status_range": "HL_BillStatus",
        "bill_ref_range": "HL_BillRefs",
        "truck_table_start": truck_start,
        "truck_table_end": truck_end,
        "owner_start": 5,
        "owner_end": owner_end,
    }


def build_trip_log(wb: openpyxl.Workbook, data: dict, ranges: dict) -> tuple:
    ws = wb.create_sheet("Trip Log")
    style_title(ws, "A1", "HARIOM LOGISTICS — TRIP LOG")
    ws.merge_cells("A1:T1")

    headers = [
        "LR No",
        "Date",
        "Sold To Party",
        "Destination",
        "Truck No",
        "Owner Name",
        "Quantity",
        "Rate",
        "Freight",
        "Diesel",
        "Cash",
        "Total Advance",
        "Commission",
        "Munsiyana",
        "Balance",
        "PAHUNCH",
        "Bill Ref",
        "Payment Date",
        "Payment Amt",
        "Particulars",
    ]
    header_row = 3
    for col, h in enumerate(headers, start=1):
        ws.cell(row=header_row, column=col, value=h)
    style_header_row(ws, header_row, len(headers))

    first_data_row = header_row + 1
    max_rows = 500
    last_data_row = first_data_row + max_rows - 1

    trips = data["trips"]
    for idx, trip in enumerate(trips.itertuples(index=False)):
        r = first_data_row + idx
        ws.cell(row=r, column=1, value=trip[0])  # LR No
        if pd.notna(trip[1]):
            ws.cell(row=r, column=2, value=trip[1].to_pydatetime() if hasattr(trip[1], "to_pydatetime") else trip[1])
            ws.cell(row=r, column=2).number_format = DATE_FMT
        ws.cell(row=r, column=3, value=trip[2])  # Party
        ws.cell(row=r, column=4, value=trip[3])  # Destination
        ws.cell(row=r, column=5, value=trip[4])  # Truck
        for col, formula in trip_formulas(r).items():
            ws.cell(row=r, column=col, value=formula)
        ws.cell(row=r, column=7, value=float(trip[6]) if pd.notna(trip[6]) else None)
        ws.cell(row=r, column=8, value=float(trip[7]) if pd.notna(trip[7]) else None)
        if pd.notna(trip[9]):
            ws.cell(row=r, column=10, value=float(trip[9]))
        if pd.notna(trip[10]):
            ws.cell(row=r, column=11, value=float(trip[10]))
        if pd.notna(trip[15]) and str(trip[15]).strip().upper() == "YES":
            ws.cell(row=r, column=16, value="YES")
        if pd.notna(trip[16]):
            ws.cell(row=r, column=17, value=normalize_bill_ref(trip[16]))
        if pd.notna(trip[17]):
            d = trip[17]
            if hasattr(d, "to_pydatetime"):
                ws.cell(row=r, column=18, value=d.to_pydatetime())
            ws.cell(row=r, column=18).number_format = DATE_FMT
        if pd.notna(trip[18]):
            ws.cell(row=r, column=19, value=float(trip[18]))
        if pd.notna(trip[19]):
            ws.cell(row=r, column=20, value=str(trip[19]))

    # Formulas for empty template rows
    for r in range(first_data_row + len(trips), last_data_row + 1):
        for col, formula in trip_formulas(r).items():
            ws.cell(row=r, column=col, value=formula)

    # Money formatting
    for col in [8, 9, 10, 11, 12, 13, 14, 15, 19]:
        for r in range(first_data_row, last_data_row + 1):
            ws.cell(row=r, column=col).number_format = MONEY_FMT

    # Data validations
    add_list_validation(ws, f"C{first_data_row}:C{last_data_row}", ranges["party_range"])
    add_list_validation(ws, f"D{first_data_row}:D{last_data_row}", ranges["dest_range"])
    add_list_validation(ws, f"E{first_data_row}:E{last_data_row}", ranges["truck_range"])
    add_list_validation(ws, f"P{first_data_row}:P{last_data_row}", ranges["pahunch_range"])
    add_list_validation(ws, f"Q{first_data_row}:Q{last_data_row}", ranges["bill_ref_range"])

    # Conditional formatting (must be after all cell writes)
    add_trip_highlight_rules(ws, first_data_row, last_data_row)

    # Instructions
    ws["A2"] = (
        "Yellow = PAHUNCH missing. Red = Bill Ref missing. "
        "Pick Party, Destination & Truck from dropdown arrows. Requires Microsoft Excel."
    )
    ws["A2"].font = Font(italic=True, color="666666", size=10)
    ws.merge_cells("A2:T2")

    col_widths = [8, 12, 30, 18, 14, 14, 10, 8, 10, 8, 8, 12, 11, 11, 10, 10, 12, 12, 12, 24]
    for i, w in enumerate(col_widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.freeze_panes = "A4"
    return ws, first_data_row, last_data_row


def build_freight_bills(wb, data, ranges, trip_first_row, trip_last_row) -> None:
    ws = wb.create_sheet("Freight Bills")
    style_title(ws, "A1", "HARIOM LOGISTICS — FREIGHT BILLS (Auto-calculated)")
    ws.merge_cells("A1:L1")
    ws["A2"] = "Link each bill to a Trip Bill Ref to auto-sum freight from Trip Log. Compare BILL AMOUNT vs Original Amt. while setting up links."
    ws["A2"].font = Font(italic=True, color="666666", size=10)
    ws.merge_cells("A2:L2")

    headers = [
        "DATE",
        "BILL No.",
        "Trip Bill Ref",
        "BILL AMOUNT",
        "Original Amt.",
        "Trip Count",
        "EPOD STATUS",
        "PENDING Amt.",
        "PASS Amt.",
        "Balance Due",
        "Running Total",
        "BILL STATUS",
    ]
    hr = 4
    for col, h in enumerate(headers, start=1):
        ws.cell(row=hr, column=col, value=h)
    style_header_row(ws, hr, len(headers))

    bills = data["bills"]
    ref_map = bill_ref_mapping()
    first_row = hr + 1
    for i, bill in enumerate(bills.itertuples(index=False)):
        r = first_row + i
        if pd.notna(bill[0]):
            d = bill[0]
            ws.cell(row=r, column=1, value=d.to_pydatetime() if hasattr(d, "to_pydatetime") else d)
            ws.cell(row=r, column=1).number_format = DATE_FMT
        ws.cell(row=r, column=2, value=bill[1])
        bill_no = str(bill[1])
        if bill_no in ref_map:
            ws.cell(row=r, column=3, value=ref_map[bill_no])
        ws.cell(row=r, column=4, value=f'=IF(C{r}="","",SUMIF(\'Trip Log\'!$Q${TRIP_DATA_START}:$Q${TRIP_DATA_END},C{r},\'Trip Log\'!$I${TRIP_DATA_START}:$I${TRIP_DATA_END}))')
        ws.cell(row=r, column=5, value=float(bill[2]) if pd.notna(bill[2]) else None)
        ws.cell(row=r, column=6, value=f'=IF(C{r}="","",COUNTIF(\'Trip Log\'!$Q${TRIP_DATA_START}:$Q${TRIP_DATA_END},C{r}))')
        ws.cell(row=r, column=7, value=bill[3] if pd.notna(bill[3]) else "")
        ws.cell(row=r, column=8, value=f'=IF(D{r}="","",D{r}-IF(I{r}="",0,I{r}))')
        ws.cell(row=r, column=9, value=bill[5] if len(bill) > 5 and pd.notna(bill[5]) else None)
        ws.cell(row=r, column=10, value=f'=IF(H{r}="","",H{r})')
        if i == 0:
            ws.cell(row=r, column=11, value=f"=IF(J{r}=\"\",0,J{r})")
        else:
            ws.cell(row=r, column=11, value=f"=IF(J{r}=\"\",\"\",K{r-1}+J{r})")
        ws.cell(row=r, column=12, value=bill[7] if len(bill) > 7 and pd.notna(bill[7]) else "PENDING")

    # Template rows for new bills
    for r in range(first_row + len(bills), first_row + 30):
        ws.cell(row=r, column=4, value=f'=IF(C{r}="","",SUMIF(\'Trip Log\'!$Q${TRIP_DATA_START}:$Q${TRIP_DATA_END},C{r},\'Trip Log\'!$I${TRIP_DATA_START}:$I${TRIP_DATA_END}))')
        ws.cell(row=r, column=6, value=f'=IF(C{r}="","",COUNTIF(\'Trip Log\'!$Q${TRIP_DATA_START}:$Q${TRIP_DATA_END},C{r}))')
        ws.cell(row=r, column=8, value=f'=IF(D{r}="","",D{r}-IF(I{r}="",0,I{r}))')
        ws.cell(row=r, column=10, value=f'=IF(H{r}="","",H{r})')
        if r == first_row:
            ws.cell(row=r, column=11, value=f"=IF(J{r}=\"\",0,J{r})")
        else:
            ws.cell(row=r, column=11, value=f"=IF(J{r}=\"\",\"\",K{r-1}+J{r})")

    last_row = first_row + 29
    add_list_validation(ws, f"C{first_row}:C{last_row}", ranges["bill_ref_range"])
    add_list_validation(ws, f"G{first_row}:G{last_row}", ranges["epod_range"])
    add_list_validation(ws, f"L{first_row}:L{last_row}", ranges["bill_status_range"])

    for col in [3, 4, 5, 8, 9, 10, 11]:
        for r in range(first_row, last_row + 1):
            ws.cell(row=r, column=col).number_format = MONEY_FMT

    add_freight_bill_highlight_rules(ws, first_row, last_row)

    # Summary
    summary_row = last_row + 3
    ws.cell(row=summary_row, column=1, value="TOTAL OUTSTANDING").font = SUBTITLE_FONT
    ws.cell(row=summary_row, column=10, value=f"=SUM(J{first_row}:J{last_row})")
    ws.cell(row=summary_row, column=10).number_format = MONEY_FMT
    ws.cell(row=summary_row, column=10).font = Font(bold=True)

    ws.cell(row=summary_row + 2, column=1, value="Original Manual Bill Total (for reference)").font = Font(italic=True)
    orig_total = float(data["bills"]["BILL AMOUNT"].sum())
    ws.cell(row=summary_row + 2, column=5, value=orig_total)
    ws.cell(row=summary_row + 2, column=5).number_format = MONEY_FMT

    widths = [12, 14, 14, 14, 14, 10, 12, 14, 12, 14, 14, 12]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A5"


def build_owner_ledger(wb, ranges, trip_first_row, trip_last_row, pay_first_row, pay_last_row) -> None:
    ws = wb.create_sheet("Owner Ledger")
    style_title(ws, "A1", "HARIOM LOGISTICS — OWNER LEDGER")
    ws.merge_cells("A1:J1")
    ws["A2"] = "Auto-calculated from Trip Log and Payments. Net Due = Balance Payable − Payments Made."
    ws["A2"].font = Font(italic=True, color="666666", size=10)
    ws.merge_cells("A2:J2")

    headers = [
        "Owner Name",
        "Total Trips",
        "Total Freight",
        "Total Commission",
        "Total Munsiyana",
        "Balance Payable",
        "Payments Made",
        "Net Due",
        "Trips w/o PAHUNCH",
        "Unbilled Trips",
    ]
    hr = 4
    for col, h in enumerate(headers, start=1):
        ws.cell(row=hr, column=col, value=h)
    style_header_row(ws, hr, len(headers))

    owner_start = ranges["owner_start"]
    owner_end = ranges["owner_end"]
    first_row = hr + 1
    for i, r in enumerate(range(first_row, first_row + (owner_end - owner_start + 1))):
        owner_row = owner_start + i
        ws.cell(row=r, column=1, value=f"='Master Data'!J{owner_row}")
        ocell = f"A{r}"
        ws.cell(row=r, column=2, value=f'=IF({ocell}="","",COUNTIF(\'Trip Log\'!$F${TRIP_DATA_START}:$F${TRIP_DATA_END},{ocell}))')
        ws.cell(row=r, column=3, value=f'=IF({ocell}="","",SUMIF(\'Trip Log\'!$F${TRIP_DATA_START}:$F${TRIP_DATA_END},{ocell},\'Trip Log\'!$I${TRIP_DATA_START}:$I${TRIP_DATA_END}))')
        ws.cell(row=r, column=4, value=f'=IF({ocell}="","",SUMIF(\'Trip Log\'!$F${TRIP_DATA_START}:$F${TRIP_DATA_END},{ocell},\'Trip Log\'!$M${TRIP_DATA_START}:$M${TRIP_DATA_END}))')
        ws.cell(row=r, column=5, value=f'=IF({ocell}="","",SUMIF(\'Trip Log\'!$F${TRIP_DATA_START}:$F${TRIP_DATA_END},{ocell},\'Trip Log\'!$N${TRIP_DATA_START}:$N${TRIP_DATA_END}))')
        ws.cell(row=r, column=6, value=f'=IF({ocell}="","",SUMIF(\'Trip Log\'!$F${TRIP_DATA_START}:$F${TRIP_DATA_END},{ocell},\'Trip Log\'!$O${TRIP_DATA_START}:$O${TRIP_DATA_END}))')
        ws.cell(row=r, column=7, value=f'=IF({ocell}="","",SUMIF(Payments!$E${PAY_DATA_START}:$E${PAY_DATA_END},{ocell},Payments!$C${PAY_DATA_START}:$C${PAY_DATA_END}))')
        ws.cell(row=r, column=8, value=f'=IF({ocell}="","",F{r}-G{r})')
        ws.cell(row=r, column=9, value=f'=IF({ocell}="","",COUNTIFS(\'Trip Log\'!$F${TRIP_DATA_START}:$F${TRIP_DATA_END},{ocell},\'Trip Log\'!$P${TRIP_DATA_START}:$P${TRIP_DATA_END},""))')
        ws.cell(row=r, column=10, value=f'=IF({ocell}="","",COUNTIFS(\'Trip Log\'!$F${TRIP_DATA_START}:$F${TRIP_DATA_END},{ocell},\'Trip Log\'!$Q${TRIP_DATA_START}:$Q${TRIP_DATA_END},""))')

    last_owner_row = first_row + (owner_end - owner_start)
    totals_row = last_owner_row + 2
    ws.cell(row=totals_row, column=1, value="GRAND TOTAL").font = SUBTITLE_FONT
    for col in range(2, 11):
        letter = get_column_letter(col)
        ws.cell(row=totals_row, column=col, value=f"=SUM({letter}{first_row}:{letter}{last_owner_row})")
        ws.cell(row=totals_row, column=col).font = Font(bold=True)

    for col in [3, 4, 5, 6, 7, 8]:
        for r in range(first_row, totals_row + 1):
            ws.cell(row=r, column=col).number_format = MONEY_FMT

    add_owner_ledger_highlight_rules(ws, first_row, last_owner_row)

    widths = [16, 10, 14, 14, 14, 14, 14, 14, 14, 12]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A5"


def build_payments(wb, data, ranges) -> tuple:
    ws = wb.create_sheet("Payments")
    style_title(ws, "A1", "HARIOM LOGISTICS — PAYMENTS")
    ws.merge_cells("A1:F1")
    ws["A2"] = "Record payments to truck owners. Select Owner from dropdown to auto-update Owner Ledger."
    ws["A2"].font = Font(italic=True, color="666666", size=10)
    ws.merge_cells("A2:F2")

    headers = ["SL. NO.", "Date", "Paid Amount", "Paid To (Name)", "Owner", "Particulars"]
    hr = 3
    for col, h in enumerate(headers, start=1):
        ws.cell(row=hr, column=col, value=h)
    style_header_row(ws, hr, len(headers))

    first_row = hr + 1
    max_row = first_row + 199
    payments = data["payments"]

    # Map known payment parties to owners where possible
    owner_hints = {
        "ABDUL MALIK JI": "VINEET",  # from particulars mentioning VINEET
    }

    for i, pay in enumerate(payments.itertuples(index=False)):
        r = first_row + i
        if pd.notna(pay[0]):
            ws.cell(row=r, column=1, value=int(pay[0]))
        if pd.notna(pay[1]):
            d = pay[1]
            ws.cell(row=r, column=2, value=d.to_pydatetime() if hasattr(d, "to_pydatetime") else d)
            ws.cell(row=r, column=2).number_format = DATE_FMT
        ws.cell(row=r, column=3, value=float(pay[2]) if pd.notna(pay[2]) else None)
        ws.cell(row=r, column=4, value=pay[3] if pd.notna(pay[3]) else "")
        party = str(pay[3]) if pd.notna(pay[3]) else ""
        if party in owner_hints:
            ws.cell(row=r, column=5, value=owner_hints[party])
        elif party in data["owners"]:
            ws.cell(row=r, column=5, value=party)
        ws.cell(row=r, column=6, value=pay[4] if pd.notna(pay[4]) else "")

    for col in [3]:
        for r in range(first_row, max_row + 1):
            ws.cell(row=r, column=col).number_format = MONEY_FMT

    add_list_validation(ws, f"E{first_row}:E{max_row}", ranges["owner_range"])

    widths = [8, 12, 14, 22, 14, 36]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A4"
    return first_row, max_row


def build_pump_sheet(wb, ranges) -> None:
    ws = wb.create_sheet("Pump Account")
    style_title(ws, "B1", "HAMARA PUMP, KATRA — DIESEL ACCOUNT")
    ws.merge_cells("B1:H1")
    ws["B2"] = "Track diesel deposits and usage per vehicle. Balance = Previous Balance + Deposit − Diesel − Cash."
    ws["B2"].font = Font(italic=True, color="666666", size=10)
    ws.merge_cells("B2:H2")

    headers = ["DATE", "Vehicle No.", "Deposit Amt.", "Diesel", "Cash", "Balance", "Notes"]
    hr = 4
    for col, h in enumerate(headers, start=2):
        ws.cell(row=hr, column=col, value=h)
    style_header_row(ws, hr, 8)

    first_row = hr + 1
    last_row = first_row + 99
    for r in range(first_row, last_row + 1):
        if r == first_row:
            ws.cell(row=r, column=7, value=f"=IF(D{r}=\"\",\"\",D{r}-E{r}-F{r})")
        else:
            ws.cell(row=r, column=7, value=f"=IF(B{r}=\"\",\"\",IF(G{r-1}=\"\",0,G{r-1})+D{r}-E{r}-F{r})")

    add_list_validation(ws, f"C{first_row}:C{last_row}", ranges["truck_range"])

    for col in [4, 5, 6, 7]:
        for r in range(first_row, last_row + 1):
            ws.cell(row=r, column=col).number_format = MONEY_FMT

    widths = {2: 12, 3: 14, 4: 12, 5: 10, 6: 10, 7: 12, 8: 20}
    for col, w in widths.items():
        ws.column_dimensions[get_column_letter(col)].width = w
    ws.freeze_panes = "B5"


def build_dashboard(wb, ranges) -> None:
    ws = wb.create_sheet("Dashboard")
    ws.sheet_properties.tabColor = "1F4E79"
    style_title(ws, "A1", "HARIOM LOGISTICS — DASHBOARD")
    ws.merge_cells("A1:D1")

    owner_last = ranges["owner_end"] - ranges["owner_start"] + 5
    cards = [
        ("A3", "Total Trips", f"=COUNTA('Trip Log'!A{TRIP_DATA_START}:A{TRIP_DATA_END})"),
        ("A5", "Total Freight (All Trips)", f"=SUM('Trip Log'!I{TRIP_DATA_START}:I{TRIP_DATA_END})"),
        ("A7", "Outstanding Owner Balance", f"=SUM('Owner Ledger'!H5:H{owner_last})"),
        ("A9", "Pending Freight Bills", "=SUM('Freight Bills'!J5:J34)"),
        ("A11", "Trips Missing PAHUNCH", f"=COUNTIFS('Trip Log'!P{TRIP_DATA_START}:P{TRIP_DATA_END},\"\",'Trip Log'!I{TRIP_DATA_START}:I{TRIP_DATA_END},\">0\")"),
        ("A13", "Trips Missing Bill Ref", f"=COUNTIFS('Trip Log'!Q{TRIP_DATA_START}:Q{TRIP_DATA_END},\"\",'Trip Log'!I{TRIP_DATA_START}:I{TRIP_DATA_END},\">0\")"),
    ]
    for cell, label, formula in cards:
        row = int(cell[1:])
        ws[cell] = label
        ws[cell].font = SUBTITLE_FONT
        ws[f"B{row}"] = formula
        if "Freight" in label or "Balance" in label or "Pending" in label:
            ws[f"B{row}"].number_format = MONEY_FMT
        ws[f"B{row}"].font = Font(bold=True, size=14, color="1F4E79")

    ws["A16"] = "Last updated"
    ws["B16"] = datetime.now()
    ws["B16"].number_format = "DD-MMM-YYYY HH:MM"

    ws["A19"] = "HOW TO USE"
    ws["A19"].font = SUBTITLE_FONT
    tips = [
        "1. Trip Log: pick Party, Destination & Truck from dropdowns — Owner, Commission & Balance auto-fill.",
        "2. Freight Bills: set Trip Bill Ref (e.g. 11, 4 EPOD) — BILL AMOUNT auto-sums matching trips.",
        "3. Payments: always select Owner so Owner Ledger stays accurate.",
        "4. Master Data: add new parties, trucks, or destinations here — dropdowns update automatically.",
        "5. Row colors on Trip Log: Yellow = PAHUNCH missing. Red = Bill Ref missing.",
    ]
    for i, tip in enumerate(tips, start=20):
        ws.cell(row=i, column=1, value=tip)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=4)
        ws.cell(row=i, column=1).font = Font(size=10, color="444444")

    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 18
    ws.column_dimensions["D"].width = 18


def build_statements_sheet(wb) -> None:
    ws = wb.create_sheet("Statements")
    ws.sheet_properties.tabColor = "548235"
    style_title(ws, "A1", "HARIOM LOGISTICS — OWNER STATEMENT PDFs")
    ws.merge_cells("A1:D1")
    ws["A2"] = "Generate monthly PDF statements per truck owner from Trip Log and Payments data."
    ws["A2"].font = Font(italic=True, color="666666", size=10)
    ws.merge_cells("A2:D2")

    ws["A4"] = "COMMAND"
    ws["A4"].font = SUBTITLE_FONT
    ws["B4"] = "python3 generate_owner_statements.py --month 2026-06"
    ws["B4"].font = Font(name="Courier New", size=10)

    commands = [
        ("List available months", "python3 generate_owner_statements.py --list-months"),
        ("Generate all owners (June 2026)", "python3 generate_owner_statements.py --month 2026-06"),
        ("Generate one owner only", "python3 generate_owner_statements.py --month 2026-06 --owner VINEET"),
        ("Custom output folder", "python3 generate_owner_statements.py --month 2026-06 --output ./statements"),
    ]
    row = 6
    ws.cell(row=5, column=1, value="Option").font = SUBTITLE_FONT
    ws.cell(row=5, column=2, value="Command").font = SUBTITLE_FONT
    for label, cmd in commands:
        ws.cell(row=row, column=1, value=label)
        ws.cell(row=row, column=2, value=cmd)
        ws.cell(row=row, column=2).font = Font(name="Courier New", size=9)
        row += 1

    ws.cell(row=row + 1, column=1, value="OUTPUT LOCATION").font = SUBTITLE_FONT
    ws.cell(row=row + 2, column=1, value="PDFs are saved to: statements/YYYY-MM/YYYY-MM_OwnerName_statement.pdf")
    ws.merge_cells(start_row=row + 2, start_column=1, end_row=row + 2, end_column=4)

    ws.cell(row=row + 4, column=1, value="EACH PDF INCLUDES").font = SUBTITLE_FONT
    bullets = [
        "Trip details: LR, date, party, destination, truck, freight, commission, balance, PAHUNCH",
        "Trip summary: totals for freight, commission, munsiyana, balance payable",
        "Payments received that month",
        "Account summary: balance payable − payments = net due",
    ]
    for i, text in enumerate(bullets, start=row + 5):
        ws.cell(row=i, column=1, value=f"• {text}")
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=4)

    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 72
    ws.column_dimensions["C"].width = 18
    ws.column_dimensions["D"].width = 18


def set_sheet_order(wb) -> None:
    order = [
        "Dashboard",
        "Trip Log",
        "Owner Ledger",
        "Freight Bills",
        "Payments",
        "Statements",
        "Pump Account",
        "Master Data",
    ]
    for i, name in enumerate(order):
        wb.move_sheet(name, offset=i - wb.sheetnames.index(name))


def main() -> None:
    data = load_source_data()
    wb = openpyxl.Workbook()
    _, ranges = build_master_data(wb, data)
    _, trip_first, trip_last = build_trip_log(wb, data, ranges)
    pay_first, pay_last = build_payments(wb, data, ranges)
    build_freight_bills(wb, data, ranges, trip_first, trip_last)
    build_owner_ledger(wb, ranges, trip_first, trip_last, pay_first, pay_last)
    build_pump_sheet(wb, ranges)
    build_dashboard(wb, ranges)
    build_statements_sheet(wb)
    create_defined_names(wb)
    set_sheet_order(wb)

    wb.calculation.fullCalcOnLoad = True
    wb.save(OUTPUT)
    print(f"Saved: {OUTPUT}")
    print(f"Sheets: {wb.sheetnames}")
    print(f"Trips migrated: {len(data['trips'])}")


if __name__ == "__main__":
    main()
