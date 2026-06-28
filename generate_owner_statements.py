#!/usr/bin/env python3
"""Generate monthly owner statement PDFs from HARIOM LOGISTICS workbook."""

from __future__ import annotations

import argparse
import calendar
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path

import pandas as pd
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_RIGHT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import (
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

DEFAULT_WORKBOOK = Path("/workspace/HARIOM LOGISTICS - Automated.xlsx")
DEFAULT_OUTPUT = Path("/workspace/statements")

BRAND_COLOR = colors.HexColor("#1F4E79")
ACCENT_COLOR = colors.HexColor("#D9E2F3")
MUTED_COLOR = colors.HexColor("#666666")


@dataclass
class TruckInfo:
    owner: str
    commission: float
    munsiyana: float


def parse_month(value: str) -> tuple[int, int]:
    try:
        year, month = value.split("-")
        month_i = int(month)
        year_i = int(year)
        if not 1 <= month_i <= 12:
            raise ValueError
        return year_i, month_i
    except ValueError as exc:
        raise argparse.ArgumentTypeError(
            f"Invalid month {value!r}. Use YYYY-MM (e.g. 2026-06)."
        ) from exc


def money(amount: float | int | None) -> str:
    if amount is None or pd.isna(amount):
        return "₹0"
    return f"₹{float(amount):,.0f}"


def load_truck_master(workbook: Path) -> dict[str, TruckInfo]:
    raw = pd.read_excel(workbook, sheet_name="Master Data", header=None)
    trucks: dict[str, TruckInfo] = {}
    for _, row in raw.iloc[4:].iterrows():
        truck_no = row.iloc[4]
        owner = row.iloc[5]
        if pd.isna(truck_no) or pd.isna(owner):
            continue
        trucks[str(truck_no).strip()] = TruckInfo(
            owner=str(owner).strip(),
            commission=float(row.iloc[6]) if pd.notna(row.iloc[6]) else 0.0,
            munsiyana=float(row.iloc[7]) if pd.notna(row.iloc[7]) else 0.0,
        )
    return trucks


def enrich_trips(trips: pd.DataFrame, trucks: dict[str, TruckInfo]) -> pd.DataFrame:
    df = trips.copy()
    df = df[df["LR No"].notna()].copy()
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")

    owners: list[str] = []
    commissions: list[float] = []
    munsiyanans: list[float] = []
    freights: list[float] = []
    advances: list[float] = []
    balances: list[float] = []

    for row in df.itertuples(index=False):
        truck = "" if pd.isna(row[4]) else str(row[4]).strip()
        info = trucks.get(truck)
        owner = row[5] if pd.notna(row[5]) else (info.owner if info else "")
        owners.append(owner)

        qty = float(row[6]) if pd.notna(row[6]) else 0.0
        rate = float(row[7]) if pd.notna(row[7]) else 0.0
        diesel = float(row[9]) if pd.notna(row[9]) else 0.0
        cash = float(row[10]) if pd.notna(row[10]) else 0.0

        freight = float(row[8]) if pd.notna(row[8]) else qty * rate
        commission = float(row[12]) if pd.notna(row[12]) else (info.commission if info else 0.0)
        munsiyana = float(row[13]) if pd.notna(row[13]) else (info.munsiyana if info else 0.0)
        advance = freight - diesel - cash
        balance = float(row[14]) if pd.notna(row[14]) else advance - commission - munsiyana

        freights.append(freight)
        commissions.append(commission)
        munsiyanans.append(munsiyana)
        advances.append(advance)
        balances.append(balance)

    df["Owner Name"] = owners
    df["Freight"] = freights
    df["Commission"] = commissions
    df["Munsiyana"] = munsiyanans
    df["Total Advance"] = advances
    df["Balance"] = balances
    return df


def load_workbook_data(workbook: Path) -> tuple[pd.DataFrame, pd.DataFrame]:
    trucks = load_truck_master(workbook)
    trips = pd.read_excel(workbook, sheet_name="Trip Log", header=2)
    trips = enrich_trips(trips, trucks)

    payments = pd.read_excel(workbook, sheet_name="Payments", header=2)
    payments = payments.dropna(how="all")
    payments["Date"] = pd.to_datetime(payments["Date"], errors="coerce")
    return trips, payments


def month_label(year: int, month: int) -> str:
    return f"{calendar.month_name[month]} {year}"


def filter_month(df: pd.DataFrame, year: int, month: int, date_col: str) -> pd.DataFrame:
    if df.empty:
        return df
    mask = (df[date_col].dt.year == year) & (df[date_col].dt.month == month)
    return df.loc[mask].copy()


def statement_filename(owner: str, year: int, month: int) -> str:
    safe_owner = "".join(ch if ch.isalnum() else "_" for ch in owner).strip("_")
    return f"{year}-{month:02d}_{safe_owner}_statement.pdf"


def build_styles():
    styles = getSampleStyleSheet()
    styles.add(
        ParagraphStyle(
            name="BrandTitle",
            parent=styles["Title"],
            fontSize=18,
            textColor=BRAND_COLOR,
            spaceAfter=4,
        )
    )
    styles.add(
        ParagraphStyle(
            name="Subtitle",
            parent=styles["Normal"],
            fontSize=11,
            textColor=MUTED_COLOR,
            spaceAfter=10,
        )
    )
    styles.add(
        ParagraphStyle(
            name="Section",
            parent=styles["Heading2"],
            fontSize=12,
            textColor=BRAND_COLOR,
            spaceBefore=8,
            spaceAfter=6,
        )
    )
    styles.add(
        ParagraphStyle(
            name="SmallMuted",
            parent=styles["Normal"],
            fontSize=8,
            textColor=MUTED_COLOR,
        )
    )
    return styles


def make_table(data: list[list], col_widths: list[float] | None = None) -> Table:
    table = Table(data, colWidths=col_widths, repeatRows=1)
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), BRAND_COLOR),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, -1), 8),
                ("ALIGN", (0, 0), (-1, 0), "CENTER"),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, ACCENT_COLOR]),
                ("TOPPADDING", (0, 0), (-1, -1), 4),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ]
        )
    )
    return table


def build_owner_pdf(
    owner: str,
    year: int,
    month: int,
    trips: pd.DataFrame,
    payments: pd.DataFrame,
    output_path: Path,
) -> None:
    styles = build_styles()
    output_path.parent.mkdir(parents=True, exist_ok=True)

    doc = SimpleDocTemplate(
        str(output_path),
        pagesize=A4,
        leftMargin=14 * mm,
        rightMargin=14 * mm,
        topMargin=14 * mm,
        bottomMargin=14 * mm,
        title=f"{owner} Statement {year}-{month:02d}",
    )

    story: list = []
    period = month_label(year, month)

    story.append(Paragraph("HARIOM LOGISTICS", styles["BrandTitle"]))
    story.append(Paragraph("Monthly Owner Statement", styles["Subtitle"]))
    story.append(Paragraph(f"<b>Owner:</b> {owner}", styles["Normal"]))
    story.append(Paragraph(f"<b>Period:</b> {period}", styles["Normal"]))
    story.append(Spacer(1, 8))

    owner_trips = trips[trips["Owner Name"] == owner].sort_values("Date")
    owner_payments = payments[payments["Owner"] == owner].sort_values("Date")

    # Trip details
    story.append(Paragraph("Trip Details", styles["Section"]))
    if owner_trips.empty:
        story.append(Paragraph("No trips recorded for this period.", styles["Normal"]))
    else:
        trip_rows = [
            [
                "LR",
                "Date",
                "Party",
                "Destination",
                "Truck",
                "Qty",
                "Freight",
                "Comm.",
                "Balance",
                "PAHUNCH",
            ]
        ]
        for row in owner_trips.itertuples(index=False):
            trip_rows.append(
                [
                    str(int(row[0])) if pd.notna(row[0]) and float(row[0]).is_integer() else str(row[0]),
                    row[1].strftime("%d-%b-%Y") if pd.notna(row[1]) else "",
                    str(row[2])[:22],
                    str(row[3])[:14],
                    str(row[4]),
                    f"{float(row[6]):g}" if pd.notna(row[6]) else "",
                    money(row[8]),
                    money(row[12]),
                    money(row[14]),
                    str(row[15]) if pd.notna(row[15]) else "",
                ]
            )
        story.append(
            make_table(
                trip_rows,
                col_widths=[24, 48, 82, 58, 52, 24, 44, 38, 44, 36],
            )
        )

    story.append(Spacer(1, 10))
    story.append(Paragraph("Trip Summary", styles["Section"]))

    total_freight = owner_trips["Freight"].sum()
    total_commission = owner_trips["Commission"].sum()
    total_munsiyana = owner_trips["Munsiyana"].sum()
    total_balance = owner_trips["Balance"].sum()
    pending_pahunch = owner_trips["PAHUNCH"].isna().sum() + (owner_trips["PAHUNCH"] == "").sum()

    summary_data = [
        ["Metric", "Value"],
        ["Total Trips", str(len(owner_trips))],
        ["Total Freight", money(total_freight)],
        ["Total Commission", money(total_commission)],
        ["Total Munsiyana", money(total_munsiyana)],
        ["Balance Payable (Trips)", money(total_balance)],
        ["Trips Pending PAHUNCH", str(int(pending_pahunch))],
    ]
    summary_table = Table(summary_data, colWidths=[180, 120])
    summary_table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), BRAND_COLOR),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, -1), 9),
                ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, ACCENT_COLOR]),
                ("ALIGN", (1, 1), (1, -1), "RIGHT"),
                ("TOPPADDING", (0, 0), (-1, -1), 5),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
            ]
        )
    )
    story.append(summary_table)

    # Payments
    story.append(Spacer(1, 10))
    story.append(Paragraph("Payments Received This Month", styles["Section"]))
    if owner_payments.empty:
        story.append(Paragraph("No payments recorded for this period.", styles["Normal"]))
        total_paid = 0.0
    else:
        pay_rows = [["Date", "Paid To", "Amount", "Particulars"]]
        for row in owner_payments.itertuples(index=False):
            pay_rows.append(
                [
                    row[1].strftime("%d-%b-%Y") if pd.notna(row[1]) else "",
                    str(row[3])[:24],
                    money(row[2]),
                    str(row[5])[:42] if pd.notna(row[5]) else "",
                ]
            )
        story.append(make_table(pay_rows, col_widths=[52, 90, 52, 210]))
        total_paid = float(owner_payments["Paid Amount"].fillna(0).sum())

    net_due = total_balance - total_paid

    story.append(Spacer(1, 10))
    story.append(Paragraph("Account Summary", styles["Section"]))
    account_data = [
        ["Description", "Amount"],
        ["Balance Payable (Trips)", money(total_balance)],
        ["Payments Received", money(total_paid)],
        ["Net Due", money(net_due)],
    ]
    account_table = Table(account_data, colWidths=[180, 120])
    account_table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), BRAND_COLOR),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, -1), 10),
                ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
                ("BACKGROUND", (0, -1), (-1, -1), ACCENT_COLOR),
                ("ALIGN", (1, 1), (1, -1), "RIGHT"),
                ("TOPPADDING", (0, 0), (-1, -1), 6),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
            ]
        )
    )
    story.append(account_table)

    story.append(Spacer(1, 16))
    story.append(
        Paragraph(
            f"Generated on {datetime.now().strftime('%d-%b-%Y %H:%M')} · HARIOM LOGISTICS",
            styles["SmallMuted"],
        )
    )
    story.append(
        Paragraph(
            "This is a computer-generated statement. Please verify trip details and PAHUNCH status.",
            styles["SmallMuted"],
        )
    )

    doc.build(story)


def available_months(trips: pd.DataFrame, payments: pd.DataFrame) -> list[tuple[int, int]]:
    months: set[tuple[int, int]] = set()
    for df, col in ((trips, "Date"), (payments, "Date")):
        if df.empty:
            continue
        valid = df[df[col].notna()]
        for dt in valid[col]:
            months.add((dt.year, dt.month))
    return sorted(months)


def generate_statements(
    workbook: Path,
    output_dir: Path,
    year: int,
    month: int,
    owners: list[str] | None = None,
) -> list[Path]:
    trips, payments = load_workbook_data(workbook)
    month_trips = filter_month(trips, year, month, "Date")
    month_payments = filter_month(payments, year, month, "Date")

    if owners:
        target_owners = owners
    else:
        target_owners = sorted(
            set(month_trips["Owner Name"].dropna().astype(str))
            | set(month_payments["Owner"].dropna().astype(str))
        )
        target_owners = [o for o in target_owners if o and o != "nan"]

    if not target_owners:
        raise SystemExit(f"No owner activity found for {year}-{month:02d}.")

    period_dir = output_dir / f"{year}-{month:02d}"
    created: list[Path] = []
    for owner in target_owners:
        out = period_dir / statement_filename(owner, year, month)
        build_owner_pdf(owner, year, month, month_trips, month_payments, out)
        created.append(out)
    return created


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Generate monthly owner statement PDFs from the HARIOM LOGISTICS workbook."
    )
    parser.add_argument(
        "--input",
        type=Path,
        default=DEFAULT_WORKBOOK,
        help="Path to HARIOM LOGISTICS - Automated.xlsx",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=DEFAULT_OUTPUT,
        help="Output directory for PDF statements",
    )
    parser.add_argument(
        "--month",
        type=parse_month,
        help="Statement month as YYYY-MM (default: latest month with data)",
    )
    parser.add_argument(
        "--owner",
        action="append",
        help="Generate for specific owner only (can be repeated)",
    )
    parser.add_argument(
        "--list-months",
        action="store_true",
        help="List months that have trip or payment data",
    )
    args = parser.parse_args()

    if not args.input.exists():
        raise SystemExit(f"Workbook not found: {args.input}")

    trips, payments = load_workbook_data(args.input)

    if args.list_months:
        months = available_months(trips, payments)
        if not months:
            print("No dated trips or payments found.")
            return
        for year, month in months:
            print(f"{year}-{month:02d}  ({month_label(year, month)})")
        return

    if args.month:
        year, month = args.month
    else:
        months = available_months(trips, payments)
        if not months:
            raise SystemExit("No dated trips or payments found. Use --month YYYY-MM.")
        year, month = months[-1]

    created = generate_statements(args.input, args.output, year, month, args.owner)
    print(f"Generated {len(created)} statement(s) for {month_label(year, month)}:")
    for path in created:
        print(f"  {path}")


if __name__ == "__main__":
    main()
