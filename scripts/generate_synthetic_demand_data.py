"""Generate a deterministic synthetic demand-history workbook for Demand Genie."""

from __future__ import annotations

import math
import random
from datetime import date
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter


OUTPUT_PATH = Path("data/Demand_Genie_Synthetic_Demand_History.xlsx")
SEED = 20260721
MONTHS = [date(year, month, 1) for year in range(2023, 2027) for month in range(1, 13)][6:42]
SKU_MULTIPLIERS = (0.55, 0.68, 0.8, 0.95, 1.1, 1.3, 1.55, 1.85)

GROUPS = (
    {
        "code": "BRG",
        "name": "Industrial Bearings",
        "base": 760,
        "seasonality": 0.13,
        "trend": 0.004,
        "volatility": 0.06,
        "profile": "Seasonal",
        "price": 78,
        "lead_time": 42,
    },
    {
        "code": "GBX",
        "name": "Gearbox Components",
        "base": 440,
        "seasonality": 0.09,
        "trend": 0.006,
        "volatility": 0.08,
        "profile": "Trending",
        "price": 215,
        "lead_time": 63,
    },
    {
        "code": "HYP",
        "name": "Hydraulic Pumps",
        "base": 285,
        "seasonality": 0.16,
        "trend": 0.003,
        "volatility": 0.12,
        "profile": "Volatile",
        "price": 540,
        "lead_time": 70,
    },
    {
        "code": "SEA",
        "name": "Sealing Systems",
        "base": 1220,
        "seasonality": 0.07,
        "trend": 0.002,
        "volatility": 0.05,
        "profile": "Stable",
        "price": 29,
        "lead_time": 28,
    },
    {
        "code": "CVL",
        "name": "Control Valves",
        "base": 330,
        "seasonality": 0.11,
        "trend": -0.002,
        "volatility": 0.1,
        "profile": "Volatile",
        "price": 310,
        "lead_time": 56,
    },
    {
        "code": "DRV",
        "name": "Electric Drives",
        "base": 245,
        "seasonality": 0.14,
        "trend": 0.007,
        "volatility": 0.09,
        "profile": "Trending",
        "price": 680,
        "lead_time": 77,
    },
    {
        "code": "SNS",
        "name": "Sensor Modules",
        "base": 510,
        "seasonality": 0.05,
        "trend": 0.01,
        "volatility": 0.13,
        "profile": "Growth",
        "price": 168,
        "lead_time": 49,
    },
    {
        "code": "CNV",
        "name": "Conveying Components",
        "base": 620,
        "seasonality": 0.17,
        "trend": 0.001,
        "volatility": 0.08,
        "profile": "Seasonal",
        "price": 115,
        "lead_time": 45,
    },
    {
        "code": "FST",
        "name": "Fastening Kits",
        "base": 1880,
        "seasonality": 0.06,
        "trend": 0.002,
        "volatility": 0.07,
        "profile": "Stable",
        "price": 16,
        "lead_time": 21,
    },
    {
        "code": "PCK",
        "name": "Packaging Equipment",
        "base": 165,
        "seasonality": 0.2,
        "trend": 0.005,
        "volatility": 0.15,
        "profile": "Intermittent",
        "price": 950,
        "lead_time": 84,
    },
)

DESCRIPTORS = (
    "Core",
    "Compact",
    "Standard",
    "Heavy Duty",
    "Precision",
    "High Capacity",
    "Service Kit",
    "Premium",
)

HEADER_FILL = PatternFill("solid", fgColor="1F5E69")
SUBHEADER_FILL = PatternFill("solid", fgColor="DCEBED")
WHITE_FONT = Font(color="FFFFFF", bold=True)
HEADER_FONT = Font(bold=True, color="17262B")
THIN_LINE = Side(style="thin", color="CBD6D7")


def calculate_demand(group: dict[str, float | int | str], sku_index: int, month_index: int, rng: random.Random) -> int:
    """Create a positive monthly demand value with useful forecasting characteristics."""
    multiplier = SKU_MULTIPLIERS[sku_index]
    phase = (sku_index * 0.61) + (len(str(group["code"])) * 0.2)
    seasonality = 1 + float(group["seasonality"]) * math.sin((month_index / 12 * 2 * math.pi) + phase)
    trend = 1 + float(group["trend"]) * month_index
    noise = 1 + rng.gauss(0, float(group["volatility"]))
    shock = 1.0

    if (sku_index + month_index * 3) % 29 == 0:
        shock *= 1.16
    if (sku_index * 5 + month_index) % 37 == 0:
        shock *= 0.82
    if group["profile"] == "Intermittent" and (month_index + sku_index) % 9 == 0:
        shock *= 0.32
    if group["profile"] == "Growth" and month_index >= 24:
        shock *= 1.08

    demand = float(group["base"]) * multiplier * seasonality * trend * noise * shock
    return max(1, int(round(demand)))


def make_sku(group: dict[str, float | int | str], sku_index: int) -> dict[str, str | float | int]:
    sku = f"{group['code']}-{sku_index + 1:02d}"
    return {
        "product_group_code": str(group["code"]),
        "product_group": str(group["name"]),
        "sku": sku,
        "description": f"{group['name']} {DESCRIPTORS[sku_index]}",
        "profile": str(group["profile"]),
        "unit_price": round(float(group["price"]) * (0.78 + sku_index * 0.07), 2),
        "lead_time": int(group["lead_time"]) + (sku_index % 4) * 3,
        "abc_class": "A" if sku_index >= 5 else "B" if sku_index >= 2 else "C",
        "xyz_class": "X" if group["profile"] in {"Stable", "Growth"} else "Y" if group["profile"] in {"Seasonal", "Trending"} else "Z",
    }


def style_sheet(sheet, widths: dict[str, float], freeze: str = "A2") -> None:
    sheet.freeze_panes = freeze
    sheet.sheet_view.showGridLines = False
    for cell in sheet[1]:
        cell.fill = HEADER_FILL
        cell.font = WHITE_FONT
        cell.alignment = Alignment(horizontal="left")
    for column, width in widths.items():
        sheet.column_dimensions[column].width = width


def add_table(sheet, name: str) -> None:
    table = Table(displayName=name, ref=f"A1:{sheet.cell(sheet.max_row, sheet.max_column).coordinate}")
    table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showFirstColumn=False, showLastColumn=False, showRowStripes=True)
    sheet.add_table(table)


def build_workbook() -> None:
    rng = random.Random(SEED)
    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)

    workbook = Workbook()
    read_me = workbook.active
    read_me.title = "Read_Me"
    read_me.append(["Demand Genie Synthetic Demand History"])
    read_me["A1"].fill = HEADER_FILL
    read_me["A1"].font = Font(color="FFFFFF", bold=True, size=15)
    read_me.merge_cells("A1:D1")
    read_me.append([])
    read_me.append(["Purpose", "Synthetic monthly demand data for forecasting and demand-planning dashboard development."])
    read_me.append(["Coverage", "36 complete months: July 2023 through June 2026."])
    read_me.append(["Scope", "10 product groups, 8 SKUs per group, 80 total SKUs, 2,880 monthly demand records."])
    read_me.append(["Upload sheet", "Use Demand_Data for a dashboard upload. It is a tidy, row-based table with one SKU-month per row."])
    read_me.append(["Business view", "Use Demand_Matrix for a wide, spreadsheet-style history view."])
    read_me.append(["Data note", "Values are deterministic synthetic data. They contain trend, seasonality, volatility, shocks, and intermittent behavior; they do not describe real products or customers."])
    read_me.append([])
    read_me.append(["Sheet", "Use"])
    read_me.append(["Demand_Data", "Tidy historical demand records for upload and analysis."])
    read_me.append(["Demand_Matrix", "One row per SKU with one column per month."])
    read_me.append(["Product_Master", "SKU attributes, price, lead time, and ABC/XYZ labels."])
    read_me.append(["Data_Dictionary", "Field definitions for Demand_Data."])
    for row in read_me.iter_rows(min_row=3, max_row=read_me.max_row, min_col=1, max_col=2):
        row[0].font = HEADER_FONT
        row[0].fill = SUBHEADER_FILL
        row[0].border = Border(bottom=THIN_LINE)
        row[1].alignment = Alignment(wrap_text=True, vertical="top")
    read_me.column_dimensions["A"].width = 22
    read_me.column_dimensions["B"].width = 110
    read_me.sheet_view.showGridLines = False

    master = workbook.create_sheet("Product_Master")
    master.append([
        "Product_Group_Code", "Product_Group", "SKU", "SKU_Description", "Demand_Profile",
        "Unit_Price_EUR", "Lead_Time_Days", "ABC_Class", "XYZ_Class",
    ])

    demand = workbook.create_sheet("Demand_Data")
    demand.append([
        "Month", "Product_Group_Code", "Product_Group", "SKU", "SKU_Description", "Demand_Units",
        "Unit_Price_EUR", "Demand_Value_EUR", "Lead_Time_Days", "Demand_Profile",
    ])

    matrix = workbook.create_sheet("Demand_Matrix")
    matrix.append([
        "Product_Group_Code", "Product_Group", "SKU", "SKU_Description", "Demand_Profile",
        *[month.strftime("%Y-%m") for month in MONTHS],
    ])

    all_skus: list[dict[str, str | float | int]] = []
    for group in GROUPS:
        for sku_index in range(8):
            sku = make_sku(group, sku_index)
            all_skus.append(sku)
            master.append([
                sku["product_group_code"], sku["product_group"], sku["sku"], sku["description"], sku["profile"],
                sku["unit_price"], sku["lead_time"], sku["abc_class"], sku["xyz_class"],
            ])

            history: list[int] = []
            for month_index, month in enumerate(MONTHS):
                demand_units = calculate_demand(group, sku_index, month_index, rng)
                history.append(demand_units)
                demand.append([
                    month, sku["product_group_code"], sku["product_group"], sku["sku"], sku["description"], demand_units,
                    sku["unit_price"], round(demand_units * float(sku["unit_price"]), 2), sku["lead_time"], sku["profile"],
                ])

            matrix.append([
                sku["product_group_code"], sku["product_group"], sku["sku"], sku["description"], sku["profile"], *history,
            ])

    dictionary = workbook.create_sheet("Data_Dictionary")
    dictionary.append(["Field", "Description", "Example"])
    dictionary_rows = (
        ("Month", "First day of the monthly demand period. Use this field as the time axis.", "2026-06-01"),
        ("Product_Group_Code", "Stable code for the main product group.", "BRG"),
        ("Product_Group", "Main product group label.", "Industrial Bearings"),
        ("SKU", "Unique stock-keeping unit identifier.", "BRG-01"),
        ("SKU_Description", "Readable SKU description.", "Industrial Bearings Core"),
        ("Demand_Units", "Observed monthly demand quantity in units.", "428"),
        ("Unit_Price_EUR", "Synthetic price per unit in EUR.", "60.84"),
        ("Demand_Value_EUR", "Demand_Units multiplied by Unit_Price_EUR.", "26039.52"),
        ("Lead_Time_Days", "Synthetic replenishment lead time in days.", "42"),
        ("Demand_Profile", "Designed behavior pattern to aid forecasting tests.", "Seasonal"),
    )
    for row in dictionary_rows:
        dictionary.append(row)

    style_sheet(master, {"A": 20, "B": 25, "C": 14, "D": 34, "E": 16, "F": 16, "G": 16, "H": 12, "I": 12})
    style_sheet(demand, {"A": 14, "B": 20, "C": 25, "D": 14, "E": 34, "F": 14, "G": 16, "H": 18, "I": 16, "J": 16})
    matrix_widths = {"A": 20, "B": 25, "C": 14, "D": 34, "E": 16}
    matrix_widths.update({get_column_letter(column): 13 for column in range(6, 6 + len(MONTHS))})
    style_sheet(matrix, matrix_widths)
    style_sheet(dictionary, {"A": 22, "B": 80, "C": 28})

    for row in demand.iter_rows(min_row=2, max_row=demand.max_row, min_col=1, max_col=10):
        row[0].number_format = "mmm yyyy"
        row[5].number_format = "#,##0"
        row[6].number_format = '#,##0.00 "EUR"'
        row[7].number_format = '#,##0.00 "EUR"'
    for row in master.iter_rows(min_row=2, max_row=master.max_row, min_col=1, max_col=9):
        row[5].number_format = '#,##0.00 "EUR"'
    for row in matrix.iter_rows(min_row=2, max_row=matrix.max_row, min_col=6, max_col=matrix.max_column):
        for cell in row:
            cell.number_format = "#,##0"

    add_table(master, "ProductMaster")
    add_table(demand, "DemandData")
    add_table(matrix, "DemandMatrix")
    add_table(dictionary, "DataDictionary")
    workbook.save(OUTPUT_PATH)
    print(f"Created {OUTPUT_PATH} with {len(all_skus)} SKUs and {demand.max_row - 1} demand records.")


if __name__ == "__main__":
    build_workbook()
