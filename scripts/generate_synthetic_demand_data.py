"""Generate deterministic synthetic demand and DDMRP planning data for Demand Genie."""

from __future__ import annotations

import calendar
import math
import random
from datetime import date, timedelta
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo


OUTPUT_PATH = Path("data/Demand_Genie_Synthetic_Demand_History.xlsx")
SEED = 20260721
SNAPSHOT_DATE = date(2026, 7, 21)
SNAPSHOT_TIMESTAMP = "2026-07-21T06:00:00Z"
LOCATION = "PLANT-01"
MONTHS = [date(year, month, 1) for year in range(2023, 2027) for month in range(1, 13)][6:42]
SKU_MULTIPLIERS = (0.55, 0.68, 0.8, 0.95, 1.1, 1.3, 1.55, 1.85)
MAKE_GROUPS = {"GBX", "HYP", "CVL", "DRV", "CNV", "PCK"}

GROUPS = (
    {"code": "BRG", "name": "Industrial Bearings", "base": 760, "seasonality": 0.13, "trend": 0.004, "volatility": 0.06, "profile": "Seasonal", "price": 78, "lead_time": 42},
    {"code": "GBX", "name": "Gearbox Components", "base": 440, "seasonality": 0.09, "trend": 0.006, "volatility": 0.08, "profile": "Trending", "price": 215, "lead_time": 63},
    {"code": "HYP", "name": "Hydraulic Pumps", "base": 285, "seasonality": 0.16, "trend": 0.003, "volatility": 0.12, "profile": "Volatile", "price": 540, "lead_time": 70},
    {"code": "SEA", "name": "Sealing Systems", "base": 1220, "seasonality": 0.07, "trend": 0.002, "volatility": 0.05, "profile": "Stable", "price": 29, "lead_time": 28},
    {"code": "CVL", "name": "Control Valves", "base": 330, "seasonality": 0.11, "trend": -0.002, "volatility": 0.1, "profile": "Volatile", "price": 310, "lead_time": 56},
    {"code": "DRV", "name": "Electric Drives", "base": 245, "seasonality": 0.14, "trend": 0.007, "volatility": 0.09, "profile": "Trending", "price": 680, "lead_time": 77},
    {"code": "SNS", "name": "Sensor Modules", "base": 510, "seasonality": 0.05, "trend": 0.01, "volatility": 0.13, "profile": "Growth", "price": 168, "lead_time": 49},
    {"code": "CNV", "name": "Conveying Components", "base": 620, "seasonality": 0.17, "trend": 0.001, "volatility": 0.08, "profile": "Seasonal", "price": 115, "lead_time": 45},
    {"code": "FST", "name": "Fastening Kits", "base": 1880, "seasonality": 0.06, "trend": 0.002, "volatility": 0.07, "profile": "Stable", "price": 16, "lead_time": 21},
    {"code": "PCK", "name": "Packaging Equipment", "base": 165, "seasonality": 0.2, "trend": 0.005, "volatility": 0.15, "profile": "Intermittent", "price": 950, "lead_time": 84},
)

DESCRIPTORS = ("Core", "Compact", "Standard", "Heavy Duty", "Precision", "High Capacity", "Service Kit", "Premium")
PLANNERS = ("Alex Planner", "Sam Buyer", "Jordan Scheduler", "Taylor Planner")

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
    supply_type = "Make" if str(group["code"]) in MAKE_GROUPS else "Buy"
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
        "supply_type": supply_type,
        "sku_index": sku_index,
    }


def round_up(value: float, multiple: int) -> int:
    return int(math.ceil(value / multiple) * multiple) if value > 0 else 0


def days_in_month(month: date) -> int:
    return calendar.monthrange(month.year, month.month)[1]


def dlt_category(dlt_days: int) -> str:
    if dlt_days >= 60:
        return "Long"
    if dlt_days >= 35:
        return "Medium"
    return "Short"


def lead_time_factor(dlt_days: int) -> float:
    return 0.3 if dlt_days >= 60 else 0.5 if dlt_days >= 35 else 0.75


def variability_factor(xyz_class: str) -> float:
    return {"X": 0.25, "Y": 0.5, "Z": 0.75}[xyz_class]


def nfp_status(net_flow: int, top_red: float, top_yellow: float, top_green: float) -> str:
    if net_flow < 0:
        return "Black"
    if net_flow <= top_red:
        return "Red"
    if net_flow <= top_yellow:
        return "Yellow"
    if net_flow <= top_green:
        return "Green"
    return "Blue"


def planning_action(status: str) -> str:
    if status == "Black":
        return "Critical: release supply recommendation and investigate immediate flow risk"
    if status == "Red":
        return "Urgent: release supply recommendation and review execution risk"
    if status == "Yellow":
        return "Replenish: release supply recommendation"
    if status == "Blue":
        return "No routine replenishment: review excess before changing open supply"
    return "No routine replenishment"


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


def format_columns(sheet, *, dates: tuple[str, ...] = (), integers: tuple[str, ...] = (), decimals: tuple[str, ...] = ()) -> None:
    header_columns = {cell.value: cell.column for cell in sheet[1]}
    for header in dates:
        if header in header_columns:
            for row in range(2, sheet.max_row + 1):
                sheet.cell(row, header_columns[header]).number_format = "yyyy-mm-dd"
    for header in integers:
        if header in header_columns:
            for row in range(2, sheet.max_row + 1):
                sheet.cell(row, header_columns[header]).number_format = "#,##0"
    for header in decimals:
        if header in header_columns:
            for row in range(2, sheet.max_row + 1):
                sheet.cell(row, header_columns[header]).number_format = "#,##0.00"


def build_ddmrp_rows(
    all_skus: list[dict[str, str | float | int]], histories: dict[str, list[int]], rng: random.Random
) -> dict[str, list[list[object]]]:
    """Build item-location DDMRP master, operational source, and recommendation records."""
    history_days = sum(days_in_month(month) for month in MONTHS[-12:])
    output = {
        "decoupling": [],
        "master": [],
        "inventory": [],
        "supply": [],
        "qualified_demand": [],
        "positions": [],
        "recommendations": [],
        "adjustments": [],
    }

    for position_index, sku_data in enumerate(all_skus):
        sku = str(sku_data["sku"])
        history = histories[sku]
        supply_type = str(sku_data["supply_type"])
        nominal_lead_time = int(sku_data["lead_time"])
        dlt_days = nominal_lead_time if supply_type == "Buy" else max(7, int(round(nominal_lead_time * 0.65)))
        baseline_adu = sum(history[-12:]) / history_days
        ltf = lead_time_factor(dlt_days)
        vf = variability_factor(str(sku_data["xyz_class"]))
        order_cycle = 14 if supply_type == "Buy" else 7
        order_multiple = 50 if baseline_adu >= 50 else 25 if baseline_adu >= 20 else 10 if baseline_adu >= 5 else 1
        moq = round_up(baseline_adu * order_cycle * 0.8, order_multiple)
        spike_horizon = 21 if dlt_days >= 60 else 14
        spike_threshold = round_up(max(baseline_adu * 1.8, order_multiple), order_multiple)
        adjustment_factor = 1.0
        adjustment_reason = "No planned adjustment"
        adjustment_status = "None"
        adjustment_from = ""
        adjustment_through = ""

        if str(sku_data["product_group_code"]) in {"CNV", "PCK"} and position_index % 2 == 0:
            adjustment_factor = 1.15
            adjustment_reason = "Synthetic approved seasonal event for dashboard testing"
            adjustment_status = "Approved synthetic demo"
            adjustment_from = SNAPSHOT_DATE
            adjustment_through = date(2026, 9, 30)
        elif str(sku_data["product_group_code"]) == "BRG" and position_index % 7 == 0:
            adjustment_factor = 0.92
            adjustment_reason = "Synthetic approved maintenance-period adjustment"
            adjustment_status = "Approved synthetic demo"
            adjustment_from = SNAPSHOT_DATE
            adjustment_through = date(2026, 8, 31)

        adjusted_adu = baseline_adu * adjustment_factor
        yellow = adjusted_adu * dlt_days
        red_base = yellow * ltf
        red_safety = red_base * vf
        red = red_base + red_safety
        green = max(moq, adjusted_adu * order_cycle, adjusted_adu * dlt_days * ltf)
        top_red = red
        top_yellow = red + yellow
        top_green = top_yellow + green
        planner = PLANNERS[position_index % len(PLANNERS)]
        decoupling_reason = (
            "Supplier variability and customer lead-time protection"
            if supply_type == "Buy"
            else "Critical operation protection with upstream decoupling"
        )
        dlt_basis = "Supplier lead time at external source" if supply_type == "Buy" else "Synthetic BOM/routing path to upstream buffer"
        upstream_source = "Approved supplier schedule" if supply_type == "Buy" else "Upstream component buffer and routing"
        buffer_profile = f"{supply_type[0]}-{dlt_category(dlt_days)[0]}-{sku_data['xyz_class']}"

        today_demand = max(1, int(round(baseline_adu * rng.uniform(0.75, 1.3))))
        past_due_demand = int(round(baseline_adu * rng.uniform(0.2, 0.8))) if position_index % 4 == 0 else 0
        qualified_spike_demand = max(spike_threshold, int(round(baseline_adu * 2.4))) if position_index % 6 == 0 else 0
        qualified_demand = past_due_demand + today_demand + qualified_spike_demand

        status_pattern = position_index % 5
        if status_pattern == 0:
            target_nfp = -min(max(1, qualified_demand * 0.45), top_green * 0.08)
        elif status_pattern == 1:
            target_nfp = top_red * 0.6
        elif status_pattern == 2:
            target_nfp = top_red + yellow * 0.55
        elif status_pattern == 3:
            target_nfp = top_yellow + green * 0.5
        else:
            target_nfp = top_green * 1.15

        total_available = max(0, int(round(target_nfp + qualified_demand)))
        if position_index % 9 == 0:
            on_hand = min(total_available, max(0, int(round(top_red * 0.4))))
        else:
            on_hand_ratio = (0.4, 0.5, 0.6, 0.75, 0.85)[status_pattern]
            on_hand = int(round(total_available * on_hand_ratio))
        open_supply = max(0, total_available - on_hand)
        net_flow = on_hand + open_supply - qualified_demand
        status = nfp_status(net_flow, top_red, top_yellow, top_green)
        triggered = net_flow <= top_yellow
        base_recommendation = max(top_green - net_flow, 0) if triggered else 0
        recommended_order = round_up(base_recommendation, order_multiple)
        execution_review = "Review open supply timing and synchronization: on hand is at or below TOR" if on_hand <= top_red else "No automatic on-hand red-zone review"

        output["decoupling"].append([
            SNAPSHOT_DATE, LOCATION, sku, sku_data["description"], supply_type, "Approved", dlt_days, dlt_basis,
            upstream_source, decoupling_reason, planner, date(2026, 6, 30), date(2026, 10, 1), "Synthetic approved demo",
        ])
        output["master"].append([
            SNAPSHOT_DATE, LOCATION, sku, sku_data["product_group"], sku_data["description"], buffer_profile, supply_type,
            baseline_adu, "Trailing 12 complete months / calendar days", dlt_days, ltf, vf, moq, order_cycle,
            order_multiple, spike_horizon, spike_threshold, adjustment_factor, adjustment_status, planner,
            "Approved synthetic demo", date(2026, 10, 1),
        ])
        output["inventory"].append([
            SNAPSHOT_DATE, SNAPSHOT_TIMESTAMP, LOCATION, sku, sku_data["description"], on_hand, 0, on_hand,
            "Available", "Complete synthetic inventory snapshot",
        ])
        output["adjustments"].append([
            LOCATION, sku, SNAPSHOT_DATE, adjustment_from, adjustment_through, adjustment_factor, adjustment_status,
            adjustment_reason, "Synthetic event data; replace with FPP3 forecast proposal before operational use", planner,
        ])

        if past_due_demand > 0:
            output["qualified_demand"].append([
                f"SO-PD-{sku}", LOCATION, sku, SNAPSHOT_DATE - timedelta(days=1), "Past_Due", past_due_demand,
                "Yes", "Past-due sales order demand", spike_threshold,
            ])
        output["qualified_demand"].append([
            f"SO-TD-{sku}", LOCATION, sku, SNAPSHOT_DATE, "Today", today_demand, "Yes", "Sales order demand due today", spike_threshold,
        ])
        if qualified_spike_demand > 0:
            output["qualified_demand"].append([
                f"SO-SP-{sku}", LOCATION, sku, SNAPSHOT_DATE + timedelta(days=3 + (position_index % 8)), "Qualified_Spike",
                qualified_spike_demand, "Yes", "Future order exceeds approved spike threshold inside spike horizon", spike_threshold,
            ])

        if open_supply > 0:
            split_supply = position_index % 3 == 0 and open_supply > order_multiple * 2
            first_supply = int(round(open_supply * 0.6)) if split_supply else open_supply
            second_supply = open_supply - first_supply
            first_status = "Late" if position_index % 9 == 0 else "Firmed"
            first_receipt = SNAPSHOT_DATE - timedelta(days=2) if first_status == "Late" else SNAPSHOT_DATE + timedelta(days=max(1, int(round(dlt_days * 0.6))))
            output["supply"].append([
                f"{'PO' if supply_type == 'Buy' else 'MO'}-{sku}-01", LOCATION, sku, supply_type, first_supply,
                first_receipt, first_status, "Yes", "Supplier A" if supply_type == "Buy" else "Assembly Cell 1",
                "Late receipt requires execution review" if first_status == "Late" else "No current synchronization exception",
            ])
            if second_supply > 0:
                output["supply"].append([
                    f"{'PO' if supply_type == 'Buy' else 'MO'}-{sku}-02", LOCATION, sku, supply_type, second_supply,
                    SNAPSHOT_DATE + timedelta(days=max(1, int(round(dlt_days * 0.9)))), "Released", "Yes",
                    "Supplier B" if supply_type == "Buy" else "Assembly Cell 2", "No current synchronization exception",
                ])

        output["positions"].append([
            sku, LOCATION, sku_data["description"], sku_data["product_group"], baseline_adu, dlt_days, ltf, vf, moq,
            order_cycle, on_hand, open_supply, past_due_demand, today_demand, qualified_spike_demand, order_multiple,
            adjustment_factor, SNAPSHOT_TIMESTAMP,
        ])
        output["recommendations"].append([
            SNAPSHOT_DATE, LOCATION, sku, sku_data["description"], sku_data["product_group"], buffer_profile, baseline_adu,
            adjustment_factor, adjusted_adu, dlt_days, ltf, vf, moq, order_cycle, order_multiple, red_base, red_safety,
            red, yellow, green, top_red, top_yellow, top_green, on_hand, open_supply, past_due_demand, today_demand,
            qualified_spike_demand, qualified_demand, net_flow, net_flow / top_green * 100 if top_green else 0, status,
            "Yes" if triggered else "No", base_recommendation, recommended_order,
            SNAPSHOT_DATE + timedelta(days=dlt_days), planning_action(status), on_hand - top_red, execution_review,
            "Proposed - planner approval required", "Complete synthetic DDMRP input set",
        ])

    return output


def build_workbook() -> None:
    rng = random.Random(SEED)
    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)

    workbook = Workbook()
    read_me = workbook.active
    read_me.title = "Read_Me"
    read_me.append(["Demand Genie Synthetic Demand and DDMRP Planning Data"])
    read_me["A1"].fill = HEADER_FILL
    read_me["A1"].font = Font(color="FFFFFF", bold=True, size=15)
    read_me.merge_cells("A1:D1")
    read_me.append([])
    read_me.append(["Purpose", "Synthetic monthly demand and complete DDMRP planning data for dashboard development. All operational values are synthetic demo data, not live ERP data."])
    read_me.append(["Coverage", "36 complete months of demand: July 2023 through June 2026. DDMRP planning snapshot: 2026-07-21 at PLANT-01."])
    read_me.append(["Scope", "10 product groups, 8 SKUs per group, 80 total SKUs, 2,880 monthly demand records, and 80 calculator-ready DDMRP item-location positions."])
    read_me.append(["Forecasting", "Use Demand_Data for FPP3 forecast training. Forecast output should later create approval-required factors in DDMRP_Adjustments; it must not be inserted directly into qualified demand."])
    read_me.append(["DDMRP planning", "DDMRP_Positions contains the complete required input grain for the DDMRP calculator. DDMRP_Recommendations is the transparent synthetic output for the snapshot."])
    read_me.append(["Data separation", "Inventory_Snapshot, Open_Supply, and Qualified_Demand are source-level operational data. DDMRP_Master and DDMRP_Decoupling contain approved synthetic configuration and DLT design."])
    read_me.append(["Data note", "Demand includes trend, seasonality, volatility, shocks, and intermittent behavior. DDMRP positions intentionally include black, red, yellow, green, and blue net-flow states for dashboard testing."])
    read_me.append([])
    read_me.append(["Sheet", "Use"])
    read_me.append(["Demand_Data", "Tidy historical demand records for upload and FPP3 analysis."])
    read_me.append(["Demand_Matrix", "One row per SKU with one column per historical month."])
    read_me.append(["Product_Master", "SKU attributes, nominal lead time, price, ABC/XYZ, default location, and supply type."])
    read_me.append(["DDMRP_Decoupling", "Approved synthetic decoupling design and DLT basis for each item-location."])
    read_me.append(["DDMRP_Master", "Approved synthetic buffer profile, ADU, factors, MOQ, order constraints, and spike settings."])
    read_me.append(["Inventory_Snapshot", "Current usable on-hand inventory at the planning snapshot."])
    read_me.append(["Open_Supply", "Open purchase/manufacturing supply with expected receipt and execution risk."])
    read_me.append(["Qualified_Demand", "Past-due, today, and qualified future spike demand used by the Net Flow Equation."])
    read_me.append(["DDMRP_Positions", "Calculator-ready daily position file with required DDMRP input columns."])
    read_me.append(["DDMRP_Recommendations", "Pre-calculated transparent synthetic order recommendations and execution review flags."])
    read_me.append(["DDMRP_Adjustments", "Current synthetic dynamic adjustment schedule. Replace with approved FPP3-derived proposals later."])
    read_me.append(["Data_Dictionary", "Field definitions for Demand_Data."])
    read_me.append(["DDMRP_Data_Dictionary", "Field definitions for the DDMRP sheets."])
    for row in read_me.iter_rows(min_row=3, max_row=read_me.max_row, min_col=1, max_col=2):
        row[0].font = HEADER_FONT
        row[0].fill = SUBHEADER_FILL
        row[0].border = Border(bottom=THIN_LINE)
        row[1].alignment = Alignment(wrap_text=True, vertical="top")
    read_me.column_dimensions["A"].width = 25
    read_me.column_dimensions["B"].width = 115
    read_me.sheet_view.showGridLines = False

    master = workbook.create_sheet("Product_Master")
    master.append([
        "Product_Group_Code", "Product_Group", "SKU", "SKU_Description", "Demand_Profile", "Unit_Price_EUR",
        "Lead_Time_Days", "ABC_Class", "XYZ_Class", "Default_Location", "Supply_Type",
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
    histories: dict[str, list[int]] = {}
    for group in GROUPS:
        for sku_index in range(8):
            sku_data = make_sku(group, sku_index)
            all_skus.append(sku_data)
            master.append([
                sku_data["product_group_code"], sku_data["product_group"], sku_data["sku"], sku_data["description"],
                sku_data["profile"], sku_data["unit_price"], sku_data["lead_time"], sku_data["abc_class"], sku_data["xyz_class"],
                LOCATION, sku_data["supply_type"],
            ])

            history: list[int] = []
            for month_index, month in enumerate(MONTHS):
                demand_units = calculate_demand(group, sku_index, month_index, rng)
                history.append(demand_units)
                demand.append([
                    month, sku_data["product_group_code"], sku_data["product_group"], sku_data["sku"], sku_data["description"],
                    demand_units, sku_data["unit_price"], round(demand_units * float(sku_data["unit_price"]), 2),
                    sku_data["lead_time"], sku_data["profile"],
                ])
            histories[str(sku_data["sku"])] = history
            matrix.append([
                sku_data["product_group_code"], sku_data["product_group"], sku_data["sku"], sku_data["description"],
                sku_data["profile"], *history,
            ])

    ddmrp = build_ddmrp_rows(all_skus, histories, rng)
    decoupling = workbook.create_sheet("DDMRP_Decoupling")
    decoupling.append([
        "Snapshot_Date", "Location", "SKU", "SKU_Description", "Supply_Type", "Decoupling_Status", "DLT_Days", "DLT_Basis",
        "Upstream_Source", "Decoupling_Rationale", "Planner_Owner", "Approved_On", "Next_Review_Date", "Approval_Status",
    ])
    for row in ddmrp["decoupling"]:
        decoupling.append(row)

    ddmrp_master = workbook.create_sheet("DDMRP_Master")
    ddmrp_master.append([
        "Snapshot_Date", "Location", "SKU", "Product_Group", "SKU_Description", "Buffer_Profile", "Supply_Type", "Baseline_ADU",
        "ADU_Calculation_Method", "DLT_Days", "Lead_Time_Factor", "Variability_Factor", "MOQ", "Order_Cycle_Days",
        "Order_Multiple", "Spike_Horizon_Days", "Spike_Threshold_Units", "Demand_Adjustment_Factor", "Adjustment_Status",
        "Planner_Owner", "Setting_Approval_Status", "Next_Review_Date",
    ])
    for row in ddmrp["master"]:
        ddmrp_master.append(row)

    inventory = workbook.create_sheet("Inventory_Snapshot")
    inventory.append([
        "Snapshot_Date", "Source_Snapshot_At", "Location", "SKU", "SKU_Description", "On_Hand_Units", "Quality_Hold_Units",
        "Usable_On_Hand_Units", "Inventory_Status", "Data_Quality_Status",
    ])
    for row in ddmrp["inventory"]:
        inventory.append(row)

    open_supply = workbook.create_sheet("Open_Supply")
    open_supply.append([
        "Supply_Order_ID", "Location", "SKU", "Supply_Type", "Open_Supply_Units", "Expected_Receipt_Date", "Order_Status",
        "Eligible_for_NFP", "Supplier_or_Work_Center", "Synchronization_Note",
    ])
    for row in ddmrp["supply"]:
        open_supply.append(row)

    qualified_demand = workbook.create_sheet("Qualified_Demand")
    qualified_demand.append([
        "Demand_Reference", "Location", "SKU", "Due_Date", "Demand_Type", "Demand_Units", "Qualified_for_NFP",
        "Qualification_Reason", "Spike_Threshold_Units",
    ])
    for row in ddmrp["qualified_demand"]:
        qualified_demand.append(row)

    positions = workbook.create_sheet("DDMRP_Positions")
    positions.append([
        "SKU", "Location", "SKU_Description", "Product_Group", "ADU", "DLT_Days", "Lead_Time_Factor", "Variability_Factor",
        "MOQ", "Order_Cycle_Days", "On_Hand", "Open_Supply", "Past_Due_Demand", "Today_Demand", "Qualified_Spike_Demand",
        "Order_Multiple", "Demand_Adjustment_Factor", "Source_Snapshot_At",
    ])
    for row in ddmrp["positions"]:
        positions.append(row)

    recommendations = workbook.create_sheet("DDMRP_Recommendations")
    recommendations.append([
        "Planning_Date", "Location", "SKU", "SKU_Description", "Product_Group", "Buffer_Profile", "ADU", "Demand_Adjustment_Factor",
        "Adjusted_ADU", "DLT_Days", "Lead_Time_Factor", "Variability_Factor", "MOQ", "Order_Cycle_Days", "Order_Multiple",
        "Red_Base", "Red_Safety", "Red_Zone", "Yellow_Zone", "Green_Zone", "TOR", "TOY", "TOG", "On_Hand", "Open_Supply",
        "Past_Due_Demand", "Today_Demand", "Qualified_Spike_Demand", "Qualified_Demand", "Net_Flow_Position", "NFP_Percent_of_TOG",
        "NFP_Status", "Replenishment_Triggered", "Base_Recommended_Units", "Recommended_Order_Units", "Suggested_Availability_Date",
        "Planning_Action", "On_Hand_Gap_to_TOR", "Execution_Review", "Planner_Release_Status", "Data_Status",
    ])
    for row in ddmrp["recommendations"]:
        recommendations.append(row)

    adjustments = workbook.create_sheet("DDMRP_Adjustments")
    adjustments.append([
        "Location", "SKU", "Snapshot_Date", "Effective_From", "Effective_Through", "Demand_Adjustment_Factor", "Adjustment_Status",
        "Adjustment_Reason", "Forecast_Integration_Note", "Planner_Owner",
    ])
    for row in ddmrp["adjustments"]:
        adjustments.append(row)

    dictionary = workbook.create_sheet("Data_Dictionary")
    dictionary.append(["Field", "Description", "Example"])
    for row in (
        ("Month", "First day of the monthly demand period. Use this field as the time axis.", "2026-06-01"),
        ("Product_Group_Code", "Stable code for the main product group.", "BRG"),
        ("Product_Group", "Main product group label.", "Industrial Bearings"),
        ("SKU", "Unique stock-keeping unit identifier.", "BRG-01"),
        ("SKU_Description", "Readable SKU description.", "Industrial Bearings Core"),
        ("Demand_Units", "Observed monthly demand quantity in units.", "428"),
        ("Unit_Price_EUR", "Synthetic price per unit in EUR.", "60.84"),
        ("Demand_Value_EUR", "Demand_Units multiplied by Unit_Price_EUR.", "26039.52"),
        ("Lead_Time_Days", "Nominal replenishment context; use approved DLT for DDMRP buffers.", "42"),
        ("Demand_Profile", "Designed behavior pattern to aid forecasting tests.", "Seasonal"),
    ):
        dictionary.append(row)

    ddmrp_dictionary = workbook.create_sheet("DDMRP_Data_Dictionary")
    ddmrp_dictionary.append(["Sheet", "Field", "Description"])
    for row in (
        ("DDMRP_Decoupling", "DLT_Days", "Approved decoupled lead time for the item-location; calculated from source/BOM design in a real implementation."),
        ("DDMRP_Master", "Baseline_ADU", "Trailing 12 complete months of demand divided by calendar days. Recalculate from approved demand/forecast policy."),
        ("DDMRP_Master", "Lead_Time_Factor", "Factor used in red base and green calculations; grouped by DLT category in this synthetic example."),
        ("DDMRP_Master", "Variability_Factor", "Factor used in red safety; grouped by synthetic XYZ class in this example."),
        ("DDMRP_Master", "Demand_Adjustment_Factor", "Time-effective dynamic factor applied to ADU. This workbook uses synthetic demo adjustments; replace with approved FPP3 proposals."),
        ("Inventory_Snapshot", "Usable_On_Hand_Units", "Inventory available for the buffered position at the snapshot. This is the On_Hand input to net flow."),
        ("Open_Supply", "Eligible_for_NFP", "Only eligible open supply is included in Net Flow Position. Receipt timing remains an execution risk."),
        ("Qualified_Demand", "Qualified_for_NFP", "Past due, today, and qualifying future spikes are included. A monthly forecast is not qualified demand."),
        ("DDMRP_Positions", "Open_Supply", "Aggregate of eligible open supply records for SKU + Location at the snapshot."),
        ("DDMRP_Positions", "Qualified_Spike_Demand", "Future daily demand above the spike threshold and inside the spike horizon."),
        ("DDMRP_Recommendations", "Net_Flow_Position", "On hand + open supply - qualified demand."),
        ("DDMRP_Recommendations", "Recommended_Order_Units", "TOG - NFP when NFP <= TOY, rounded up to the valid order multiple."),
        ("DDMRP_Recommendations", "Execution_Review", "On-hand and synchronization warning separate from the net-flow supply recommendation."),
        ("DDMRP_Adjustments", "Forecast_Integration_Note", "Forecast-derived adjustments require FPP3 validation and planner approval before operational use."),
    ):
        ddmrp_dictionary.append(row)

    style_sheet(master, {"A": 20, "B": 25, "C": 14, "D": 34, "E": 16, "F": 16, "G": 16, "H": 12, "I": 12, "J": 18, "K": 14})
    style_sheet(demand, {"A": 14, "B": 20, "C": 25, "D": 14, "E": 34, "F": 14, "G": 16, "H": 18, "I": 16, "J": 16})
    matrix_widths = {"A": 20, "B": 25, "C": 14, "D": 34, "E": 16}
    matrix_widths.update({get_column_letter(column): 13 for column in range(6, 6 + len(MONTHS))})
    style_sheet(matrix, matrix_widths)
    style_sheet(decoupling, {"A": 14, "B": 14, "C": 14, "D": 34, "E": 12, "F": 18, "G": 12, "H": 42, "I": 34, "J": 50, "K": 18, "L": 14, "M": 16, "N": 24})
    style_sheet(ddmrp_master, {"A": 14, "B": 14, "C": 14, "D": 25, "E": 34, "F": 16, "G": 12, "H": 14, "I": 42, "J": 12, "K": 16, "L": 17, "M": 12, "N": 18, "O": 15, "P": 18, "Q": 20, "R": 24, "S": 22, "T": 18, "U": 25, "V": 16})
    style_sheet(inventory, {"A": 14, "B": 24, "C": 14, "D": 14, "E": 34, "F": 18, "G": 18, "H": 22, "I": 16, "J": 34})
    style_sheet(open_supply, {"A": 20, "B": 14, "C": 14, "D": 14, "E": 20, "F": 22, "G": 14, "H": 18, "I": 24, "J": 42})
    style_sheet(qualified_demand, {"A": 20, "B": 14, "C": 14, "D": 14, "E": 18, "F": 14, "G": 18, "H": 48, "I": 22})
    style_sheet(positions, {"A": 14, "B": 14, "C": 34, "D": 25, "E": 12, "F": 12, "G": 16, "H": 17, "I": 12, "J": 18, "K": 14, "L": 16, "M": 18, "N": 14, "O": 24, "P": 15, "Q": 24, "R": 24})
    style_sheet(recommendations, {get_column_letter(index): 18 for index in range(1, 42)})
    recommendations.column_dimensions["D"].width = 34
    recommendations.column_dimensions["E"].width = 25
    recommendations.column_dimensions["AG"].width = 54
    recommendations.column_dimensions["AM"].width = 44
    recommendations.column_dimensions["AO"].width = 30
    style_sheet(adjustments, {"A": 14, "B": 14, "C": 14, "D": 16, "E": 18, "F": 24, "G": 24, "H": 46, "I": 70, "J": 18})
    style_sheet(dictionary, {"A": 22, "B": 80, "C": 28})
    style_sheet(ddmrp_dictionary, {"A": 26, "B": 28, "C": 105})

    format_columns(master, integers=("Lead_Time_Days",))
    format_columns(demand, dates=("Month",), integers=("Demand_Units", "Lead_Time_Days"))
    for row in demand.iter_rows(min_row=2, max_row=demand.max_row, min_col=1, max_col=10):
        row[6].number_format = '#,##0.00 "EUR"'
        row[7].number_format = '#,##0.00 "EUR"'
    for row in master.iter_rows(min_row=2, max_row=master.max_row, min_col=1, max_col=master.max_column):
        row[5].number_format = '#,##0.00 "EUR"'
    for row in matrix.iter_rows(min_row=2, max_row=matrix.max_row, min_col=6, max_col=matrix.max_column):
        for cell in row:
            cell.number_format = "#,##0"
    format_columns(decoupling, dates=("Snapshot_Date", "Approved_On", "Next_Review_Date"), integers=("DLT_Days",))
    format_columns(ddmrp_master, dates=("Snapshot_Date", "Next_Review_Date"), integers=("DLT_Days", "MOQ", "Order_Cycle_Days", "Order_Multiple", "Spike_Horizon_Days", "Spike_Threshold_Units"), decimals=("Baseline_ADU", "Lead_Time_Factor", "Variability_Factor", "Demand_Adjustment_Factor"))
    format_columns(inventory, dates=("Snapshot_Date",), integers=("On_Hand_Units", "Quality_Hold_Units", "Usable_On_Hand_Units"))
    format_columns(open_supply, dates=("Expected_Receipt_Date",), integers=("Open_Supply_Units",))
    format_columns(qualified_demand, dates=("Due_Date",), integers=("Demand_Units", "Spike_Threshold_Units"))
    format_columns(positions, integers=("DLT_Days", "MOQ", "Order_Cycle_Days", "On_Hand", "Open_Supply", "Past_Due_Demand", "Today_Demand", "Qualified_Spike_Demand", "Order_Multiple"), decimals=("ADU", "Lead_Time_Factor", "Variability_Factor", "Demand_Adjustment_Factor"))
    format_columns(recommendations, dates=("Planning_Date", "Suggested_Availability_Date"), integers=("DLT_Days", "MOQ", "Order_Cycle_Days", "Order_Multiple", "On_Hand", "Open_Supply", "Past_Due_Demand", "Today_Demand", "Qualified_Spike_Demand", "Qualified_Demand", "Net_Flow_Position", "Base_Recommended_Units", "Recommended_Order_Units"), decimals=("ADU", "Demand_Adjustment_Factor", "Adjusted_ADU", "Lead_Time_Factor", "Variability_Factor", "Red_Base", "Red_Safety", "Red_Zone", "Yellow_Zone", "Green_Zone", "TOR", "TOY", "TOG", "NFP_Percent_of_TOG", "On_Hand_Gap_to_TOR"))
    format_columns(adjustments, dates=("Snapshot_Date", "Effective_From", "Effective_Through"), decimals=("Demand_Adjustment_Factor",))

    add_table(master, "ProductMaster")
    add_table(demand, "DemandData")
    add_table(matrix, "DemandMatrix")
    add_table(decoupling, "DDMRPDecoupling")
    add_table(ddmrp_master, "DDMRPMaster")
    add_table(inventory, "InventorySnapshot")
    add_table(open_supply, "OpenSupply")
    add_table(qualified_demand, "QualifiedDemand")
    add_table(positions, "DDMRPPositions")
    add_table(recommendations, "DDMRPRecommendations")
    add_table(adjustments, "DDMRPAdjustments")
    add_table(dictionary, "DataDictionary")
    add_table(ddmrp_dictionary, "DDMRPDataDictionary")

    workbook.save(OUTPUT_PATH)
    print(
        f"Created {OUTPUT_PATH} with {len(all_skus)} SKUs, {demand.max_row - 1} demand records, "
        f"and {recommendations.max_row - 1} DDMRP recommendations."
    )


if __name__ == "__main__":
    build_workbook()
