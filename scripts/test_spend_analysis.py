#!/usr/bin/env python3
"""Regression tests for strict spend and demand-calendar decision gates."""

from __future__ import annotations

from argparse import Namespace

from analyze_spend import item_segmentation, spend_quality


def test_strict_control_tolerance() -> None:
    control_total = 143_686_430.21
    rows = [
        {
            "Spend_Base": control_total - 1.0,
            "Base_Currency": "EUR",
            "Source_System": "TEST",
            "Transaction_ID": "TX-1",
            "Line_ID": "1",
            "Supplier_ID": "SUP-1",
            "Category_L1": "Test",
            "Supplier_Normalized_ID": "SUP-1",
            "Supplier_Parent_ID": "PAR-1",
            "On_Contract_Parsed": True,
            "Positive_Spend": control_total - 1.0,
        }
    ]
    quality = spend_quality(1, rows, {}, Namespace(control_total=control_total))
    control = next(row for row in quality if row["Check"] == "Control-total reconciliation")
    assert control["Status"] == "FAIL", "A EUR 1 mismatch must not pass a large-ledger tolerance"
    assert "tolerance 0.01" in control["Detail"]


def test_globally_missing_month_is_incomplete() -> None:
    rows = []
    for part in ("SKU-1", "SKU-2"):
        for period, demand in (("2026-01", 10.0), ("2026-03", 12.0)):
            rows.append(
                {
                    "Part_ID": part,
                    "Location": "PLANT-01",
                    "Period": period,
                    "Demand_Units": demand,
                    "Unit_Cost_Base": 5.0,
                    "Criticality": "Medium",
                    "Lifecycle_Status": "Active",
                }
            )
    args = Namespace(
        periods_per_year=12,
        xyz_x=0.5,
        xyz_y=1.0,
        adi_boundary=1.32,
        cv2_boundary=0.49,
        abc_a=0.8,
        abc_b=0.95,
    )
    segmentation, quality = item_segmentation(rows, args)
    assert all(row["Complete_History"] == "No" for row in segmentation)
    incomplete = next(row for row in quality if row["Check"] == "Incomplete demand histories")
    assert incomplete["Status"] == "FAIL"
    assert incomplete["Value"] == 2


def main() -> None:
    test_strict_control_tolerance()
    test_globally_missing_month_is_incomplete()
    print("Spend-analysis regression tests: PASS")


if __name__ == "__main__":
    main()
