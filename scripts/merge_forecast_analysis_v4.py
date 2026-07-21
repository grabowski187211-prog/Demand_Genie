#!/usr/bin/env python3
"""Merge fpp3 and TiRex2 results into an auditable Dashboard V4 forecast package."""

from __future__ import annotations

import argparse
import hashlib
import json
from datetime import UTC, datetime
from pathlib import Path

import numpy as np
import pandas as pd


COMPLEXITY = {
    "Mean": 1,
    "Naive": 1,
    "Seasonal_Naive": 1,
    "Drift": 2,
    "ETS": 3,
    "ARIMA": 4,
    "TiRex2": 5,
}
PRACTICAL_TOLERANCE = 0.02


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as source:
        for block in iter(lambda: source.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def select_models(scores: pd.DataFrame) -> pd.DataFrame:
    selected_parts: list[pd.DataFrame] = []
    for _, group in scores.groupby("SKU", sort=True):
        ranked = group.sort_values(["RMSE", "MAE", "Model"]).copy()
        best_rmse = float(ranked.iloc[0]["RMSE"])
        raw_best = str(ranked.iloc[0]["Model"])
        second_rmse = float(ranked.iloc[1]["RMSE"])
        eligible_limit = best_rmse * (1.0 + PRACTICAL_TOLERANCE)
        ranked["Within_2pct_of_Best"] = ranked["RMSE"] <= eligible_limit + 1e-12
        eligible = ranked[ranked["Within_2pct_of_Best"]].copy()
        eligible["Complexity_Rank"] = eligible["Model"].map(COMPLEXITY).fillna(99)
        selected_model = str(eligible.sort_values(["Complexity_Rank", "RMSE", "MAE", "Model"]).iloc[0]["Model"])
        ranked["Complexity_Rank"] = ranked["Model"].map(COMPLEXITY).fillna(99).astype(int)
        ranked["Raw_Best_Model"] = raw_best
        ranked["Selected_Model"] = selected_model
        ranked["Selected"] = ranked["Model"].eq(selected_model)
        ranked["Near_Tie"] = second_rmse <= eligible_limit + 1e-12
        ranked["Raw_Winner_Margin_Pct"] = (second_rmse - best_rmse) / best_rmse if best_rmse else 0.0
        ranked["Selection_Rationale"] = np.where(
            ranked["Selected"],
            np.where(
                selected_model == raw_best,
                "Lowest rolling-origin RMSE",
                "Simplest model within 2% of lowest rolling-origin RMSE",
            ),
            "Candidate",
        )
        ranked["RMSE_Rank"] = ranked["RMSE"].rank(method="min").astype(int)
        selected_parts.append(ranked)
    return pd.concat(selected_parts, ignore_index=True)


def build_model_summary(scores: pd.DataFrame) -> pd.DataFrame:
    raw_wins = scores[scores["Model"].eq(scores["Raw_Best_Model"])]["Model"].value_counts()
    selected_wins = scores[scores["Selected"]]["Model"].value_counts()
    rows: list[dict[str, object]] = []
    for model, group in scores.groupby("Model", sort=True):
        points = float(group["Evaluation_Points"].sum())
        actual = float(group["Actual_Units_Sum"].sum())
        rows.append(
            {
                "Model": model,
                "Engine": "TiRex2 0.1.1" if model == "TiRex2" else "R fpp3 1.0.2",
                "Selected_SKUs": int(selected_wins.get(model, 0)),
                "Raw_RMSE_Wins": int(raw_wins.get(model, 0)),
                "Macro_MASE": float(group["MASE"].mean()),
                "Macro_RMSSE": float(group["RMSSE"].mean()),
                "Global_RMSE": float(np.sqrt(group["Squared_Error_Sum"].sum() / points)),
                "Volume_WAPE": float(group["Absolute_Error_Sum"].sum() / actual),
                "Mean_Bias": float(group["Error_Sum"].sum() / points),
                "Empirical_Coverage_80": float(group["Covered_80_Count"].sum() / points),
                "Evaluation_Points": int(points),
                "Evaluation_Origins_Per_SKU": int(group["Evaluation_Origins"].min()),
                "Evaluation_Horizon": int(group["Evaluation_Horizon"].min()),
            }
        )
    return pd.DataFrame(rows).sort_values(["Macro_MASE", "Global_RMSE", "Model"])


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("forecast_dir", nargs="?", default="data/forecast-v4")
    parser.add_argument("--input", default="data/Demand_Genie_Synthetic_Demand_History.xlsx")
    args = parser.parse_args()

    forecast_dir = Path(args.forecast_dir).resolve()
    input_display = Path(args.input).as_posix()
    input_path = Path(args.input).resolve()
    required_paths = {
        "classical_scores": forecast_dir / "classical_model_selection.csv",
        "tirex_scores": forecast_dir / "tirex2_model_selection.csv",
        "classical_forecasts": forecast_dir / "classical_all_model_forecasts.csv",
        "tirex_forecasts": forecast_dir / "tirex2_all_model_forecasts.csv",
        "classical_backtests": forecast_dir / "classical_backtest_predictions.csv",
        "tirex_backtests": forecast_dir / "tirex2_backtest_predictions.csv",
        "decomposition": forecast_dir / "decomposition.csv",
        "decomposition_features": forecast_dir / "decomposition_features.csv",
        "tirex_metadata": forecast_dir / "tirex2_run_metadata.csv",
        "classical_metadata": forecast_dir / "classical_run_metadata.csv",
    }
    missing = [str(path) for path in required_paths.values() if not path.exists()]
    if missing:
        raise FileNotFoundError(f"Missing forecast artifact(s): {', '.join(missing)}")

    workbook_hash = sha256_file(input_path)
    tirex_metadata = pd.read_csv(required_paths["tirex_metadata"]).iloc[0]
    if str(tirex_metadata["Input_SHA256"]) != workbook_hash:
        raise ValueError("TiRex2 artifact input hash does not match the current workbook")

    classical_scores = pd.read_csv(required_paths["classical_scores"])
    tirex_scores = pd.read_csv(required_paths["tirex_scores"])
    for column in sorted(set(classical_scores.columns).union(tirex_scores.columns)):
        if column not in classical_scores:
            classical_scores[column] = np.nan
        if column not in tirex_scores:
            tirex_scores[column] = np.nan
    scores = pd.concat([classical_scores[tirex_scores.columns], tirex_scores], ignore_index=True)
    scores = select_models(scores)

    all_forecasts = pd.concat(
        [pd.read_csv(required_paths["classical_forecasts"]), pd.read_csv(required_paths["tirex_forecasts"])],
        ignore_index=True,
        sort=False,
    )
    selected_lookup = scores[scores["Selected"]][["SKU", "Selected_Model"]]
    selected_forecasts = all_forecasts.merge(selected_lookup, on="SKU", how="inner")
    selected_forecasts = selected_forecasts[selected_forecasts["Model"].eq(selected_forecasts["Selected_Model"])].copy()
    selected_forecasts = selected_forecasts.sort_values(["SKU", "Month"])
    if len(selected_forecasts) != scores["SKU"].nunique() * 12:
        raise ValueError("Selected forecasts must contain exactly 12 rows per SKU")
    selected_forecasts["Selected_Model"] = selected_forecasts["Model"]

    model_summary = build_model_summary(scores)
    backtests = pd.concat(
        [pd.read_csv(required_paths["classical_backtests"]), pd.read_csv(required_paths["tirex_backtests"])],
        ignore_index=True,
        sort=False,
    ).sort_values(["SKU", "Model", "Origin", "Horizon_Step"])

    selected_metrics = scores[scores["Selected"]].copy()
    features = pd.read_csv(required_paths["decomposition_features"])
    exceptions = selected_metrics.merge(
        features[["SKU", "Trend_Strength", "Seasonal_Strength", "Seasonal_Cycles", "Evidence_Quality"]],
        on="SKU",
        how="left",
    )
    exceptions["Calibration_Review"] = (exceptions["Coverage_80"] < 0.65) | (exceptions["Coverage_80"] > 0.95)
    exceptions["Review_Required"] = exceptions["Near_Tie"] | exceptions["Calibration_Review"]
    exceptions["Review_Note"] = exceptions.apply(
        lambda row: "; ".join(
            part
            for part in (
                "Near-tied candidate models; planner judgment is material" if row["Near_Tie"] else "",
                "Empirical 80% interval coverage needs review" if row["Calibration_Review"] else "",
                "TiRex2 provides a native 80% interval but no native 95% interval" if row["Selected_Model"] == "TiRex2" else "",
                "Only three annual cycles; seasonal evidence is limited",
            )
            if part
        ),
        axis=1,
    )

    classical_metadata = pd.read_csv(required_paths["classical_metadata"]).iloc[0]
    run_metadata = pd.DataFrame(
        [
            {
                "Forecast_Package": "Demand Genie V4",
                "Input_File": input_display,
                "Input_SHA256": workbook_hash,
                "History_Start": classical_metadata["History_Start"],
                "History_End": classical_metadata["History_End"],
                "SKU_Count": int(classical_metadata["SKU_Count"]),
                "Candidate_Models": ", ".join(sorted(scores["Model"].unique())),
                "Selection_Protocol": "Seven expanding-window origins; 24-30 training months; six forecast months per origin",
                "Selection_Metric": "Per-SKU RMSE over 42 out-of-sample points",
                "Practical_Equivalence_Rule": "Prefer the lowest-complexity candidate within 2% of minimum RMSE",
                "Scaling": classical_metadata["Scaling"],
                "Final_Forecast_Horizon_Months": 12,
                "Classical_Engine": f"fpp3 {classical_metadata['FPP3_Version']}",
                "Foundation_Model": f"{tirex_metadata['Model_Repository']} tirex-2 {tirex_metadata['Package_Version']}",
                "TiRex2_Revision": tirex_metadata["Model_Revision"],
                "TiRex2_Checkpoint_SHA256": tirex_metadata["Checkpoint_SHA256"],
                "TiRex2_Inference_Seconds": tirex_metadata["Inference_Seconds"],
                "Interval_Note": "Classical candidates expose native 80/95% model intervals; TiRex2 exposes native marginal q10-q90 only",
                "Decomposition": classical_metadata["Decomposition"],
                "Seasonality_Caveat": "36 monthly observations provide only three annual cycles",
                "Generated_At_UTC": datetime.now(UTC).isoformat(),
            }
        ]
    )

    output_paths = {
        "model_selection.csv": scores.sort_values(["SKU", "RMSE", "MAE"]),
        "model_summary.csv": model_summary,
        "all_model_forecasts.csv": all_forecasts.sort_values(["SKU", "Model", "Month"]),
        "forecast_results.csv": selected_forecasts,
        "rolling_origin_predictions.csv": backtests,
        "forecast_exceptions.csv": exceptions.sort_values("SKU"),
        "run_metadata.csv": run_metadata,
    }
    for name, frame in output_paths.items():
        frame.to_csv(forecast_dir / name, index=False)

    artifact_files = sorted(path for path in forecast_dir.iterdir() if path.is_file() and path.name != "artifact_manifest.json")
    manifest = {
        "package": "Demand Genie forecast analysis V4",
        "input_file": input_display,
        "input_sha256": workbook_hash,
        "generated_at_utc": datetime.now(UTC).isoformat(),
        "artifacts": {
            path.name: {"sha256": sha256_file(path), "bytes": path.stat().st_size}
            for path in artifact_files
        },
    }
    (forecast_dir / "artifact_manifest.json").write_text(json.dumps(manifest, indent=2), encoding="utf-8")
    print(
        f"Merged {scores['Model'].nunique()} candidates for {scores['SKU'].nunique()} SKUs; "
        f"TiRex2 selected for {int((selected_metrics['Selected_Model'] == 'TiRex2').sum())} SKUs"
    )


if __name__ == "__main__":
    main()
