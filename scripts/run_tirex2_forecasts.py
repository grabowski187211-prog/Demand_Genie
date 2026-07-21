#!/usr/bin/env python3
"""Run TiRex2 on the same rolling origins used by the fpp3 V4 pipeline."""

from __future__ import annotations

import argparse
import hashlib
import json
import math
import platform
import time
from importlib.metadata import version
from pathlib import Path

import numpy as np
import pandas as pd
import torch
import yaml
from huggingface_hub import snapshot_download
from tirex2 import TimeseriesType, load_model


REPO_ID = "NX-AI/TiRex-2"
INITIAL_MONTHS = 24
EVALUATION_HORIZON = 6
ORIGIN_COUNT = 7
FORECAST_HORIZON = 12
MODEL_NAME = "TiRex2"


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as source:
        for block in iter(lambda: source.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def pinball_loss(actual: float, prediction: float, quantile: float) -> float:
    error = actual - prediction
    return max(quantile * error, (quantile - 1.0) * error)


def rearrange_quantiles(prediction: np.ndarray) -> tuple[np.ndarray, np.ndarray]:
    """Clip at zero and monotonically rearrange marginal quantiles per horizon."""
    clipped = np.maximum(prediction, 0.0)
    raw_crossings = np.sum(np.diff(clipped, axis=0) < 0, axis=0)
    return np.sort(clipped, axis=0), raw_crossings


def forecast_batch(model, values: np.ndarray, length: int, horizon: int, batch_size: int) -> tuple[list[np.ndarray], float]:
    timeseries = [
        TimeseriesType(
            target=torch.from_numpy(row[:length].astype(np.float32, copy=False)).unsqueeze(0),
            past_covariates=None,
            future_covariates=None,
        )
        for row in values
    ]
    started = time.monotonic()
    result = model.forecast(
        timeseries,
        prediction_length=horizon,
        output_type="numpy",
        batch_size=batch_size,
        tta_diff=True,
        tta_sign_flip=False,
    )
    return result, time.monotonic() - started


def summarize_scores(frame: pd.DataFrame) -> pd.DataFrame:
    rows: list[dict[str, object]] = []
    for sku, group in frame.groupby("SKU", sort=True):
        errors = group["Error"].to_numpy(float)
        abs_errors = np.abs(errors)
        squared_errors = errors**2
        actual = group["Actual_Units"].to_numpy(float)
        rows.append(
            {
                "Product_Group_Code": group["Product_Group_Code"].iloc[0],
                "Product_Group": group["Product_Group"].iloc[0],
                "SKU_Description": group["SKU_Description"].iloc[0],
                "SKU": sku,
                "Model": MODEL_NAME,
                "ME": float(errors.mean()),
                "RMSE": float(np.sqrt(squared_errors.mean())),
                "MAE": float(abs_errors.mean()),
                "MASE": float(group["Scaled_Absolute_Error"].mean()),
                "RMSSE": float(np.sqrt(group["Scaled_Squared_Error"].mean())),
                "WAPE": float(abs_errors.sum() / actual.sum()),
                "Coverage_80": float(group["Covered_80"].mean()),
                "Coverage_95": math.nan,
                "Mean_PI80_Width": float((group["PI80_Upper_Units"] - group["PI80_Lower_Units"]).mean()),
                "Mean_PI95_Width": math.nan,
                "Mean_Pinball_Loss": float(group["Mean_Pinball_Loss"].mean()),
                "Quantile_Crossings": int(group["Quantile_Crossings"].sum()),
                "Raw_Quantile_Crossings": int(group["Raw_Quantile_Crossings"].sum()),
                "Rearranged_Points": int(group["Quantile_Rearranged"].sum()),
                "Error_Sum": float(errors.sum()),
                "Absolute_Error_Sum": float(abs_errors.sum()),
                "Squared_Error_Sum": float(squared_errors.sum()),
                "Actual_Units_Sum": float(actual.sum()),
                "Covered_80_Count": int(group["Covered_80"].sum()),
                "Covered_95_Count": 0,
                "Evaluation_Origins": int(group["Origin"].nunique()),
                "Evaluation_Points": int(len(group)),
                "Evaluation_Horizon": EVALUATION_HORIZON,
                "Demand_Profile": group["Demand_Profile"].iloc[0],
            }
        )
    return pd.DataFrame(rows)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser()
    parser.add_argument("input", nargs="?", default="data/Demand_Genie_Synthetic_Demand_History.xlsx")
    parser.add_argument("output_dir", nargs="?", default="data/forecast-v4")
    parser.add_argument("--sheet", default="Demand_Data")
    parser.add_argument("--batch-size", type=int, default=8)
    parser.add_argument("--local-only", action="store_true")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    input_display = Path(args.input).as_posix()
    input_path = Path(args.input).resolve()
    output_dir = Path(args.output_dir).resolve()
    output_dir.mkdir(parents=True, exist_ok=True)
    if args.batch_size < 1:
        raise ValueError("batch size must be at least one")

    raw = pd.read_excel(input_path, sheet_name=args.sheet)
    required = {"Month", "SKU", "Demand_Units"}
    missing = required.difference(raw.columns)
    if missing:
        raise ValueError(f"Missing required columns: {', '.join(sorted(missing))}")
    demand = raw.copy()
    demand["Month"] = pd.to_datetime(demand["Month"])
    demand["SKU"] = demand["SKU"].astype(str).str.strip()
    demand["Demand_Units"] = pd.to_numeric(demand["Demand_Units"], errors="raise")
    if demand[["Month", "SKU", "Demand_Units"]].isna().any().any():
        raise ValueError("Demand data contains missing key or value fields")
    if (demand["Demand_Units"] < 0).any():
        raise ValueError("Demand_Units must be non-negative")
    if demand.duplicated(["SKU", "Month"]).any():
        raise ValueError("Each SKU and Month must occur exactly once")

    counts = demand.groupby("SKU")["Month"].nunique()
    required_history = INITIAL_MONTHS + ORIGIN_COUNT - 1 + EVALUATION_HORIZON
    if not counts.eq(required_history).all():
        raise ValueError(f"Every SKU must have exactly {required_history} complete monthly observations")

    demand = demand.sort_values(["SKU", "Month"]).reset_index(drop=True)
    skus = sorted(demand["SKU"].unique())
    matrices = [demand.loc[demand["SKU"].eq(sku), "Demand_Units"].to_numpy(np.float32) for sku in skus]
    months = [demand.loc[demand["SKU"].eq(sku), "Month"].to_numpy() for sku in skus]
    values = np.stack(matrices)
    metadata = {
        sku: demand.loc[demand["SKU"].eq(sku)].iloc[0].to_dict()
        for sku in skus
    }

    snapshot_started = time.monotonic()
    snapshot_path = Path(
        snapshot_download(
            repo_id=REPO_ID,
            allow_patterns=["model-config.yaml", "model.ckpt"],
            local_files_only=args.local_only,
        )
    )
    snapshot_seconds = time.monotonic() - snapshot_started
    checkpoint_path = snapshot_path / "model.ckpt"
    config_path = snapshot_path / "model-config.yaml"
    config = yaml.safe_load(config_path.read_text(encoding="utf-8"))

    torch.manual_seed(0)
    load_started = time.monotonic()
    model = load_model(snapshot_path, device="cpu")
    model_load_seconds = time.monotonic() - load_started
    quantiles = [float(value) for value in model._quantile_levels()]
    if quantiles != sorted(quantiles):
        raise RuntimeError(f"TiRex2 quantile levels are not ordered: {quantiles}")
    q50_index = quantiles.index(0.5)
    q10_index = quantiles.index(0.1)
    q90_index = quantiles.index(0.9)
    quantile_columns = [f"Q{int(round(q * 100)):02d}_Units" for q in quantiles]

    inference_seconds = 0.0
    backtest_rows: list[dict[str, object]] = []
    for origin in range(1, ORIGIN_COUNT + 1):
        training_months = INITIAL_MONTHS + origin - 1
        forecasts, elapsed = forecast_batch(model, values, training_months, EVALUATION_HORIZON, args.batch_size)
        inference_seconds += elapsed
        for sku_index, sku in enumerate(skus):
            prediction = np.asarray(forecasts[sku_index], dtype=float)[0]
            if prediction.shape != (len(quantiles), EVALUATION_HORIZON):
                raise RuntimeError(f"Unexpected TiRex2 output shape for {sku}: {prediction.shape}")
            train = values[sku_index, :training_months].astype(float)
            seasonal_errors = train[12:] - train[:-12]
            mase_scale = float(np.mean(np.abs(seasonal_errors)))
            rmsse_scale = float(np.mean(seasonal_errors**2))
            rearranged, raw_crossings = rearrange_quantiles(prediction)
            for step in range(EVALUATION_HORIZON):
                actual = float(values[sku_index, training_months + step])
                point = float(rearranged[q50_index, step])
                error = actual - point
                q_values = rearranged[:, step]
                row = {
                    "Product_Group_Code": metadata[sku].get("Product_Group_Code", ""),
                    "Product_Group": metadata[sku].get("Product_Group", ""),
                    "SKU_Description": metadata[sku].get("SKU_Description", ""),
                    "SKU": sku,
                    "Model": MODEL_NAME,
                    "Month": pd.Timestamp(months[sku_index][training_months + step]).date().isoformat(),
                    "Forecast_Units": point,
                    "PI80_Lower_Units": float(rearranged[q10_index, step]),
                    "PI80_Upper_Units": float(rearranged[q90_index, step]),
                    "PI95_Lower_Units": math.nan,
                    "PI95_Upper_Units": math.nan,
                    "Actual_Units": actual,
                    "MASE_Scale": mase_scale,
                    "RMSSE_Scale": rmsse_scale,
                    "Origin": origin,
                    "Training_Months": training_months,
                    "Horizon_Step": step + 1,
                    "Error": error,
                    "Absolute_Error": abs(error),
                    "Squared_Error": error**2,
                    "Scaled_Absolute_Error": abs(error) / mase_scale if mase_scale > 0 else math.nan,
                    "Scaled_Squared_Error": error**2 / rmsse_scale if rmsse_scale > 0 else math.nan,
                    "Covered_80": float(rearranged[q10_index, step]) <= actual <= float(rearranged[q90_index, step]),
                    "Covered_95": math.nan,
                    "Mean_Pinball_Loss": float(np.mean([pinball_loss(actual, qv, q) for qv, q in zip(q_values, quantiles)])),
                    "Quantile_Crossings": int(np.sum(np.diff(q_values) < 0)),
                    "Raw_Quantile_Crossings": int(raw_crossings[step]),
                    "Quantile_Rearranged": bool(raw_crossings[step]),
                    "Demand_Profile": metadata[sku].get("Demand_Profile", ""),
                }
                row.update({column: float(value) for column, value in zip(quantile_columns, q_values)})
                backtest_rows.append(row)

    backtest = pd.DataFrame(backtest_rows)
    scores = summarize_scores(backtest)

    final_forecasts, elapsed = forecast_batch(model, values, values.shape[1], FORECAST_HORIZON, args.batch_size)
    inference_seconds += elapsed
    final_rows: list[dict[str, object]] = []
    for sku_index, sku in enumerate(skus):
        prediction = np.asarray(final_forecasts[sku_index], dtype=float)[0]
        rearranged, raw_crossings = rearrange_quantiles(prediction)
        future_months = pd.date_range(
            pd.Timestamp(months[sku_index][-1]) + pd.offsets.MonthBegin(1),
            periods=FORECAST_HORIZON,
            freq="MS",
        )
        for step, month in enumerate(future_months):
            q_values = rearranged[:, step]
            row = {
                "Product_Group_Code": metadata[sku].get("Product_Group_Code", ""),
                "Product_Group": metadata[sku].get("Product_Group", ""),
                "SKU_Description": metadata[sku].get("SKU_Description", ""),
                "Month": month.date().isoformat(),
                "SKU": sku,
                "Model": MODEL_NAME,
                "Forecast_Units": float(rearranged[q50_index, step]),
                "PI80_Lower_Units": float(rearranged[q10_index, step]),
                "PI80_Upper_Units": float(rearranged[q90_index, step]),
                "PI95_Lower_Units": math.nan,
                "PI95_Upper_Units": math.nan,
                "Interval_Method": "TiRex2 marginal q10-q90 after monotone rearrangement; no native 95% interval",
                "Demand_Profile": metadata[sku].get("Demand_Profile", ""),
                "Quantile_Crossings": int(np.sum(np.diff(q_values) < 0)),
                "Raw_Quantile_Crossings": int(raw_crossings[step]),
                "Quantile_Rearranged": bool(raw_crossings[step]),
            }
            row.update({column: float(value) for column, value in zip(quantile_columns, q_values)})
            final_rows.append(row)
    final_forecast = pd.DataFrame(final_rows)

    metadata_row = {
        "Engine": "TiRex2",
        "Package": "tirex-2",
        "Package_Version": version("tirex-2"),
        "Model_Repository": REPO_ID,
        "Model_Revision": snapshot_path.name,
        "Checkpoint_SHA256": sha256_file(checkpoint_path),
        "Checkpoint_Bytes": checkpoint_path.stat().st_size,
        "Input_File": input_display,
        "Input_SHA256": sha256_file(input_path),
        "Sheet": args.sheet,
        "History_Start": demand["Month"].min().date().isoformat(),
        "History_End": demand["Month"].max().date().isoformat(),
        "SKU_Count": len(skus),
        "Initial_Training_Months": INITIAL_MONTHS,
        "Evaluation_Horizon_Months": EVALUATION_HORIZON,
        "Evaluation_Origins": ORIGIN_COUNT,
        "Final_Forecast_Horizon_Months": FORECAST_HORIZON,
        "Point_Forecast": "q50",
        "Native_Interval": "q10-q90 (80% marginal interval)",
        "Quantiles": ",".join(str(q) for q in quantiles),
        "Quantile_Postprocessing": "Non-negative clipping followed by monotone rearrangement per forecast horizon",
        "Device": "cpu",
        "Batch_Size": args.batch_size,
        "TTA_Diff": True,
        "TTA_Sign_Flip": False,
        "Snapshot_Resolution_Seconds": round(snapshot_seconds, 3),
        "Model_Load_Seconds": round(model_load_seconds, 3),
        "Inference_Seconds": round(inference_seconds, 3),
        "Python": platform.python_version(),
        "PyTorch": torch.__version__,
        "Model_Context_Limit": config.get("context_len"),
        "Model_Horizon_Limit": config.get("future_len"),
        "Model_Patch_Size": config.get("patch_size"),
        "Generated_At_UTC": pd.Timestamp.now(tz="UTC").isoformat(),
    }

    backtest.to_csv(output_dir / "tirex2_backtest_predictions.csv", index=False)
    scores.to_csv(output_dir / "tirex2_model_selection.csv", index=False)
    final_forecast.to_csv(output_dir / "tirex2_all_model_forecasts.csv", index=False)
    pd.DataFrame([metadata_row]).to_csv(output_dir / "tirex2_run_metadata.csv", index=False)
    (output_dir / "tirex2_run_metadata.json").write_text(json.dumps(metadata_row, indent=2), encoding="utf-8")
    print(
        f"TiRex2 complete: {len(backtest):,} backtest rows, {len(final_forecast):,} final rows, "
        f"{inference_seconds:.1f}s inference"
    )


if __name__ == "__main__":
    main()
