#!/usr/bin/env Rscript

"%||%" <- function(value, fallback) if (is.null(value) || length(value) == 0L) fallback else value

suppressPackageStartupMessages({
  library(fpp3)
  library(readxl)
  library(readr)
})

args <- commandArgs(trailingOnly = TRUE)
input_display <- if (length(args) >= 1L) args[[1L]] else "data/Demand_Genie_Synthetic_Demand_History.xlsx"
input_path <- normalizePath(input_display, mustWork = TRUE)
output_dir <- if (length(args) >= 2L) args[[2L]] else "data/forecast-v4"
sheet_name <- if (length(args) >= 3L) args[[3L]] else "Demand_Data"

initial_months <- 24L
evaluation_horizon <- 6L
origin_count <- 7L
forecast_horizon <- 12L

dir.create(output_dir, recursive = TRUE, showWarnings = FALSE)

required_columns <- c("Month", "SKU", "Demand_Units")
raw_demand <- read_excel(input_path, sheet = sheet_name, .name_repair = "minimal")
missing_columns <- setdiff(required_columns, names(raw_demand))
if (length(missing_columns) > 0L) {
  stop(sprintf("Missing required column(s): %s", paste(missing_columns, collapse = ", ")), call. = FALSE)
}

demand <- raw_demand |>
  mutate(
    Month = as.Date(Month),
    SKU = trimws(as.character(SKU)),
    Demand_Units = as.numeric(Demand_Units)
  ) |>
  select(any_of(c(
    "Month", "Product_Group_Code", "Product_Group", "SKU", "SKU_Description",
    "Demand_Units", "Demand_Profile"
  )))

if (any(is.na(demand$Month))) stop("Month contains missing or invalid values.", call. = FALSE)
if (any(is.na(demand$SKU) | demand$SKU == "")) stop("SKU contains missing or blank values.", call. = FALSE)
if (any(is.na(demand$Demand_Units))) stop("Demand_Units contains missing or non-numeric values.", call. = FALSE)
if (any(demand$Demand_Units < 0)) stop("Demand_Units must be non-negative.", call. = FALSE)
if (any(format(demand$Month, "%d") != "01")) stop("Month must be the first day of each monthly period.", call. = FALSE)

duplicates <- demand |> count(SKU, Month, name = "Rows") |> filter(Rows > 1L)
if (nrow(duplicates) > 0L) stop("Each SKU and Month must occur exactly once.", call. = FALSE)

monthly_calendar <- seq(min(demand$Month), max(demand$Month), by = "month")
missing_periods <- tidyr::crossing(SKU = sort(unique(demand$SKU)), Month = monthly_calendar) |>
  anti_join(select(demand, SKU, Month), by = c("SKU", "Month"))
if (nrow(missing_periods) > 0L) {
  stop(sprintf("Found %d missing SKU-month period(s).", nrow(missing_periods)), call. = FALSE)
}

history_counts <- demand |> count(SKU, name = "History_Months")
required_history <- initial_months + origin_count - 1L + evaluation_horizon
if (any(history_counts$History_Months < required_history)) {
  stop(sprintf("Every SKU needs at least %d months for the declared rolling-origin protocol.", required_history), call. = FALSE)
}

metadata_columns <- intersect(
  c("SKU", "Product_Group_Code", "Product_Group", "SKU_Description", "Demand_Profile"),
  names(demand)
)
sku_metadata <- demand |> select(all_of(metadata_columns)) |> distinct(SKU, .keep_all = TRUE)

demand_ts <- demand |>
  select(Month, SKU, Demand_Units) |>
  mutate(Month = yearmonth(Month)) |>
  as_tsibble(key = SKU, index = Month)

fit_candidates <- function(data) {
  data |>
    model(
      Mean = MEAN(Demand_Units),
      Naive = NAIVE(Demand_Units),
      Drift = RW(Demand_Units ~ drift()),
      Seasonal_Naive = SNAIVE(Demand_Units),
      ETS = ETS(Demand_Units),
      ARIMA = ARIMA(Demand_Units)
    )
}

backtest_parts <- vector("list", origin_count)
for (origin_index in seq_len(origin_count)) {
  training_months <- initial_months + origin_index - 1L
  training <- demand_ts |> group_by_key() |> slice_head(n = training_months)
  actual <- demand_ts |>
    group_by_key() |>
    slice((training_months + 1L):(training_months + evaluation_horizon)) |>
    as_tibble() |>
    transmute(SKU, Month, Actual_Units = Demand_Units)

  scales <- training |>
    as_tibble() |>
    group_by(SKU) |>
    arrange(Month, .by_group = TRUE) |>
    mutate(Seasonal_Error = Demand_Units - lag(Demand_Units, 12L)) |>
    summarise(
      MASE_Scale = mean(abs(Seasonal_Error), na.rm = TRUE),
      RMSSE_Scale = mean(Seasonal_Error^2, na.rm = TRUE),
      .groups = "drop"
    )

  forecast_rows <- fit_candidates(training) |>
    forecast(h = evaluation_horizon) |>
    hilo(level = c(80, 95)) |>
    unpack_hilo(cols = all_of(c("80%", "95%"))) |>
    as_tibble() |>
    transmute(
      SKU,
      Model = .model,
      Month,
      Forecast_Units = pmax(as.numeric(.mean), 0),
      PI80_Lower_Units = pmax(as.numeric(`80%_lower`), 0),
      PI80_Upper_Units = pmax(as.numeric(`80%_upper`), 0),
      PI95_Lower_Units = pmax(as.numeric(`95%_lower`), 0),
      PI95_Upper_Units = pmax(as.numeric(`95%_upper`), 0)
    ) |>
    inner_join(actual, by = c("SKU", "Month")) |>
    left_join(scales, by = "SKU") |>
    group_by(SKU, Model) |>
    arrange(Month, .by_group = TRUE) |>
    mutate(
      Origin = origin_index,
      Training_Months = training_months,
      Horizon_Step = row_number(),
      Error = Actual_Units - Forecast_Units,
      Absolute_Error = abs(Error),
      Squared_Error = Error^2,
      Scaled_Absolute_Error = if_else(MASE_Scale > 0, Absolute_Error / MASE_Scale, NA_real_),
      Scaled_Squared_Error = if_else(RMSSE_Scale > 0, Squared_Error / RMSSE_Scale, NA_real_),
      Covered_80 = Actual_Units >= PI80_Lower_Units & Actual_Units <= PI80_Upper_Units,
      Covered_95 = Actual_Units >= PI95_Lower_Units & Actual_Units <= PI95_Upper_Units
    ) |>
    ungroup()

  backtest_parts[[origin_index]] <- forecast_rows
}

backtest_predictions <- bind_rows(backtest_parts) |>
  mutate(Month = as.Date(Month)) |>
  left_join(sku_metadata, by = "SKU") |>
  relocate(any_of(c("Product_Group_Code", "Product_Group", "SKU_Description")), .before = SKU) |>
  arrange(SKU, Model, Origin, Horizon_Step)

classical_model_selection <- backtest_predictions |>
  group_by(SKU, Model) |>
  summarise(
    ME = mean(Error),
    RMSE = sqrt(mean(Squared_Error)),
    MAE = mean(Absolute_Error),
    MASE = mean(Scaled_Absolute_Error, na.rm = TRUE),
    RMSSE = sqrt(mean(Scaled_Squared_Error, na.rm = TRUE)),
    WAPE = sum(Absolute_Error) / sum(Actual_Units),
    Coverage_80 = mean(Covered_80),
    Coverage_95 = mean(Covered_95),
    Mean_PI80_Width = mean(PI80_Upper_Units - PI80_Lower_Units),
    Mean_PI95_Width = mean(PI95_Upper_Units - PI95_Lower_Units),
    Error_Sum = sum(Error),
    Absolute_Error_Sum = sum(Absolute_Error),
    Squared_Error_Sum = sum(Squared_Error),
    Actual_Units_Sum = sum(Actual_Units),
    Covered_80_Count = sum(Covered_80),
    Covered_95_Count = sum(Covered_95),
    Evaluation_Origins = n_distinct(Origin),
    Evaluation_Points = n(),
    Evaluation_Horizon = evaluation_horizon,
    .groups = "drop"
  ) |>
  left_join(sku_metadata, by = "SKU") |>
  relocate(any_of(c("Product_Group_Code", "Product_Group", "SKU_Description")), .before = SKU) |>
  arrange(SKU, RMSE, MAE)

classical_all_model_forecasts <- fit_candidates(demand_ts) |>
  forecast(h = forecast_horizon) |>
  hilo(level = c(80, 95)) |>
  unpack_hilo(cols = all_of(c("80%", "95%"))) |>
  as_tibble() |>
  transmute(
    Month = as.Date(Month),
    SKU,
    Model = .model,
    Forecast_Units = pmax(as.numeric(.mean), 0),
    PI80_Lower_Units = pmax(as.numeric(`80%_lower`), 0),
    PI80_Upper_Units = pmax(as.numeric(`80%_upper`), 0),
    PI95_Lower_Units = pmax(as.numeric(`95%_lower`), 0),
    PI95_Upper_Units = pmax(as.numeric(`95%_upper`), 0),
    Interval_Method = "Native fable 80% and 95% model intervals"
  ) |>
  left_join(sku_metadata, by = "SKU") |>
  relocate(any_of(c("Product_Group_Code", "Product_Group", "SKU_Description")), .before = SKU) |>
  arrange(SKU, Model, Month)

stl_components <- demand_ts |>
  model(STL = STL(Demand_Units ~ trend(window = 13) + season(window = "periodic"), robust = TRUE)) |>
  components() |>
  as_tibble() |>
  transmute(
    Month = as.Date(Month),
    SKU,
    Observed = Demand_Units,
    Trend = trend,
    Seasonal = season_year,
    Remainder = remainder,
    Seasonally_Adjusted = season_adjust
  ) |>
  left_join(sku_metadata, by = "SKU") |>
  relocate(any_of(c("Product_Group_Code", "Product_Group", "SKU_Description")), .before = SKU) |>
  arrange(SKU, Month)

strength_value <- function(signal, remainder) {
  denominator <- var(signal + remainder, na.rm = TRUE)
  if (!is.finite(denominator) || denominator <= 0) return(0)
  max(0, min(1, 1 - var(remainder, na.rm = TRUE) / denominator))
}

decomposition_features <- stl_components |>
  group_by(SKU) |>
  summarise(
    Trend_Strength = strength_value(Trend, Remainder),
    Seasonal_Strength = strength_value(Seasonal, Remainder),
    Remainder_SD = sd(Remainder),
    Seasonal_Cycles = n() / 12,
    Evidence_Quality = if_else(Seasonal_Cycles < 4, "Limited: fewer than four annual cycles", "Review"),
    .groups = "drop"
  ) |>
  left_join(sku_metadata, by = "SKU") |>
  relocate(any_of(c("Product_Group_Code", "Product_Group", "SKU_Description")), .before = SKU) |>
  arrange(SKU)

run_metadata <- tibble(
  Engine = "R fpp3",
  FPP3_Version = as.character(packageVersion("fpp3")),
  Fable_Version = as.character(packageVersion("fable")),
  Feasts_Version = as.character(packageVersion("feasts")),
  Input_File = input_display,
  Sheet = sheet_name,
  History_Start = as.character(min(demand$Month)),
  History_End = as.character(max(demand$Month)),
  SKU_Count = n_distinct(demand$SKU),
  Initial_Training_Months = initial_months,
  Evaluation_Horizon_Months = evaluation_horizon,
  Evaluation_Origins = origin_count,
  Final_Forecast_Horizon_Months = forecast_horizon,
  Selection_Metric = "Per-SKU rolling-origin RMSE; identical 7 x 6-month origins",
  Scaling = "MASE and RMSSE denominators use only each origin's training history with seasonal lag 12",
  Decomposition = "Robust STL, period 12, periodic seasonal window, trend window 13",
  Generated_At_UTC = format(Sys.time(), tz = "UTC", usetz = TRUE)
)

write_csv(backtest_predictions, file.path(output_dir, "classical_backtest_predictions.csv"), na = "")
write_csv(classical_model_selection, file.path(output_dir, "classical_model_selection.csv"), na = "")
write_csv(classical_all_model_forecasts, file.path(output_dir, "classical_all_model_forecasts.csv"), na = "")
write_csv(stl_components, file.path(output_dir, "decomposition.csv"), na = "")
write_csv(decomposition_features, file.path(output_dir, "decomposition_features.csv"), na = "")
write_csv(run_metadata, file.path(output_dir, "classical_run_metadata.csv"), na = "")

message("Classical rolling-origin analysis complete: ", normalizePath(output_dir))
