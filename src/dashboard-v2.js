(() => {
  "use strict";

  const REQUIRED_SHEETS = ["Demand_Data", "DDMRP_Positions"];
  const SHEET_PURPOSES = {
    Demand_Data: "Monthly history for time-series analysis",
    Product_Master: "Item attributes and supply policy",
    DDMRP_Decoupling: "Approved decoupling-point rationale",
    DDMRP_Master: "Buffer profile and planning parameters",
    Inventory_Snapshot: "Usable on-hand by item and location",
    Open_Supply: "Eligible open purchase and production supply",
    Qualified_Demand: "Past-due, today, and qualified spike demand",
    DDMRP_Positions: "Atomic net-flow calculation inputs",
    DDMRP_Recommendations: "Auditable baseline calculation output",
    DDMRP_Adjustments: "Approved demand adjustment history",
    Forecast_Results: "Point forecasts and prediction intervals",
    Model_Selection: "Holdout model accuracy and selection",
  };
  const STATUS_ORDER = { Black: 0, Red: 1, Yellow: 2, Green: 3, Blue: 4 };
  const MODEL_LABELS = {
    Auto: "Automatic winner",
    Mean: "Mean",
    Naive: "Naive",
    Drift: "Drift",
    Seasonal_Naive: "Seasonal naive",
    ETS: "ETS",
    ARIMA: "ARIMA",
  };
  const state = {
    data: cloneData(SAMPLE_DATA),
    view: "planning",
    scenario: "baseline",
    group: "All",
    location: "All",
    status: "All",
    search: "",
    selectedKey: null,
    selectedSku: null,
    selectedBufferKey: null,
    horizon: 6,
    forecastModel: "Auto",
    actionLog: [],
    decisions: {},
    pendingAction: null,
    sourceMode: "fpp3",
  };

  const el = (id) => document.getElementById(id);
  const num = (value, fallback = 0) => {
    const parsed = Number(value);
    return Number.isFinite(parsed) ? parsed : fallback;
  };
  const round = (value, digits = 0) => {
    const factor = 10 ** digits;
    return Math.round(num(value) * factor) / factor;
  };
  const clamp = (value, low, high) => Math.min(high, Math.max(low, value));
  const esc = (value) => String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
  const fmt = (value, digits = 0) => Number.isFinite(Number(value))
    ? new Intl.NumberFormat("en-US", { maximumFractionDigits: digits, minimumFractionDigits: digits }).format(Number(value))
    : "-";
  const pct = (value) => `${fmt(num(value) * 100, 0)}%`;
  const dateValue = (value) => {
    if (value instanceof Date && !Number.isNaN(value.getTime())) return value;
    if (typeof value === "number" && window.XLSX?.SSF) {
      const parsed = XLSX.SSF.parse_date_code(value);
      if (parsed) return new Date(Date.UTC(parsed.y, parsed.m - 1, parsed.d));
    }
    const date = new Date(value);
    return Number.isNaN(date.getTime()) ? null : date;
  };
  const isoDate = (value) => {
    const date = dateValue(value);
    return date ? date.toISOString().slice(0, 10) : "-";
  };
  const monthLabel = (value, short = true) => {
    const date = dateValue(value);
    return date ? new Intl.DateTimeFormat("en", { month: short ? "short" : "long", year: "numeric", timeZone: "UTC" }).format(date) : "-";
  };
  const addDays = (value, days) => {
    const date = dateValue(value) || new Date();
    date.setUTCDate(date.getUTCDate() + Number(days));
    return date.toISOString().slice(0, 10);
  };
  const addMonths = (value, months) => {
    const date = dateValue(value) || new Date();
    return new Date(Date.UTC(date.getUTCFullYear(), date.getUTCMonth() + months, 1)).toISOString().slice(0, 10);
  };
  const daysInMonth = (value) => {
    const date = dateValue(value);
    return date ? new Date(Date.UTC(date.getUTCFullYear(), date.getUTCMonth() + 1, 0)).getUTCDate() : 30;
  };
  const keyFor = (row) => `${row.SKU}|${row.Location || ""}`;

  function cloneData(data) {
    return JSON.parse(JSON.stringify(data));
  }

  function normalizeRows(rows) {
    return (rows || []).map((row) => Object.fromEntries(Object.entries(row).map(([key, value]) => {
      const normalizedKey = String(key).trim();
      if (value instanceof Date) {
        const year = value.getFullYear();
        const month = String(value.getMonth() + 1).padStart(2, "0");
        const day = String(value.getDate()).padStart(2, "0");
        return [normalizedKey, `${year}-${month}-${day}`];
      }
      return [normalizedKey, value];
    })));
  }

  function sheet(name) {
    return state.data.sheets?.[name] || [];
  }

  function productLookup() {
    const lookup = new Map();
    sheet("Product_Master").forEach((row) => lookup.set(row.SKU, row));
    sheet("Demand_Data").forEach((row) => {
      if (!lookup.has(row.SKU)) lookup.set(row.SKU, row);
    });
    return lookup;
  }

  function snapshotDate() {
    const metadata = sheet("Forecast_Run_Metadata")[0];
    const positionDate = sheet("DDMRP_Positions")[0]?.Source_Snapshot_At;
    const masterDate = sheet("DDMRP_Master")[0]?.Snapshot_Date;
    return isoDate(positionDate || masterDate || metadata?.Last_Observed_Month || new Date());
  }

  function fppSelection(sku) {
    const rows = sheet("Model_Selection").filter((row) => row.SKU === sku);
    return rows.find((row) => row.Selected === true || String(row.Selected).toUpperCase() === "TRUE") || rows[0] || null;
  }

  function forecastFactor(sku) {
    const position = sheet("DDMRP_Positions").find((row) => row.SKU === sku);
    const baselineAdu = num(position?.ADU);
    if (!baselineAdu) return num(position?.Demand_Adjustment_Factor, 1) || 1;
    const forecast = getForecastBundle(sku, "Auto", 1).forecast[0];
    if (!forecast) return num(position?.Demand_Adjustment_Factor, 1) || 1;
    return clamp((num(forecast.point) / daysInMonth(forecast.month)) / baselineAdu, 0.25, 3);
  }

  function calculatePosition(position, scenario = state.scenario) {
    const approvedFactor = num(position.Demand_Adjustment_Factor, 1) || 1;
    const factor = scenario === "forecast" ? forecastFactor(position.SKU) : approvedFactor;
    const adu = num(position.ADU);
    const dlt = num(position.DLT_Days);
    const ltf = num(position.Lead_Time_Factor);
    const vf = num(position.Variability_Factor);
    const adjustedAdu = adu * factor;
    const yellow = adjustedAdu * dlt;
    const redBase = yellow * ltf;
    const redSafety = redBase * vf;
    const red = redBase + redSafety;
    const green = Math.max(
      num(position.MOQ),
      adjustedAdu * num(position.Order_Cycle_Days),
      adjustedAdu * dlt * ltf,
    );
    const tor = red;
    const toy = red + yellow;
    const tog = toy + green;
    const qualifiedDemand = num(position.Past_Due_Demand) + num(position.Today_Demand) + num(position.Qualified_Spike_Demand);
    const nfp = num(position.On_Hand) + num(position.Open_Supply) - qualifiedDemand;
    let status = "Blue";
    if (nfp < 0) status = "Black";
    else if (nfp <= tor) status = "Red";
    else if (nfp <= toy) status = "Yellow";
    else if (nfp <= tog) status = "Green";
    const triggered = nfp <= toy;
    const baseQty = Math.max(0, tog - nfp);
    const multiple = Math.max(1, num(position.Order_Multiple, 1));
    const recommended = triggered ? Math.ceil(baseQty / multiple) * multiple : 0;
    const availability = addDays(snapshotDate(), dlt);
    return {
      ...position,
      factor,
      approvedFactor,
      adjustedAdu,
      yellow,
      redBase,
      redSafety,
      red,
      green,
      tor,
      toy,
      tog,
      qualifiedDemand,
      nfp,
      nfpPct: tog ? nfp / tog : 0,
      status,
      triggered,
      recommended,
      availability,
      executionRisk: num(position.On_Hand) <= tor,
      key: keyFor(position),
    };
  }

  function allPositions() {
    const products = productLookup();
    return sheet("DDMRP_Positions").map((row) => {
      const product = products.get(row.SKU) || {};
      return calculatePosition({
        ...row,
        Product_Group: row.Product_Group || product.Product_Group || "Unassigned",
        SKU_Description: row.SKU_Description || product.SKU_Description || row.SKU,
      });
    });
  }

  function filteredPositions() {
    const query = state.search.trim().toLowerCase();
    return allPositions().filter((row) => {
      if (state.group !== "All" && row.Product_Group !== state.group) return false;
      if (state.location !== "All" && row.Location !== state.location) return false;
      if (state.status !== "All" && row.status !== state.status) return false;
      return !query || `${row.SKU} ${row.SKU_Description}`.toLowerCase().includes(query);
    }).sort((a, b) => STATUS_ORDER[a.status] - STATUS_ORDER[b.status] || a.nfpPct - b.nfpPct || a.SKU.localeCompare(b.SKU));
  }

  function statusMarkup(status) {
    const normalized = Object.hasOwn(STATUS_ORDER, status) ? status : "Neutral";
    return `<span class="status status-${normalized.toLowerCase()}">${esc(status)}</span>`;
  }

  function bufferMarkup(row, large = false) {
    const total = Math.max(row.tog, 1);
    const redShare = clamp((row.red / total) * 100, 0, 100);
    const yellowShare = clamp((row.yellow / total) * 100, 0, 100 - redShare);
    const greenShare = Math.max(0, 100 - redShare - yellowShare);
    const marker = clamp((row.nfp / total) * 100, 0, 100);
    return `<div class="${large ? "detail-buffer" : ""}">
      <div class="buffer-bar" style="--red-share:${redShare}%;--yellow-share:${yellowShare}%;--green-share:${greenShare}%;--marker-position:${marker}%" aria-label="Net flow ${fmt(row.nfp)} within top of green ${fmt(row.tog)}">
        <span class="red"></span><span class="yellow"></span><span class="green"></span><i class="buffer-marker"></i>
      </div>
      <div class="buffer-caption"><span>NFP ${fmt(row.nfp)}</span><span>TOG ${fmt(row.tog)}</span></div>
    </div>`;
  }

  function decisionLabel(row) {
    const decision = state.decisions[row.key];
    if (!decision) return row.triggered ? "Proposed" : "No release";
    return decision.action === "Release" ? "Released" : "Held";
  }

  function renderPlanning() {
    const positions = filteredPositions();
    if (!state.selectedKey || !positions.some((row) => row.key === state.selectedKey)) state.selectedKey = positions[0]?.key || null;
    const critical = positions.filter((row) => row.status === "Black" || row.status === "Red");
    const quantity = positions.reduce((sum, row) => sum + row.recommended, 0);
    const risks = positions.filter((row) => row.executionRisk).length;
    el("critical-count").textContent = fmt(critical.length);
    el("critical-context").textContent = `of ${positions.length} positions`;
    el("release-quantity").textContent = fmt(quantity);
    el("execution-risk-count").textContent = fmt(risks);
    el("planning-headline").textContent = critical.length
      ? `${critical.length} critical net-flow ${critical.length === 1 ? "position requires" : "positions require"} review.`
      : "No black or red net-flow positions in the current filter.";
    el("planning-subline").textContent = state.scenario === "forecast"
      ? "Forecast-derived ADU factors are a planning proposal and do not change qualified demand."
      : "Approved factors are active; positions are ranked by status and buffer penetration.";
    el("filter-context").textContent = `${positions.length} buffered positions | ${snapshotDate()}`;
    el("planning-body").innerHTML = positions.length ? positions.map((row) => `<tr data-selectable data-key="${esc(row.key)}" tabindex="0" class="${row.key === state.selectedKey ? "selected" : ""}">
      <td>${statusMarkup(row.status)}</td>
      <td class="item-cell"><strong>${esc(row.SKU)}</strong><span title="${esc(row.SKU_Description)}">${esc(row.SKU_Description)}</span></td>
      <td class="buffer-cell">${bufferMarkup(row)}</td>
      <td class="number">${fmt(row.On_Hand)}</td>
      <td class="number">${fmt(row.Open_Supply)}</td>
      <td class="number">${fmt(row.qualifiedDemand)}</td>
      <td class="number ${row.recommended ? "critical-text" : ""}">${fmt(row.recommended)}</td>
      <td class="availability-date">${esc(row.availability)}</td>
    </tr>`).join("") : '<tr><td colspan="8" class="empty-state">No positions match the active filters.</td></tr>';
    document.querySelectorAll("#planning-body tr[data-selectable]").forEach((row) => {
      const select = () => { state.selectedKey = row.dataset.key; renderPlanning(); };
      row.addEventListener("click", select);
      row.addEventListener("keydown", (event) => { if (event.key === "Enter" || event.key === " ") select(); });
    });
    renderPlanningDetail(positions.find((row) => row.key === state.selectedKey));
  }

  function renderPlanningDetail(row) {
    if (!row) {
      el("planning-detail").innerHTML = '<p class="empty-state">Select a position to inspect its planning evidence.</p>';
      return;
    }
    const action = row.status === "Black"
      ? "Investigate the immediate flow deficit and release or expedite supply."
      : row.triggered
        ? "Review the recommendation and release replenishment to restore top of green."
        : "No replenishment trigger. Monitor execution and demand qualification.";
    const scenarioText = state.scenario === "forecast"
      ? `Forecast proposal factor ${fmt(row.factor, 2)} replaces approved factor ${fmt(row.approvedFactor, 2)} for this scenario.`
      : `Approved demand adjustment factor ${fmt(row.factor, 2)} is applied.`;
    el("planning-detail").innerHTML = `<header class="detail-header">
      <p class="section-kicker">Planning evidence</p>
      <div class="detail-title-row"><h2>${esc(row.SKU)} / ${esc(row.Location)}</h2>${statusMarkup(row.status)}</div>
      <p>${esc(row.SKU_Description)} | ${esc(row.Product_Group)}</p>
    </header>
    <section class="detail-section">
      <h3>Net-flow position</h3>${bufferMarkup(row, true)}
      <dl class="formula-list">
        <div class="formula-row"><dt>Usable on hand</dt><dd>${fmt(row.On_Hand)}</dd></div>
        <div class="formula-row"><dt>Eligible open supply</dt><dd>+ ${fmt(row.Open_Supply)}</dd></div>
        <div class="formula-row"><dt>Qualified demand</dt><dd>- ${fmt(row.qualifiedDemand)}</dd></div>
        <div class="formula-row"><dt>Net-flow position</dt><dd>${fmt(row.nfp)}</dd></div>
      </dl>
    </section>
    <section class="detail-section">
      <h3>Qualified demand</h3>
      <dl class="formula-list">
        <div class="formula-row"><dt>Past due</dt><dd>${fmt(row.Past_Due_Demand)}</dd></div>
        <div class="formula-row"><dt>Due today</dt><dd>${fmt(row.Today_Demand)}</dd></div>
        <div class="formula-row"><dt>Qualified spikes</dt><dd>${fmt(row.Qualified_Spike_Demand)}</dd></div>
      </dl>
      <p>Forecast demand is excluded from this net-flow demand term.</p>
    </section>
    <section class="detail-section">
      <h3>Buffer calculation</h3>
      <p>${esc(scenarioText)}</p>
      <dl class="formula-list">
        <div class="formula-row"><dt>Adjusted ADU</dt><dd>${fmt(row.ADU, 2)} x ${fmt(row.factor, 2)} = ${fmt(row.adjustedAdu, 2)}</dd></div>
        <div class="formula-row"><dt>Top of red</dt><dd>${fmt(row.tor)}</dd></div>
        <div class="formula-row"><dt>Top of yellow</dt><dd>${fmt(row.toy)}</dd></div>
        <div class="formula-row"><dt>Top of green</dt><dd>${fmt(row.tog)}</dd></div>
      </dl>
    </section>
    <section class="detail-section">
      <h3>Planner decision</h3>
      <p class="detail-callout ${row.status === "Black" || row.status === "Red" ? "critical" : ""}">${esc(action)}</p>
      <dl class="formula-list">
        <div class="formula-row"><dt>Recommended order</dt><dd>${fmt(row.recommended)} units</dd></div>
        <div class="formula-row"><dt>Suggested availability</dt><dd>${esc(row.availability)}</dd></div>
        <div class="formula-row"><dt>Decision status</dt><dd>${esc(decisionLabel(row))}</dd></div>
      </dl>
      <div class="detail-actions">
        <button class="command primary planner-action" type="button" data-action="Release" ${row.recommended ? "" : "disabled"}>Release ${fmt(row.recommended)}</button>
        <button class="command planner-action" type="button" data-action="Hold">Hold / review</button>
      </div>
    </section>`;
    document.querySelectorAll(".planner-action").forEach((button) => button.addEventListener("click", () => openActionDialog(row, button.dataset.action)));
  }

  function historyFor(sku) {
    return sheet("Demand_Data")
      .filter((row) => row.SKU === sku)
      .map((row) => ({ month: isoDate(row.Month), value: num(row.Demand_Units) }))
      .sort((a, b) => a.month.localeCompare(b.month));
  }

  function modelForecast(values, model, horizon) {
    const n = values.length;
    if (!n) return Array(horizon).fill(0);
    if (model === "Mean") return Array(horizon).fill(values.reduce((a, b) => a + b, 0) / n);
    if (model === "Naive") return Array(horizon).fill(values[n - 1]);
    if (model === "Drift") {
      const slope = n > 1 ? (values[n - 1] - values[0]) / (n - 1) : 0;
      return Array.from({ length: horizon }, (_, i) => values[n - 1] + slope * (i + 1));
    }
    if (model === "Seasonal_Naive" && n >= 12) {
      return Array.from({ length: horizon }, (_, i) => values[n - 12 + (i % 12)]);
    }
    if (model === "ETS") return etsForecast(values, horizon);
    return Array(horizon).fill(values[n - 1]);
  }

  function etsForecast(values, horizon) {
    const n = values.length;
    const seasonal = n >= 24;
    const period = seasonal ? 12 : 1;
    const firstMean = values.slice(0, period).reduce((a, b) => a + b, 0) / period;
    const seasons = seasonal ? values.slice(0, period).map((value) => value - firstMean) : [0];
    let level = values[0];
    let trend = n > 1 ? values[1] - values[0] : 0;
    const alpha = 0.35;
    const beta = 0.12;
    const gamma = 0.2;
    values.forEach((value, index) => {
      if (!index) return;
      const seasonIndex = index % period;
      const priorLevel = level;
      level = alpha * (value - seasons[seasonIndex]) + (1 - alpha) * (level + trend);
      trend = beta * (level - priorLevel) + (1 - beta) * trend;
      seasons[seasonIndex] = gamma * (value - level) + (1 - gamma) * seasons[seasonIndex];
    });
    return Array.from({ length: horizon }, (_, i) => Math.max(0, level + trend * (i + 1) + seasons[(n + i) % period]));
  }

  function evaluateModels(history) {
    const values = history.map((row) => row.value);
    const holdout = Math.min(6, Math.max(3, Math.floor(values.length * 0.2)));
    const train = values.slice(0, -holdout);
    const actual = values.slice(-holdout);
    const scaleErrors = train.slice(12).map((value, index) => Math.abs(value - train[index]));
    const squareScaleErrors = train.slice(12).map((value, index) => (value - train[index]) ** 2);
    const maseScale = scaleErrors.length ? scaleErrors.reduce((a, b) => a + b, 0) / scaleErrors.length : 1;
    const rmsseScale = squareScaleErrors.length ? squareScaleErrors.reduce((a, b) => a + b, 0) / squareScaleErrors.length : 1;
    return ["Seasonal_Naive", "ETS", "Drift", "Naive", "Mean"].map((model) => {
      const predicted = modelForecast(train, model, holdout);
      const errors = actual.map((value, index) => value - predicted[index]);
      const me = errors.reduce((a, b) => a + b, 0) / errors.length;
      const mae = errors.reduce((a, b) => a + Math.abs(b), 0) / errors.length;
      const rmse = Math.sqrt(errors.reduce((a, b) => a + b ** 2, 0) / errors.length);
      return {
        model,
        ME: me,
        MAE: mae,
        RMSE: rmse,
        MASE: maseScale ? mae / maseScale : null,
        RMSSE: rmsseScale ? rmse / Math.sqrt(rmsseScale) : null,
        errors,
      };
    }).sort((a, b) => a.RMSE - b.RMSE);
  }

  function browserForecast(sku, requestedModel, horizon) {
    const history = historyFor(sku);
    const comparisons = evaluateModels(history);
    const selected = requestedModel === "Auto" ? comparisons[0]?.model || "Naive" : requestedModel;
    const values = history.map((row) => row.value);
    const points = modelForecast(values, selected, horizon);
    const validation = comparisons.find((row) => row.model === selected) || comparisons[0];
    const sigma = Math.max(1, num(validation?.RMSE, 1));
    const lastMonth = history.at(-1)?.month || new Date().toISOString();
    const forecast = points.map((point, index) => {
      const growth = Math.sqrt(index + 1);
      return {
        month: addMonths(lastMonth, index + 1),
        point: Math.max(0, point),
        lower80: Math.max(0, point - 1.282 * sigma * growth),
        upper80: point + 1.282 * sigma * growth,
        lower95: Math.max(0, point - 1.96 * sigma * growth),
        upper95: point + 1.96 * sigma * growth,
      };
    });
    return { history, forecast, comparisons, selected, metrics: validation, source: "browser" };
  }

  function getForecastBundle(sku, requestedModel = state.forecastModel, horizon = state.horizon) {
    const history = historyFor(sku);
    const selectionRows = sheet("Model_Selection").filter((row) => row.SKU === sku);
    const selectedRow = fppSelection(sku);
    const packaged = sheet("Forecast_Results").filter((row) => row.SKU === sku).sort((a, b) => isoDate(a.Month).localeCompare(isoDate(b.Month)));
    if (state.sourceMode === "fpp3" && requestedModel === "Auto" && packaged.length) {
      const comparisons = selectionRows.map((row) => ({
        model: row[".model"] || row.Model || row.Selected_Model,
        ME: num(row.ME, NaN),
        RMSE: num(row.RMSE, NaN),
        MAE: num(row.MAE, NaN),
        MASE: num(row.MASE, NaN),
        RMSSE: num(row.RMSSE, NaN),
      })).filter((row) => row.model).sort((a, b) => a.RMSE - b.RMSE);
      return {
        history,
        forecast: packaged.slice(0, horizon).map((row) => ({
          month: isoDate(row.Month),
          point: num(row.Forecast_Units),
          lower80: num(row.PI80_Lower_Units),
          upper80: num(row.PI80_Upper_Units),
          lower95: num(row.PI95_Lower_Units),
          upper95: num(row.PI95_Upper_Units),
        })),
        comparisons,
        selected: selectedRow?.Selected_Model || selectedRow?.[".model"] || packaged[0]?.Selected_Model || "Selected model",
        metrics: comparisons.find((row) => row.model === selectedRow?.Selected_Model) || comparisons[0],
        source: "fpp3",
      };
    }
    return browserForecast(sku, requestedModel, horizon);
  }

  function renderForecast() {
    const sku = state.selectedSku;
    if (!sku) return;
    const product = productLookup().get(sku) || {};
    const bundle = getForecastBundle(sku);
    el("forecast-source-note").textContent = bundle.source === "fpp3"
      ? "R fpp3 | final-horizon holdout | 80% and 95% intervals"
      : "Browser fallback | final-horizon holdout | approximate intervals";
    el("engine-state").textContent = bundle.source === "fpp3" ? `Forecast engine: ${state.data.forecastEngine}` : "Forecast engine: local browser fallback";
    el("selected-model").textContent = MODEL_LABELS[bundle.selected] || bundle.selected;
    el("model-context").textContent = state.forecastModel === "Auto" ? "Automatic holdout winner" : "Planner-selected comparison";
    el("forecast-rmse").textContent = fmt(bundle.metrics?.RMSE, 1);
    el("forecast-bias").textContent = fmt(bundle.metrics?.ME, 1);
    el("forecast-headline").textContent = `${sku}: ${MODEL_LABELS[bundle.selected] || bundle.selected} selected for ${state.horizon} months.`;
    el("forecast-subline").textContent = bundle.source === "fpp3"
      ? "Evaluation uses withheld observations; forecast uncertainty remains visible."
      : "Uploaded history was evaluated locally; intervals are residual-scale approximations.";
    el("forecast-chart-title").textContent = `${sku} | ${product.SKU_Description || "Demand history"}`;
    renderForecastChart(bundle);
    renderModelComparison(bundle);
    renderAdjustments(sku, bundle);
  }

  function renderModelComparison(bundle) {
    const rows = bundle.comparisons.filter((row) => Number.isFinite(row.RMSE));
    const maxRmse = Math.max(...rows.map((row) => row.RMSE), 1);
    el("model-comparison").innerHTML = rows.length ? rows.map((row) => `<div class="model-row ${row.model === bundle.selected ? "selected" : ""}">
      <div class="model-name"><span>${esc(MODEL_LABELS[row.model] || row.model)}${row.model === bundle.selected ? " | selected" : ""}</span><span>RMSE ${fmt(row.RMSE, 1)}</span></div>
      <div class="model-meter" style="--meter:${clamp((row.RMSE / maxRmse) * 100, 2, 100)}%"><span></span></div>
      <div class="model-metrics"><span>MAE ${fmt(row.MAE, 1)}</span><span>Bias ${fmt(row.ME, 1)}</span><span>MASE ${fmt(row.MASE, 2)}</span></div>
    </div>`).join("") : '<p class="empty-state">No model evaluation is available for this item.</p>';
  }

  function renderForecastChart(bundle) {
    const svg = el("forecast-chart");
    const history = bundle.history.slice(-36);
    const future = bundle.forecast;
    const combined = [...history.map((row) => ({ month: row.month, value: row.value })), ...future.map((row) => ({ month: row.month, value: row.point }))];
    if (!combined.length) {
      svg.innerHTML = '<text x="20" y="40" class="svg-label">No demand history available.</text>';
      return;
    }
    const width = 900;
    const height = 360;
    const margin = { top: 16, right: 18, bottom: 44, left: 56 };
    const plotW = width - margin.left - margin.right;
    const plotH = height - margin.top - margin.bottom;
    const minValue = Math.min(0, ...future.map((row) => row.lower95));
    const maxValue = Math.max(...history.map((row) => row.value), ...future.map((row) => row.upper95), 1) * 1.05;
    const y = (value) => margin.top + plotH - ((value - minValue) / (maxValue - minValue || 1)) * plotH;
    const x = (index) => margin.left + (index / Math.max(1, combined.length - 1)) * plotW;
    const linePath = (points, offset, accessor) => points.map((row, index) => `${index ? "L" : "M"}${x(offset + index).toFixed(1)},${y(accessor(row)).toFixed(1)}`).join(" ");
    const areaPath = (points, offset, low, high) => {
      const upper = points.map((row, index) => `${index ? "L" : "M"}${x(offset + index).toFixed(1)},${y(high(row)).toFixed(1)}`).join(" ");
      const lower = [...points].reverse().map((row, reverseIndex) => {
        const index = points.length - 1 - reverseIndex;
        return `L${x(offset + index).toFixed(1)},${y(low(row)).toFixed(1)}`;
      }).join(" ");
      return `${upper} ${lower} Z`;
    };
    const actualPath = linePath(history, 0, (row) => row.value);
    const forecastOffset = Math.max(0, history.length - 1);
    const connectedForecast = history.length ? [{ month: history.at(-1).month, point: history.at(-1).value, lower80: history.at(-1).value, upper80: history.at(-1).value, lower95: history.at(-1).value, upper95: history.at(-1).value }, ...future] : future;
    const yTicks = Array.from({ length: 5 }, (_, i) => minValue + ((maxValue - minValue) * i) / 4);
    const labelEvery = combined.length > 30 ? 6 : 3;
    const xLabels = combined.map((row, index) => ({ row, index })).filter(({ index }) => index % labelEvery === 0 || index === combined.length - 1);
    svg.setAttribute("viewBox", `0 0 ${width} ${height}`);
    svg.innerHTML = `
      ${yTicks.map((tick) => `<line class="svg-grid" x1="${margin.left}" x2="${width - margin.right}" y1="${y(tick)}" y2="${y(tick)}"></line><text class="svg-label" x="${margin.left - 8}" y="${y(tick) + 3}" text-anchor="end">${fmt(tick)}</text>`).join("")}
      ${future.length ? `<path class="svg-pi95" d="${areaPath(connectedForecast, forecastOffset, (row) => row.lower95, (row) => row.upper95)}"></path><path class="svg-pi80" d="${areaPath(connectedForecast, forecastOffset, (row) => row.lower80, (row) => row.upper80)}"></path>` : ""}
      <line class="svg-axis" x1="${margin.left}" x2="${width - margin.right}" y1="${margin.top + plotH}" y2="${margin.top + plotH}"></line>
      ${history.length ? `<path class="svg-actual" d="${actualPath}"></path>` : ""}
      ${connectedForecast.length ? `<path class="svg-forecast" d="${linePath(connectedForecast, forecastOffset, (row) => row.point)}"></path>` : ""}
      ${future.length && history.length ? `<line class="svg-boundary" x1="${x(history.length - 0.5)}" x2="${x(history.length - 0.5)}" y1="${margin.top}" y2="${margin.top + plotH}"></line>` : ""}
      ${xLabels.map(({ row, index }) => `<text class="svg-label" x="${x(index)}" y="${height - 17}" text-anchor="middle">${esc(monthLabel(row.month).replace(" ", " '"))}</text>`).join("")}
      ${history.map((row, index) => `<circle class="svg-point chart-hit" cx="${x(index)}" cy="${y(row.value)}" r="3" data-tip="${esc(monthLabel(row.month, false))}: actual ${fmt(row.value)}"></circle>`).join("")}
      ${future.map((row, index) => `<circle class="svg-point forecast chart-hit" cx="${x(history.length + index)}" cy="${y(row.point)}" r="3" data-tip="${esc(monthLabel(row.month, false))}: forecast ${fmt(row.point)} | 95% ${fmt(row.lower95)}-${fmt(row.upper95)}"></circle>`).join("")}`;
    const tooltip = el("forecast-tooltip");
    svg.querySelectorAll(".chart-hit").forEach((point) => {
      point.addEventListener("pointerenter", (event) => {
        tooltip.textContent = point.dataset.tip;
        tooltip.style.display = "block";
        tooltip.style.left = `${Math.min(event.offsetX + 10, svg.clientWidth - 190)}px`;
        tooltip.style.top = `${Math.max(4, event.offsetY - 28)}px`;
      });
      point.addEventListener("pointerleave", () => { tooltip.style.display = "none"; });
    });
  }

  function renderAdjustments(sku, bundle) {
    const position = sheet("DDMRP_Positions").find((row) => row.SKU === sku);
    const baselineAdu = num(position?.ADU);
    const dlt = num(position?.DLT_Days);
    el("adjustment-body").innerHTML = bundle.forecast.length ? bundle.forecast.map((row) => {
      const forecastAdu = row.point / daysInMonth(row.month);
      const factor = baselineAdu ? forecastAdu / baselineAdu : 1;
      return `<tr><td>${esc(monthLabel(row.month, false))}</td><td class="number">${fmt(row.point)}</td><td class="number">${fmt(row.lower95)}-${fmt(row.upper95)}</td><td class="number">${fmt(forecastAdu, 2)}</td><td class="number">${baselineAdu ? fmt(baselineAdu, 2) : "-"}</td><td class="number">${baselineAdu ? fmt(factor, 2) : "-"}</td><td>${dlt ? esc(addDays(row.month, -dlt)) : "-"}</td><td>Proposed</td></tr>`;
    }).join("") : '<tr><td colspan="8" class="empty-state">No forecast is available.</td></tr>';
  }

  function filteredSupply() {
    const products = productLookup();
    const plan = new Map(allPositions().map((row) => [row.key, row]));
    const query = state.search.trim().toLowerCase();
    return sheet("Open_Supply").map((row) => {
      const product = products.get(row.SKU) || {};
      const position = plan.get(keyFor(row));
      const due = dateValue(row.Expected_Receipt_Date);
      const snap = dateValue(snapshotDate());
      const days = due && snap ? Math.round((due - snap) / 86400000) : 9999;
      const closed = /received|closed|cancel/i.test(String(row.Order_Status));
      let risk = "Monitor";
      if (!closed && days < 0) risk = "Late";
      else if (!closed && days <= 7) risk = "Due soon";
      else if (position?.executionRisk) risk = "Inventory risk";
      return { ...row, product, position, days, risk };
    }).filter((row) => {
      if (state.group !== "All" && row.product.Product_Group !== state.group) return false;
      if (state.location !== "All" && row.Location !== state.location) return false;
      if (state.status !== "All" && row.position?.status !== state.status) return false;
      return !query || `${row.SKU} ${row.product.SKU_Description || ""} ${row.Supply_Order_ID}`.toLowerCase().includes(query);
    }).sort((a, b) => ({ Late: 0, "Due soon": 1, "Inventory risk": 2, Monitor: 3 })[a.risk] - ({ Late: 0, "Due soon": 1, "Inventory risk": 2, Monitor: 3 })[b.risk] || a.days - b.days);
  }

  function renderExecution() {
    const supply = filteredSupply();
    const late = supply.filter((row) => row.risk === "Late").length;
    const onhandRisk = filteredPositions().filter((row) => row.executionRisk).length;
    const total = supply.reduce((sum, row) => sum + num(row.Open_Supply_Units), 0);
    el("late-count").textContent = fmt(late);
    el("onhand-risk-count").textContent = fmt(onhandRisk);
    el("open-supply-total").textContent = fmt(total);
    el("execution-headline").textContent = late ? `${late} late ${late === 1 ? "receipt requires" : "receipts require"} intervention.` : "No late receipts in the current filter.";
    el("execution-body").innerHTML = supply.length ? supply.map((row) => `<tr>
      <td>${statusMarkup(row.risk === "Late" ? "Red" : row.risk === "Due soon" || row.risk === "Inventory risk" ? "Yellow" : "Green")}<br><small>${esc(row.risk)}</small></td>
      <td>${esc(row.Supply_Order_ID)}</td><td class="item-cell"><strong>${esc(row.SKU)}</strong><span>${esc(row.product.SKU_Description || "")}</span></td>
      <td>${esc(row.Location)}</td><td>${esc(row.Supply_Type)}</td><td class="number">${fmt(row.Open_Supply_Units)}</td>
      <td class="${row.risk === "Late" ? "late-text" : ""}">${esc(isoDate(row.Expected_Receipt_Date))}</td><td>${esc(row.Order_Status)}</td>
      <td>${esc(row.Supplier_or_Work_Center)}</td><td>${esc(row.Synchronization_Note)}</td>
    </tr>`).join("") : '<tr><td colspan="10" class="empty-state">No open supply lines match the active filters.</td></tr>';
  }

  function renderBuffers() {
    const positions = filteredPositions();
    const masters = new Map(sheet("DDMRP_Master").map((row) => [keyFor(row), row]));
    if (!state.selectedBufferKey || !positions.some((row) => row.key === state.selectedBufferKey)) state.selectedBufferKey = positions[0]?.key || null;
    el("buffer-body").innerHTML = positions.length ? positions.map((row) => {
      const master = masters.get(row.key) || {};
      return `<tr data-selectable data-key="${esc(row.key)}" tabindex="0" class="${row.key === state.selectedBufferKey ? "selected" : ""}">
        <td class="item-cell"><strong>${esc(row.SKU)}</strong><span>${esc(row.Location)}</span></td><td>${esc(master.Buffer_Profile || "-")}</td>
        <td class="number">${fmt(row.ADU, 2)}</td><td class="number">${fmt(row.DLT_Days)}</td><td class="number">${fmt(row.Lead_Time_Factor, 2)}</td>
        <td class="number">${fmt(row.Variability_Factor, 2)}</td><td class="number">${fmt(row.MOQ)}</td><td class="number">${fmt(row.Order_Cycle_Days)}</td>
        <td class="number">${fmt(row.Demand_Adjustment_Factor, 2)}</td><td>${esc(master.Setting_Approval_Status || "Not supplied")}</td>
      </tr>`;
    }).join("") : '<tr><td colspan="10" class="empty-state">No buffer settings match the active filters.</td></tr>';
    document.querySelectorAll("#buffer-body tr[data-selectable]").forEach((row) => {
      const select = () => { state.selectedBufferKey = row.dataset.key; renderBuffers(); };
      row.addEventListener("click", select);
      row.addEventListener("keydown", (event) => { if (event.key === "Enter" || event.key === " ") select(); });
    });
    const selected = positions.find((row) => row.key === state.selectedBufferKey);
    renderBufferDetail(selected, selected ? masters.get(selected.key) || {} : {});
  }

  function renderBufferDetail(row, master) {
    if (!row) {
      el("buffer-detail").innerHTML = '<p class="empty-state">Select a buffer to review its settings.</p>';
      return;
    }
    el("buffer-detail").innerHTML = `<header class="detail-header">
      <p class="section-kicker">Buffer definition</p><div class="detail-title-row"><h2>${esc(row.SKU)} / ${esc(row.Location)}</h2><span>${esc(master.Buffer_Profile || "-")}</span></div>
      <p>${esc(row.SKU_Description)}</p>
    </header>
    <section class="detail-section"><h3>Current scenario</h3>${bufferMarkup(row, true)}
      <p>${state.scenario === "forecast" ? "Forecast-derived proposal shown; approved master data remains unchanged." : "Approved demand adjustment and master settings shown."}</p>
      <dl class="formula-list">
        <div class="formula-row"><dt>Red base</dt><dd>${fmt(row.redBase)}</dd></div><div class="formula-row"><dt>Red safety</dt><dd>${fmt(row.redSafety)}</dd></div>
        <div class="formula-row"><dt>Yellow zone</dt><dd>${fmt(row.yellow)}</dd></div><div class="formula-row"><dt>Green zone</dt><dd>${fmt(row.green)}</dd></div>
      </dl>
    </section>
    <section class="detail-section"><h3>Policy evidence</h3><dl class="formula-list">
      <div class="formula-row"><dt>Supply type</dt><dd>${esc(master.Supply_Type || "-")}</dd></div>
      <div class="formula-row"><dt>ADU method</dt><dd>${esc(master.ADU_Calculation_Method || "-")}</dd></div>
      <div class="formula-row"><dt>Spike horizon</dt><dd>${fmt(master.Spike_Horizon_Days)} days</dd></div>
      <div class="formula-row"><dt>Spike threshold</dt><dd>${fmt(master.Spike_Threshold_Units)}</dd></div>
      <div class="formula-row"><dt>Order multiple</dt><dd>${fmt(row.Order_Multiple)}</dd></div>
    </dl></section>
    <section class="detail-section"><h3>Governance</h3><dl class="formula-list">
      <div class="formula-row"><dt>Planner owner</dt><dd>${esc(master.Planner_Owner || "-")}</dd></div>
      <div class="formula-row"><dt>Approval</dt><dd>${esc(master.Setting_Approval_Status || "-")}</dd></div>
      <div class="formula-row"><dt>Next review</dt><dd>${esc(isoDate(master.Next_Review_Date))}</dd></div>
    </dl></section>`;
  }

  function renderData() {
    const sheets = state.data.sheets || {};
    const requiredMissing = REQUIRED_SHEETS.filter((name) => !(sheets[name] || []).length);
    const forecastSkus = new Set(sheet("Forecast_Results").map((row) => row.SKU));
    el("demand-record-count").textContent = fmt(sheet("Demand_Data").length);
    el("position-count").textContent = fmt(sheet("DDMRP_Positions").length);
    el("forecast-coverage").textContent = fmt(forecastSkus.size || new Set(sheet("Demand_Data").map((row) => row.SKU)).size);
    el("data-headline").textContent = requiredMissing.length ? `Workbook is missing ${requiredMissing.join(" and ")}.` : "Workbook is complete for forecasting and DDMRP.";
    el("data-subline").textContent = state.sourceMode === "fpp3" ? "Packaged R/fpp3 results and planning source classes are auditable." : "Forecasts are calculated in this browser session from loaded demand history.";
    el("sheet-body").innerHTML = Object.entries(SHEET_PURPOSES).map(([name, purpose]) => {
      const count = (sheets[name] || []).length;
      const required = REQUIRED_SHEETS.includes(name);
      const status = count ? "Loaded" : required ? "Missing" : "Optional / absent";
      return `<tr><td>${esc(name)}</td><td>${esc(purpose)}</td><td class="number">${fmt(count)}</td><td class="${count ? "sheet-ok" : required ? "sheet-missing" : ""}">${status}</td></tr>`;
    }).join("");
    renderAudit();
  }

  function renderAudit() {
    el("audit-body").innerHTML = state.actionLog.length ? [...state.actionLog].reverse().map((row) => `<tr><td>${esc(row.time)}</td><td>${esc(row.sku)}</td><td>${esc(row.action)}</td><td class="number">${fmt(row.quantity)}</td><td>${esc(row.reason)}</td></tr>`).join("") : '<tr><td colspan="5" class="empty-state">No planner decisions recorded in this session.</td></tr>';
  }

  function populateFilters() {
    const products = productLookup();
    const positions = sheet("DDMRP_Positions");
    const groups = [...new Set([...positions.map((row) => row.Product_Group || products.get(row.SKU)?.Product_Group), ...sheet("Demand_Data").map((row) => row.Product_Group)].filter(Boolean))].sort();
    const locations = [...new Set(positions.map((row) => row.Location).filter(Boolean))].sort();
    el("group-filter").innerHTML = '<option value="All">All groups</option>' + groups.map((value) => `<option value="${esc(value)}">${esc(value)}</option>`).join("");
    el("location-filter").innerHTML = '<option value="All">All locations</option>' + locations.map((value) => `<option value="${esc(value)}">${esc(value)}</option>`).join("");
    el("group-filter").value = groups.includes(state.group) ? state.group : "All";
    el("location-filter").value = locations.includes(state.location) ? state.location : "All";
    el("status-filter").value = state.status;
    const eligibleSkus = [...new Set(sheet("Demand_Data").map((row) => row.SKU).filter(Boolean))].filter((sku) => state.group === "All" || products.get(sku)?.Product_Group === state.group).sort();
    if (!state.selectedSku || !eligibleSkus.includes(state.selectedSku)) state.selectedSku = eligibleSkus[0] || null;
    el("forecast-sku").innerHTML = eligibleSkus.map((sku) => `<option value="${esc(sku)}">${esc(sku)} | ${esc(products.get(sku)?.SKU_Description || "")}</option>`).join("");
    el("forecast-sku").value = state.selectedSku || "";
  }

  function renderAll() {
    populateFilters();
    renderPlanning();
    renderForecast();
    renderExecution();
    renderBuffers();
    renderData();
    updateDataAlert();
  }

  function updateDataAlert(message = "") {
    const missing = REQUIRED_SHEETS.filter((name) => !sheet(name).length);
    const alert = el("data-alert");
    const text = message || (missing.length ? `Incomplete workbook: ${missing.join(", ")} ${missing.length === 1 ? "is" : "are"} required for the full planning workflow.` : "");
    alert.textContent = text;
    alert.hidden = !text;
  }

  function switchView(view) {
    state.view = view;
    document.querySelectorAll(".tab").forEach((tab) => {
      const active = tab.dataset.view === view;
      tab.classList.toggle("active", active);
      tab.setAttribute("aria-selected", String(active));
    });
    document.querySelectorAll(".view").forEach((panel) => {
      const active = panel.id === `view-${view}`;
      panel.classList.toggle("active", active);
      panel.hidden = !active;
    });
    if (view === "forecast") requestAnimationFrame(renderForecast);
  }

  function openActionDialog(row, action) {
    state.pendingAction = { row, action };
    el("dialog-title").textContent = `${action} ${row.SKU} recommendation`;
    el("dialog-summary").innerHTML = `<strong>${esc(row.SKU)} / ${esc(row.Location)}</strong><br>${action === "Release" ? `${fmt(row.recommended)} units, suggested availability ${esc(row.availability)}` : "Recommendation held for planner review"}`;
    el("action-reason").value = "";
    el("action-dialog").showModal();
  }

  function confirmAction(event) {
    event.preventDefault();
    const reason = el("action-reason").value.trim();
    if (!reason) {
      el("action-reason").focus();
      return;
    }
    const { row, action } = state.pendingAction;
    const entry = {
      time: new Date().toLocaleString(),
      sku: row.SKU,
      action,
      quantity: action === "Release" ? row.recommended : 0,
      reason,
    };
    state.actionLog.push(entry);
    state.decisions[row.key] = entry;
    el("action-dialog").close();
    showToast(`${row.SKU}: ${action.toLowerCase()} decision recorded in the session log.`);
    renderPlanning();
    renderAudit();
  }

  let toastTimer;
  function showToast(message) {
    clearTimeout(toastTimer);
    el("toast").textContent = message;
    el("toast").classList.add("visible");
    toastTimer = setTimeout(() => el("toast").classList.remove("visible"), 3200);
  }

  function csvCell(value) {
    const text = String(value ?? "");
    return /[",\n]/.test(text) ? `"${text.replaceAll('"', '""')}"` : text;
  }

  function downloadCsv(filename, rows) {
    if (!rows.length) return showToast("There is no data to export for the active selection.");
    const headers = Object.keys(rows[0]);
    const csv = [headers, ...rows.map((row) => headers.map((header) => row[header]))].map((row) => row.map(csvCell).join(",")).join("\n");
    const url = URL.createObjectURL(new Blob([csv], { type: "text/csv;charset=utf-8" }));
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = filename;
    anchor.click();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  }

  function exportPlan() {
    downloadCsv("demand-genie-v2-recommendations.csv", filteredPositions().map((row) => ({
      SKU: row.SKU,
      Location: row.Location,
      Scenario: state.scenario,
      NFP_Status: row.status,
      Net_Flow_Position: round(row.nfp, 2),
      TOR: round(row.tor, 2),
      TOY: round(row.toy, 2),
      TOG: round(row.tog, 2),
      Qualified_Demand: round(row.qualifiedDemand, 2),
      Recommended_Order_Units: round(row.recommended, 0),
      Suggested_Availability_Date: row.availability,
      Planner_Decision: decisionLabel(row),
    })));
  }

  function exportForecast() {
    const bundle = getForecastBundle(state.selectedSku);
    downloadCsv(`demand-genie-v2-forecast-${state.selectedSku}.csv`, bundle.forecast.map((row) => ({
      SKU: state.selectedSku,
      Month: row.month,
      Model: bundle.selected,
      Forecast_Units: round(row.point, 2),
      PI80_Lower: round(row.lower80, 2),
      PI80_Upper: round(row.upper80, 2),
      PI95_Lower: round(row.lower95, 2),
      PI95_Upper: round(row.upper95, 2),
      Engine: bundle.source === "fpp3" ? state.data.forecastEngine : "Browser fallback",
    })));
  }

  async function loadWorkbook(file) {
    try {
      updateDataAlert(`Loading ${file.name}...`);
      const bytes = await file.arrayBuffer();
      const workbook = XLSX.read(bytes, { type: "array", cellDates: true });
      const sheets = {};
      Object.keys(SHEET_PURPOSES).forEach((name) => {
        if (!workbook.SheetNames.includes(name)) return;
        sheets[name] = normalizeRows(XLSX.utils.sheet_to_json(workbook.Sheets[name], { defval: null, raw: true }));
      });
      if (file.name === SAMPLE_DATA.fileName && sheets.Demand_Data?.length === SAMPLE_DATA.sheets.Demand_Data?.length) {
        ["Forecast_Results", "Model_Selection", "Forecast_Exceptions", "Forecast_Run_Metadata"].forEach((name) => { sheets[name] = cloneData(SAMPLE_DATA.sheets[name] || []); });
        state.sourceMode = "fpp3";
      } else {
        state.sourceMode = sheets.Forecast_Results?.length && sheets.Model_Selection?.length ? "fpp3" : "browser";
      }
      state.data = {
        fileName: file.name,
        sourceType: "Uploaded workbook",
        forecastEngine: state.sourceMode === "fpp3" ? "Packaged R fpp3 output" : "Browser fallback",
        sheets,
      };
      state.group = "All";
      state.location = "All";
      state.status = "All";
      state.search = "";
      state.selectedKey = null;
      state.selectedSku = null;
      state.selectedBufferKey = null;
      el("group-filter").value = "All";
      el("location-filter").value = "All";
      el("status-filter").value = "All";
      el("search-filter").value = "";
      el("source-kind").textContent = "Uploaded workbook";
      el("source-file").textContent = file.name;
      el("source-freshness").textContent = `Snapshot ${snapshotDate()}`;
      renderAll();
      const missing = REQUIRED_SHEETS.filter((name) => !sheet(name).length);
      updateDataAlert(missing.length ? "" : state.sourceMode === "browser" ? "Forecasts use a browser holdout comparison because packaged FPP3 outputs were not present." : "");
      showToast(`${file.name} loaded: ${fmt(sheet("Demand_Data").length)} demand rows and ${fmt(sheet("DDMRP_Positions").length)} positions.`);
    } catch (error) {
      console.error(error);
      updateDataAlert(`Could not load ${file.name}. Confirm that it is an unencrypted Excel workbook with named header rows.`);
    } finally {
      el("workbook-input").value = "";
    }
  }

  function bindEvents() {
    document.querySelectorAll(".tab").forEach((tab) => tab.addEventListener("click", () => switchView(tab.dataset.view)));
    el("upload-button").addEventListener("click", () => el("workbook-input").click());
    el("workbook-input").addEventListener("change", (event) => { if (event.target.files[0]) loadWorkbook(event.target.files[0]); });
    el("group-filter").addEventListener("change", (event) => { state.group = event.target.value; renderAll(); });
    el("location-filter").addEventListener("change", (event) => { state.location = event.target.value; renderAll(); });
    el("status-filter").addEventListener("change", (event) => { state.status = event.target.value; renderAll(); });
    el("search-filter").addEventListener("input", (event) => { state.search = event.target.value; renderPlanning(); renderExecution(); renderBuffers(); });
    document.querySelectorAll("#scenario-control button").forEach((button) => button.addEventListener("click", () => {
      state.scenario = button.dataset.scenario;
      document.querySelectorAll("#scenario-control button").forEach((item) => {
        const active = item === button;
        item.classList.toggle("active", active);
        item.setAttribute("aria-pressed", String(active));
      });
      renderPlanning();
      renderBuffers();
    }));
    el("forecast-sku").addEventListener("change", (event) => { state.selectedSku = event.target.value; renderForecast(); });
    el("forecast-model").addEventListener("change", (event) => { state.forecastModel = event.target.value; renderForecast(); });
    document.querySelectorAll("#horizon-control button").forEach((button) => button.addEventListener("click", () => {
      state.horizon = Number(button.dataset.horizon);
      document.querySelectorAll("#horizon-control button").forEach((item) => {
        const active = item === button;
        item.classList.toggle("active", active);
        item.setAttribute("aria-pressed", String(active));
      });
      renderForecast();
    }));
    el("export-plan").addEventListener("click", exportPlan);
    el("export-forecast").addEventListener("click", exportForecast);
    el("export-audit").addEventListener("click", () => downloadCsv("demand-genie-v2-planner-action-log.csv", state.actionLog));
    el("dialog-cancel").addEventListener("click", () => el("action-dialog").close());
    el("action-form").addEventListener("submit", confirmAction);
  }

  function initialize() {
    state.sourceMode = sheet("Forecast_Results").length ? "fpp3" : "browser";
    bindEvents();
    renderAll();
    el("source-kind").textContent = state.data.sourceType;
    el("source-file").textContent = state.data.fileName;
    el("source-freshness").textContent = `Snapshot ${snapshotDate()}`;
  }

  initialize();
})();
