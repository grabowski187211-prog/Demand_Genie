(() => {
  "use strict";

  const REQUIRED_SHEETS = ["Demand_Data", "DDMRP_Positions", "Spend_Lines"];
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
    All_Model_Forecasts: "Candidate forecasts refit on all history",
    Model_Selection: "Rolling-origin model accuracy and selection",
    Model_Summary: "Cross-SKU candidate scorecard",
    Forecast_Exceptions: "Forecast review and calibration flags",
    Forecast_Run_Metadata: "Versioned engine, model, and input hashes",
    Decomposition: "Robust STL components by SKU and month",
    Decomposition_Features: "Trend and seasonal strength diagnostics",
    Spend_Lines: "Signed procurement transaction ledger",
    Supplier_Master: "Normalized supplier and commercial-risk master",
    Contracts: "Contract coverage and renewal register",
    Category_Taxonomy: "Approved procurement classification",
    Category_Risk: "Kraljic impact and risk evidence",
    Spend_Control: "Finance control-total reconciliation",
    Demand_History: "Consumption fact for item segmentation",
    Spend_Summary: "Reconciled spend analysis summary",
    Supplier_Summary: "Supplier-parent Pareto analysis",
    Category_Summary: "Category Pareto, concentration, and Kraljic output",
    Monthly_Spend: "Monthly net, gross, and credit trend",
    Opportunity_Flags: "Evidence-based commercial review prompts",
    Spend_Data_Quality: "Spend and demand quality gates",
    Item_Segmentation: "ABC/XYZ and ADI/CV2 item behavior",
  };
  const STATUS_ORDER = { Data: -1, Black: 0, Red: 1, Yellow: 2, Green: 3, Blue: 4 };
  const POSITION_NUMERIC_FIELDS = [
    "ADU", "DLT_Days", "Lead_Time_Factor", "Variability_Factor", "MOQ", "Order_Cycle_Days",
    "On_Hand", "Open_Supply", "Past_Due_Demand", "Today_Demand", "Qualified_Spike_Demand",
    "Order_Multiple", "Demand_Adjustment_Factor",
  ];
  const MODEL_LABELS = {
    Auto: "Automatic winner",
    Mean: "Mean",
    Naive: "Naive",
    Drift: "Drift",
    Seasonal_Naive: "Seasonal naive",
    ETS: "ETS",
    ARIMA: "ARIMA",
    TiRex2: "TiRex2",
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
    forecastMode: "forecast",
    actionLog: [],
    decisions: {},
    pendingAction: null,
    sourceMode: "packaged",
    spendPeriod: "All",
    category: "All",
    supplier: "All",
    contract: "All",
    commercialSearch: "",
    paretoDimension: "category",
    segmentationMode: "behavior",
    selectedCategory: null,
  };

  const el = (id) => document.getElementById(id);
  const num = (value, fallback = 0) => {
    const parsed = Number(value);
    return Number.isFinite(parsed) ? parsed : fallback;
  };
  const nullableNum = (value) => {
    if (value === null || value === undefined || value === "") return null;
    const parsed = Number(value);
    return Number.isFinite(parsed) ? parsed : null;
  };
  const hasNumber = (value) => value !== null && value !== "" && Number.isFinite(Number(value));
  const truthValue = (value) => value === true || ["true", "yes", "1"].includes(String(value ?? "").toLowerCase());
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
  const money = (value, compact = true) => {
    const amount = num(value);
    if (!compact) return new Intl.NumberFormat("en-US", { style: "currency", currency: "EUR", maximumFractionDigits: 0 }).format(amount);
    const absolute = Math.abs(amount);
    const sign = amount < 0 ? "-" : "";
    if (absolute >= 1000000) return `${sign}EUR ${fmt(absolute / 1000000, absolute >= 10000000 ? 1 : 2)}m`;
    if (absolute >= 1000) return `${sign}EUR ${fmt(absolute / 1000, 1)}k`;
    return `${sign}EUR ${fmt(absolute, 0)}`;
  };
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

  function positionValidation(position) {
    const missing = POSITION_NUMERIC_FIELDS.filter((field) => !hasNumber(position[field]));
    const invalid = [];
    if (hasNumber(position.ADU) && num(position.ADU) < 0) invalid.push("ADU");
    if (hasNumber(position.DLT_Days) && num(position.DLT_Days) <= 0) invalid.push("DLT_Days");
    ["Lead_Time_Factor", "Variability_Factor", "MOQ", "On_Hand", "Open_Supply", "Past_Due_Demand", "Today_Demand", "Qualified_Spike_Demand"].forEach((field) => {
      if (hasNumber(position[field]) && num(position[field]) < 0) invalid.push(field);
    });
    if (hasNumber(position.Order_Cycle_Days) && num(position.Order_Cycle_Days) <= 0) invalid.push("Order_Cycle_Days");
    if (hasNumber(position.Order_Multiple) && num(position.Order_Multiple) <= 0) invalid.push("Order_Multiple");
    if (hasNumber(position.Demand_Adjustment_Factor) && num(position.Demand_Adjustment_Factor) <= 0) invalid.push("Demand_Adjustment_Factor");
    return [...new Set([...missing, ...invalid])];
  }

  function calculatePosition(position, scenario = state.scenario) {
    const invalidFields = positionValidation(position);
    if (invalidFields.length) {
      return {
        ...position,
        calculationValid: false,
        invalidFields,
        factor: null,
        adjustedAdu: null,
        yellow: null,
        redBase: null,
        redSafety: null,
        red: null,
        green: null,
        tor: null,
        toy: null,
        tog: null,
        qualifiedDemand: null,
        nfp: null,
        nfpPct: -1,
        status: "Data",
        triggered: false,
        recommended: 0,
        availability: "Blocked",
        executionRisk: false,
        key: keyFor(position),
      };
    }
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
      calculationValid: true,
      invalidFields: [],
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

  function allPositions(scenario = state.scenario) {
    const products = productLookup();
    return sheet("DDMRP_Positions").map((row) => {
      const product = products.get(row.SKU) || {};
      return calculatePosition({
        ...row,
        Product_Group: row.Product_Group || product.Product_Group || "Unassigned",
        SKU_Description: row.SKU_Description || product.SKU_Description || row.SKU,
      }, scenario);
    });
  }

  function filteredPositions(scenario = state.scenario) {
    const query = state.search.trim().toLowerCase();
    return allPositions(scenario).filter((row) => {
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
    if (!row.calculationValid) return `<p class="invalid-calculation">Calculation blocked: ${esc(row.invalidFields.join(", "))}</p>`;
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
    const invalid = positions.filter((row) => !row.calculationValid);
    const proposals = positions.filter((row) => row.calculationValid && row.triggered).length;
    const risks = positions.filter((row) => row.executionRisk).length;
    el("critical-count").textContent = fmt(critical.length);
    el("critical-context").textContent = `of ${positions.length} positions`;
    el("release-quantity").textContent = fmt(proposals);
    el("release-context").textContent = `${invalid.length ? `${invalid.length} blocked; ` : ""}triggered positions`;
    el("execution-risk-count").textContent = fmt(risks);
    el("planning-headline").textContent = invalid.length
      ? `${invalid.length} ${invalid.length === 1 ? "position is" : "positions are"} blocked by invalid planning inputs.`
      : critical.length
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
    if (!row.calculationValid) {
      el("planning-detail").innerHTML = `<header class="detail-header"><p class="section-kicker">Planning data gate</p><div class="detail-title-row"><h2>${esc(row.SKU)} / ${esc(row.Location)}</h2>${statusMarkup("Data")}</div><p>${esc(row.SKU_Description)}</p></header><section class="detail-section"><h3>Calculation blocked</h3><p class="detail-callout critical">Correct these fields before using a buffer status or order recommendation: ${esc(row.invalidFields.join(", "))}.</p></section>`;
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
    const supported = new Set(["Mean", "Naive", "Drift", "Seasonal_Naive", "ETS"]);
    const selected = requestedModel === "Auto" || !supported.has(requestedModel) ? comparisons[0]?.model || "Naive" : requestedModel;
    const values = history.map((row) => row.value);
    const points = modelForecast(values, selected, horizon);
    const validation = comparisons.find((row) => row.model === selected) || comparisons[0];
    const lastMonth = history.at(-1)?.month || new Date().toISOString();
    const forecast = points.map((point, index) => {
      return {
        month: addMonths(lastMonth, index + 1),
        point: Math.max(0, point),
        lower80: null,
        upper80: null,
        lower95: null,
        upper95: null,
      };
    });
    return { history, forecast, comparisons, selected, metrics: validation, source: "browser", has80: false, has95: false };
  }

  function getForecastBundle(sku, requestedModel = state.forecastModel, horizon = state.horizon) {
    const history = historyFor(sku);
    const selectionRows = sheet("Model_Selection").filter((row) => row.SKU === sku);
    const selectedRow = fppSelection(sku);
    const selectedModel = selectedRow?.Selected_Model || selectedRow?.Model;
    const effectiveModel = requestedModel === "Auto" ? selectedModel : requestedModel;
    const packaged = sheet("All_Model_Forecasts")
      .filter((row) => row.SKU === sku && row.Model === effectiveModel)
      .sort((a, b) => isoDate(a.Month).localeCompare(isoDate(b.Month)));
    if (state.sourceMode === "packaged" && effectiveModel && packaged.length) {
      const comparisons = selectionRows.map((row) => ({
        model: row.Model || row[".model"] || row.Selected_Model,
        ME: num(row.ME, NaN),
        RMSE: num(row.RMSE, NaN),
        MAE: num(row.MAE, NaN),
        MASE: num(row.MASE, NaN),
        RMSSE: num(row.RMSSE, NaN),
        coverage80: num(row.Coverage_80, NaN),
        selected: truthValue(row.Selected),
        nearTie: truthValue(row.Near_Tie),
        rawBest: row.Raw_Best_Model,
        rank: num(row.RMSE_Rank, NaN),
        rationale: row.Selection_Rationale,
        origins: num(row.Evaluation_Origins, NaN),
        points: num(row.Evaluation_Points, NaN),
      })).filter((row) => row.model).sort((a, b) => a.RMSE - b.RMSE);
      const forecast = packaged.slice(0, horizon).map((row) => ({
        month: isoDate(row.Month),
        point: num(row.Forecast_Units),
        lower80: nullableNum(row.PI80_Lower_Units),
        upper80: nullableNum(row.PI80_Upper_Units),
        lower95: nullableNum(row.PI95_Lower_Units),
        upper95: nullableNum(row.PI95_Upper_Units),
      }));
      return {
        history,
        forecast,
        comparisons,
        selected: effectiveModel,
        metrics: comparisons.find((row) => row.model === effectiveModel) || comparisons[0],
        source: "packaged",
        intervalMethod: packaged[0]?.Interval_Method || "Packaged model interval",
        has80: forecast.every((row) => row.lower80 !== null && row.upper80 !== null),
        has95: forecast.every((row) => row.lower95 !== null && row.upper95 !== null),
      };
    }
    return browserForecast(sku, requestedModel, horizon);
  }

  function renderForecast() {
    const sku = state.selectedSku;
    if (!sku) return;
    const decompositionMode = state.forecastMode === "decomposition";
    el("forecast-results").hidden = decompositionMode;
    el("decomposition-results").hidden = !decompositionMode;
    if (decompositionMode) {
      renderDecomposition(sku);
      return;
    }
    const product = productLookup().get(sku) || {};
    const bundle = getForecastBundle(sku);
    el("forecast-source-note").textContent = bundle.source === "packaged"
      ? `${state.data.forecastProtocol || "7 rolling origins x 6 months"} | display ${state.horizon} months`
      : "Browser benchmarks only | six-month holdout | no probabilistic intervals";
    el("engine-state").textContent = `Forecast: ${bundle.source === "packaged" ? state.data.forecastEngine : "local benchmark fallback"} | Spend: ${state.data.spendEngine || "local auditable analysis"}`;
    el("selected-model").textContent = MODEL_LABELS[bundle.selected] || bundle.selected;
    el("model-context").textContent = state.forecastModel === "Auto" ? "7 origins / 42 errors" : "Planner-selected comparison";
    el("forecast-rmse").textContent = fmt(bundle.metrics?.RMSE, 1);
    el("forecast-bias").textContent = fmt(bundle.metrics?.ME, 1);
    el("forecast-headline").textContent = `${sku}: ${MODEL_LABELS[bundle.selected] || bundle.selected} evaluated at six months; showing ${state.horizon}.`;
    el("forecast-subline").textContent = bundle.source === "packaged"
      ? "Every candidate used identical rolling origins; forecast output remains separate from qualified demand."
      : "Only transparent browser benchmarks are available; ARIMA, TiRex2, and defensible intervals require an offline run.";
    el("forecast-chart-title").textContent = `${sku} | ${product.SKU_Description || "Demand history"}`;
    el("forecast-interval-legend").textContent = bundle.has95 ? "80% / 95%" : bundle.has80 ? "native 80%" : "interval unavailable";
    const notes = [];
    if (bundle.metrics?.nearTie) notes.push("Near tie: at least two candidates are within 2% RMSE.");
    if (Number.isFinite(bundle.metrics?.coverage80)) notes.push(`Empirical 80% coverage ${pct(bundle.metrics.coverage80)}.`);
    if (bundle.selected === "TiRex2") notes.push("TiRex2 uses q50 and a native marginal q10-q90 interval; no native 95% interval is claimed.");
    const review = el("forecast-review-note");
    review.textContent = notes.join(" ");
    review.hidden = !notes.length;
    const portfolio = sheet("Model_Summary").find((row) => row.Model === bundle.selected);
    el("model-portfolio-note").textContent = portfolio
      ? `${fmt(portfolio.Selected_SKUs)} of 80 SKUs selected | macro MASE ${fmt(portfolio.Macro_MASE, 2)} | portfolio 80% coverage ${pct(portfolio.Empirical_Coverage_80)}`
      : "Every candidate uses the same out-of-sample observations.";
    renderForecastChart(bundle);
    renderModelComparison(bundle);
    renderAdjustments(sku, bundle);
  }

  function renderModelComparison(bundle) {
    const rows = bundle.comparisons.filter((row) => Number.isFinite(row.RMSE));
    const maxRmse = Math.max(...rows.map((row) => row.RMSE), 1);
    el("model-comparison").innerHTML = rows.length ? rows.map((row) => `<div class="model-row ${row.model === bundle.selected ? "selected" : ""}">
      <div class="model-name"><span>${esc(MODEL_LABELS[row.model] || row.model)}${row.model === bundle.selected ? " | selected" : row.model === row.rawBest ? " | lowest RMSE" : ""}</span><span>RMSE ${fmt(row.RMSE, 1)}</span></div>
      <div class="model-meter" style="--meter:${clamp((row.RMSE / maxRmse) * 100, 2, 100)}%"><span></span></div>
      <div class="model-metrics"><span>MAE ${fmt(row.MAE, 1)}</span><span>MASE ${fmt(row.MASE, 2)}</span><span>80% ${Number.isFinite(row.coverage80) ? pct(row.coverage80) : "n/a"}</span></div>
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
    const outerLow = (row) => bundle.has95 ? row.lower95 : bundle.has80 ? row.lower80 : row.point;
    const outerHigh = (row) => bundle.has95 ? row.upper95 : bundle.has80 ? row.upper80 : row.point;
    const minValue = Math.min(0, ...future.map(outerLow));
    const maxValue = Math.max(...history.map((row) => row.value), ...future.map(outerHigh), 1) * 1.05;
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
      ${future.length && bundle.has95 ? `<path class="svg-pi95" d="${areaPath(connectedForecast, forecastOffset, (row) => row.lower95, (row) => row.upper95)}"></path>` : ""}
      ${future.length && bundle.has80 ? `<path class="svg-pi80" d="${areaPath(connectedForecast, forecastOffset, (row) => row.lower80, (row) => row.upper80)}"></path>` : ""}
      <line class="svg-axis" x1="${margin.left}" x2="${width - margin.right}" y1="${margin.top + plotH}" y2="${margin.top + plotH}"></line>
      ${history.length ? `<path class="svg-actual" d="${actualPath}"></path>` : ""}
      ${connectedForecast.length ? `<path class="svg-forecast" d="${linePath(connectedForecast, forecastOffset, (row) => row.point)}"></path>` : ""}
      ${future.length && history.length ? `<line class="svg-boundary" x1="${x(history.length - 0.5)}" x2="${x(history.length - 0.5)}" y1="${margin.top}" y2="${margin.top + plotH}"></line>` : ""}
      ${xLabels.map(({ row, index }) => `<text class="svg-label" x="${x(index)}" y="${height - 17}" text-anchor="middle">${esc(monthLabel(row.month).replace(" ", " '"))}</text>`).join("")}
      ${history.map((row, index) => `<circle class="svg-point chart-hit" tabindex="0" cx="${x(index)}" cy="${y(row.value)}" r="3" data-tip="${esc(monthLabel(row.month, false))}: actual ${fmt(row.value)}"></circle>`).join("")}
      ${future.map((row, index) => `<circle class="svg-point forecast chart-hit" tabindex="0" cx="${x(history.length + index)}" cy="${y(row.point)}" r="3" data-tip="${esc(monthLabel(row.month, false))}: forecast ${fmt(row.point)}${bundle.has80 ? ` | 80% ${fmt(row.lower80)}-${fmt(row.upper80)}` : ""}"></circle>`).join("")}`;
    const tooltip = el("forecast-tooltip");
    svg.querySelectorAll(".chart-hit").forEach((point) => {
      const show = (event) => {
        tooltip.textContent = point.dataset.tip;
        tooltip.style.display = "block";
        const box = point.getBoundingClientRect();
        const parent = svg.getBoundingClientRect();
        tooltip.style.left = `${Math.min(box.left - parent.left + 10, svg.clientWidth - 190)}px`;
        tooltip.style.top = `${Math.max(4, box.top - parent.top - 28)}px`;
      };
      const hide = () => { tooltip.style.display = "none"; };
      point.addEventListener("pointerenter", show);
      point.addEventListener("pointerleave", hide);
      point.addEventListener("focus", show);
      point.addEventListener("blur", hide);
    });
  }

  function renderAdjustments(sku, bundle) {
    const position = sheet("DDMRP_Positions").find((row) => row.SKU === sku);
    const baselineAdu = num(position?.ADU);
    const dlt = num(position?.DLT_Days);
    el("adjustment-interval-label").textContent = bundle.has95 ? "95% interval" : bundle.has80 ? "Native 80% interval" : "Interval";
    el("adjustment-body").innerHTML = bundle.forecast.length ? bundle.forecast.map((row) => {
      const forecastAdu = row.point / daysInMonth(row.month);
      const factor = baselineAdu ? forecastAdu / baselineAdu : 1;
      const interval = bundle.has95 ? `${fmt(row.lower95)}-${fmt(row.upper95)}` : bundle.has80 ? `${fmt(row.lower80)}-${fmt(row.upper80)}` : "Not available";
      return `<tr><td>${esc(monthLabel(row.month, false))}</td><td class="number">${fmt(row.point)}</td><td class="number">${interval}</td><td class="number">${fmt(forecastAdu, 2)}</td><td class="number">${baselineAdu ? fmt(baselineAdu, 2) : "-"}</td><td class="number">${baselineAdu ? fmt(factor, 2) : "-"}</td><td>${dlt ? esc(addDays(row.month, -dlt)) : "-"}</td><td>Proposed</td></tr>`;
    }).join("") : '<tr><td colspan="8" class="empty-state">No forecast is available.</td></tr>';
  }

  function browserDecomposition(sku) {
    const history = historyFor(sku);
    const values = history.map((row) => row.value);
    const n = values.length;
    const meanX = (n - 1) / 2;
    const meanY = values.reduce((sum, value) => sum + value, 0) / Math.max(1, n);
    const slopeDenominator = values.reduce((sum, _value, index) => sum + (index - meanX) ** 2, 0);
    const slope = slopeDenominator ? values.reduce((sum, value, index) => sum + (index - meanX) * (value - meanY), 0) / slopeDenominator : 0;
    const trend = values.map((_value, index) => meanY + slope * (index - meanX));
    const detrended = values.map((value, index) => value - trend[index]);
    const seasonalPattern = Array.from({ length: 12 }, (_, month) => {
      const members = detrended.filter((_value, index) => index % 12 === month);
      return members.length ? members.reduce((sum, value) => sum + value, 0) / members.length : 0;
    });
    const seasonalCenter = seasonalPattern.reduce((sum, value) => sum + value, 0) / 12;
    const seasonal = values.map((_value, index) => seasonalPattern[index % 12] - seasonalCenter);
    return history.map((row, index) => ({ month: row.month, observed: row.value, trend: trend[index], seasonal: seasonal[index], remainder: row.value - trend[index] - seasonal[index] }));
  }

  function componentStrength(signal, remainder) {
    const variance = (values) => {
      const mean = values.reduce((sum, value) => sum + value, 0) / Math.max(1, values.length);
      return values.reduce((sum, value) => sum + (value - mean) ** 2, 0) / Math.max(1, values.length);
    };
    const denominator = variance(signal.map((value, index) => value + remainder[index]));
    return denominator ? clamp(1 - variance(remainder) / denominator, 0, 1) : 0;
  }

  function renderDecomposition(sku) {
    const product = productLookup().get(sku) || {};
    const packaged = sheet("Decomposition").filter((row) => row.SKU === sku).sort((a, b) => isoDate(a.Month).localeCompare(isoDate(b.Month)));
    const rows = packaged.length ? packaged.map((row) => ({
      month: isoDate(row.Month), observed: num(row.Observed), trend: num(row.Trend), seasonal: num(row.Seasonal), remainder: num(row.Remainder),
    })) : browserDecomposition(sku);
    const feature = sheet("Decomposition_Features").find((row) => row.SKU === sku);
    const remainder = rows.map((row) => row.remainder);
    const trendStrength = feature ? num(feature.Trend_Strength) : componentStrength(rows.map((row) => row.trend), remainder);
    const seasonalStrength = feature ? num(feature.Seasonal_Strength) : componentStrength(rows.map((row) => row.seasonal), remainder);
    const cycles = feature ? num(feature.Seasonal_Cycles) : rows.length / 12;
    el("forecast-source-note").textContent = packaged.length ? "Robust STL | period 12 | three observed annual cycles" : "Browser additive decomposition | offline robust STL unavailable";
    el("trend-strength").textContent = fmt(trendStrength, 2);
    el("seasonal-strength").textContent = fmt(seasonalStrength, 2);
    el("seasonal-cycles").textContent = fmt(cycles, 1);
    el("decomposition-chart-title").textContent = `${sku} | ${product.SKU_Description || "Demand history"}`;
    el("decomposition-subline").textContent = `${fmt(cycles, 1)} annual cycles provide limited evidence; component strength is descriptive, not a planning release signal.`;

    const svg = el("decomposition-chart");
    if (!rows.length) {
      svg.innerHTML = '<text x="20" y="40" class="svg-label">No demand history available.</text>';
      return;
    }
    const width = 900;
    const height = 590;
    const margin = { top: 18, right: 18, bottom: 42, left: 96 };
    const gap = 16;
    const panelH = (height - margin.top - margin.bottom - gap * 3) / 4;
    const plotW = width - margin.left - margin.right;
    const x = (index) => margin.left + index / Math.max(1, rows.length - 1) * plotW;
    const panels = [
      { key: "observed", label: "Observed", className: "observed" },
      { key: "trend", label: "Trend", className: "trend" },
      { key: "seasonal", label: "Seasonal", className: "seasonal" },
      { key: "remainder", label: "Remainder", className: "remainder" },
    ];
    const panelMarkup = panels.map((panel, panelIndex) => {
      const values = rows.map((row) => row[panel.key]);
      let low = Math.min(...values);
      let high = Math.max(...values);
      if (panel.key === "seasonal" || panel.key === "remainder") {
        const span = Math.max(Math.abs(low), Math.abs(high), 1);
        low = -span;
        high = span;
      }
      const padding = Math.max((high - low) * 0.08, 1);
      low -= padding;
      high += padding;
      const top = margin.top + panelIndex * (panelH + gap);
      const y = (value) => top + panelH - (value - low) / (high - low || 1) * panelH;
      const path = rows.map((row, index) => `${index ? "L" : "M"}${x(index).toFixed(1)},${y(row[panel.key]).toFixed(1)}`).join(" ");
      const zero = low <= 0 && high >= 0 ? `<line class="decomp-zero" x1="${margin.left}" x2="${width - margin.right}" y1="${y(0)}" y2="${y(0)}"></line>` : "";
      return `<rect class="decomp-band" x="${margin.left}" y="${top}" width="${plotW}" height="${panelH}"></rect>${[0, 0.5, 1].map((step) => { const value = low + (high - low) * step; return `<line class="svg-grid" x1="${margin.left}" x2="${width - margin.right}" y1="${y(value)}" y2="${y(value)}"></line><text class="svg-label decomp-y-label" x="${margin.left - 8}" y="${y(value) + 3}" text-anchor="end">${fmt(value)}</text>`; }).join("")}${zero}<text class="decomp-label" x="8" y="${top + panelH / 2 + 4}">${panel.label}</text><path class="decomp-line ${panel.className}" d="${path}"></path>`;
    }).join("");
    const labels = rows.map((row, index) => ({ row, index })).filter(({ index }) => index % 6 === 0 || index === rows.length - 1);
    svg.setAttribute("viewBox", `0 0 ${width} ${height}`);
    svg.innerHTML = `${panelMarkup}${labels.map(({ row, index }) => `<text class="svg-label decomp-x-label" x="${x(index)}" y="${height - 13}" text-anchor="${index === 0 ? "start" : index === rows.length - 1 ? "end" : "middle"}">${esc(monthLabel(row.month).replace(" ", " '"))}</text>`).join("")}`;
  }

  function truthy(value) {
    return ["yes", "true", "1", "on contract", "preferred"].includes(String(value ?? "").trim().toLowerCase());
  }

  function spendParent(row) {
    return {
      id: row.Supplier_Parent_ID || row.Supplier_Normalized_ID || row.Supplier_ID || "UNMAPPED",
      name: row.Supplier_Parent_Name || row.Supplier_Normalized_Name || row.Supplier_Name || "Unmapped supplier",
    };
  }

  function spendMonths() {
    return [...new Set(sheet("Spend_Lines").map((row) => isoDate(row.Transaction_Date).slice(0, 7)).filter((value) => /^\d{4}-\d{2}$/.test(value)))].sort();
  }

  function activeSpendMonths() {
    const months = spendMonths();
    if (state.spendPeriod === "Current") return new Set(months.slice(-12));
    if (state.spendPeriod === "Prior") return new Set(months.slice(-24, -12));
    return new Set(months);
  }

  function filteredSpendLines() {
    const months = activeSpendMonths();
    const query = state.commercialSearch.trim().toLowerCase();
    return sheet("Spend_Lines").filter((row) => {
      const parent = spendParent(row);
      const month = isoDate(row.Transaction_Date).slice(0, 7);
      if (!months.has(month)) return false;
      if (state.category !== "All" && row.Category_L1 !== state.category) return false;
      if (state.supplier !== "All" && parent.id !== state.supplier) return false;
      if (state.contract !== "All" && (truthy(row.On_Contract) ? "Yes" : "No") !== state.contract) return false;
      return !query || `${parent.name} ${row.Supplier_Name || ""} ${row.Category_L1 || ""} ${row.Part_ID || ""}`.toLowerCase().includes(query);
    });
  }

  function summarizeSpend(lines) {
    const positive = lines.reduce((sum, row) => sum + Math.max(0, num(row.Spend_Base)), 0);
    const net = lines.reduce((sum, row) => sum + num(row.Spend_Base), 0);
    const credits = lines.reduce((sum, row) => sum + Math.min(0, num(row.Spend_Base)), 0);
    const known = lines.reduce((sum, row) => sum + (String(row.On_Contract ?? "").trim() ? Math.max(0, num(row.Spend_Base)) : 0), 0);
    const contracted = lines.reduce((sum, row) => sum + (truthy(row.On_Contract) ? Math.max(0, num(row.Spend_Base)) : 0), 0);
    return {
      net,
      positive,
      credits,
      contractShare: known ? contracted / known : null,
      suppliers: new Set(lines.map((row) => spendParent(row).id)).size,
      transactions: new Set(lines.map((row) => `${row.Source_System || ""}|${row.Transaction_ID || ""}`)).size,
    };
  }

  function assignSpendPareto(rows) {
    const ordered = [...rows].sort((a, b) => b.positive - a.positive || a.name.localeCompare(b.name));
    const total = ordered.reduce((sum, row) => sum + row.positive, 0);
    let cumulative = 0;
    ordered.forEach((row, index) => {
      const priorShare = total ? cumulative / total : 0;
      cumulative += row.positive;
      row.rank = index + 1;
      row.share = total ? row.positive / total : 0;
      row.cumulativeShare = total ? cumulative / total : 0;
      row.abc = priorShare < 0.8 ? "A" : priorShare < 0.95 ? "B" : "C";
    });
    return ordered;
  }

  function groupedSpend(lines, dimension) {
    const groups = new Map();
    lines.forEach((row) => {
      const parent = spendParent(row);
      const id = dimension === "supplier" ? parent.id : row.Category_L1 || "Unclassified";
      const name = dimension === "supplier" ? parent.name : row.Category_L1 || "Unclassified";
      if (!groups.has(id)) groups.set(id, { id, name, positive: 0, net: 0, credits: 0, contracted: 0, known: 0, suppliers: new Set(), categories: new Set(), transactions: new Set(), impactTotal: 0, riskTotal: 0, scoreSpend: 0 });
      const group = groups.get(id);
      const spend = num(row.Spend_Base);
      const positive = Math.max(0, spend);
      group.positive += positive;
      group.net += spend;
      group.credits += Math.min(0, spend);
      if (String(row.On_Contract ?? "").trim()) group.known += positive;
      if (truthy(row.On_Contract)) group.contracted += positive;
      group.suppliers.add(parent.id);
      group.categories.add(row.Category_L1 || "Unclassified");
      group.transactions.add(`${row.Source_System || ""}|${row.Transaction_ID || ""}`);
      if (positive && Number.isFinite(Number(row.Business_Impact_Score)) && Number.isFinite(Number(row.Supply_Risk_Score))) {
        group.impactTotal += num(row.Business_Impact_Score) * positive;
        group.riskTotal += num(row.Supply_Risk_Score) * positive;
        group.scoreSpend += positive;
      }
    });
    return assignSpendPareto([...groups.values()].map((group) => ({
      ...group,
      supplierCount: group.suppliers.size,
      categoryCount: group.categories.size,
      transactionCount: group.transactions.size,
      contractShare: group.known ? group.contracted / group.known : null,
      impact: group.scoreSpend ? group.impactTotal / group.scoreSpend : null,
      risk: group.scoreSpend ? group.riskTotal / group.scoreSpend : null,
    })));
  }

  function commercialFlags(categoryRows) {
    const flags = [];
    categoryRows.forEach((row) => {
      const supplierSpend = new Map();
      filteredSpendLines().filter((line) => (line.Category_L1 || "Unclassified") === row.name).forEach((line) => {
        const parent = spendParent(line).id;
        supplierSpend.set(parent, (supplierSpend.get(parent) || 0) + Math.max(0, num(line.Spend_Base)));
      });
      const topShare = row.positive ? Math.max(0, ...supplierSpend.values()) / row.positive : 0;
      if (topShare >= 0.8) flags.push({ priority: "High", scope: row.name, flag: "Supplier concentration", evidence: `Top supplier share ${pct(topShare)}`, spend: row.positive });
      if (row.contractShare !== null && row.contractShare < 0.8) flags.push({ priority: row.contractShare < 0.5 ? "High" : "Medium", scope: row.name, flag: "Contract leakage", evidence: `On-contract share ${pct(row.contractShare)}`, spend: row.positive });
      if (row.supplierCount >= 5) flags.push({ priority: "Medium", scope: row.name, flag: "Supplier fragmentation", evidence: `${row.supplierCount} supplier parents`, spend: row.positive });
    });
    return flags.sort((a, b) => ({ High: 0, Medium: 1 })[a.priority] - ({ High: 0, Medium: 1 })[b.priority] || b.spend - a.spend);
  }

  function renderSpend() {
    const gate = spendDecisionGate();
    if (gate.blocked) {
      el("net-spend").textContent = "-";
      el("contract-coverage").textContent = "-";
      el("supplier-count").textContent = "-";
      el("spend-headline").textContent = "Financial interpretation is blocked.";
      el("spend-subline").textContent = gate.reason;
      el("commercial-context").textContent = "Decision gate blocked";
      renderPareto([]);
      renderSpendTrend([]);
      renderSupplierTable([]);
      el("opportunity-body").innerHTML = '<tr><td colspan="4" class="empty-state">Resolve the spend decision gate before reviewing opportunities.</td></tr>';
      return;
    }
    const lines = filteredSpendLines();
    const summary = summarizeSpend(lines);
    const categories = groupedSpend(lines, "category");
    const suppliers = groupedSpend(lines, "supplier");
    const flags = commercialFlags(categories);
    el("net-spend").textContent = money(summary.net);
    el("spend-scope").textContent = `${fmt(summary.transactions)} signed transactions`;
    el("contract-coverage").textContent = summary.contractShare === null ? "Unknown" : pct(summary.contractShare);
    el("supplier-count").textContent = fmt(summary.suppliers);
    el("spend-headline").textContent = flags.length ? `${flags.filter((row) => row.priority === "High").length} high-priority commercial reviews need evidence.` : "No commercial exception meets the current review thresholds.";
    el("spend-subline").textContent = `Gross positive spend ${money(summary.positive)}; credits ${money(summary.credits)}; no savings are claimed.`;
    const months = [...activeSpendMonths()].sort();
    el("commercial-context").textContent = months.length ? `${months[0]} to ${months.at(-1)} | ${fmt(lines.length)} signed lines` : "No spend lines in scope";
    renderPareto(state.paretoDimension === "supplier" ? suppliers : categories);
    renderSpendTrend(lines);
    renderSupplierTable(suppliers);
    el("opportunity-body").innerHTML = flags.length ? flags.map((row) => `<tr><td><span class="review-priority ${row.priority.toLowerCase()}">${esc(row.priority)}</span></td><td><strong>${esc(row.flag)}</strong><small>${esc(row.scope)}</small></td><td>${esc(row.evidence)}<small>Review prompt; no savings claim</small></td><td class="number">${esc(money(row.spend))}</td></tr>`).join("") : '<tr><td colspan="4" class="empty-state">No review prompt meets the current thresholds.</td></tr>';
  }

  function renderPareto(rows) {
    el("pareto-title").textContent = `${state.paretoDimension === "supplier" ? "Supplier" : "Category"} spend Pareto`;
    const svg = el("pareto-chart");
    const visible = rows.slice(0, 12);
    if (!visible.length) {
      svg.innerHTML = '<text x="20" y="40" class="svg-label">No positive spend in the active filter.</text>';
      return;
    }
    const width = 900;
    const rowHeight = 30;
    const height = 72 + visible.length * rowHeight;
    const margin = { top: 38, right: 52, bottom: 28, left: 185 };
    const plotW = width - margin.left - margin.right;
    const maxValue = Math.max(...visible.map((row) => row.positive), 1);
    const x = (value) => margin.left + (value / maxValue) * plotW;
    const cumulativePoints = visible.map((row, index) => `${index ? "L" : "M"}${margin.left + row.cumulativeShare * plotW},${margin.top + index * rowHeight + 11}`).join(" ");
    svg.setAttribute("viewBox", `0 0 ${width} ${height}`);
    svg.innerHTML = `<text class="axis-title" x="${margin.left}" y="11">Cumulative share (top axis)</text>${[0, 0.5, 1].map((value) => `<line class="svg-grid" x1="${margin.left + value * plotW}" x2="${margin.left + value * plotW}" y1="${margin.top - 7}" y2="${height - margin.bottom}"></line><text class="svg-label" x="${margin.left + value * plotW}" y="25" text-anchor="middle">${pct(value)}</text><text class="svg-label" x="${margin.left + value * plotW}" y="${height - 8}" text-anchor="middle">${money(maxValue * value)}</text>`).join("")}
      ${visible.map((row, index) => {
        const y = margin.top + index * rowHeight;
        return `<text class="svg-row-label" x="${margin.left - 9}" y="${y + 14}" text-anchor="end">${esc(row.name.length > 25 ? `${row.name.slice(0, 23)}...` : row.name)}</text><rect class="pareto-bar class-${row.abc.toLowerCase()}" x="${margin.left}" y="${y}" width="${Math.max(1, x(row.positive) - margin.left)}" height="18"></rect><text class="svg-value-label" x="${Math.min(width - margin.right + 5, x(row.positive) + 6)}" y="${y + 14}">${pct(row.share)}</text>`;
      }).join("")}
      <line class="pareto-threshold" x1="${margin.left + plotW * 0.8}" x2="${margin.left + plotW * 0.8}" y1="${margin.top - 7}" y2="${height - margin.bottom}"></line><text class="pareto-threshold-label" x="${margin.left + plotW * 0.8 - 4}" y="${margin.top - 11}" text-anchor="end">80% cumulative</text>
      <path class="pareto-line" d="${cumulativePoints}"></path>
      ${visible.map((row, index) => `<circle class="pareto-point" cx="${margin.left + row.cumulativeShare * plotW}" cy="${margin.top + index * rowHeight + 11}" r="3.5"><title>${esc(row.name)}: cumulative ${pct(row.cumulativeShare)}</title></circle>`).join("")}`;
  }

  function renderSpendTrend(lines) {
    const monthMap = new Map();
    lines.forEach((row) => {
      const month = isoDate(row.Transaction_Date).slice(0, 7);
      monthMap.set(month, (monthMap.get(month) || 0) + num(row.Spend_Base));
    });
    const rows = [...monthMap].sort(([a], [b]) => a.localeCompare(b)).map(([month, value]) => ({ month, value }));
    const svg = el("spend-trend-chart");
    el("trend-note").textContent = rows.length ? `${rows[0].month} to ${rows.at(-1).month}` : "No selected period";
    if (!rows.length) {
      svg.innerHTML = '<text x="20" y="40" class="svg-label">No monthly spend in the active filter.</text>';
      return;
    }
    const width = 760;
    const height = 260;
    const margin = { top: 18, right: 18, bottom: 38, left: 64 };
    const plotW = width - margin.left - margin.right;
    const plotH = height - margin.top - margin.bottom;
    const min = Math.min(0, ...rows.map((row) => row.value));
    const max = Math.max(1, ...rows.map((row) => row.value)) * 1.06;
    const x = (index) => margin.left + index / Math.max(1, rows.length - 1) * plotW;
    const y = (value) => margin.top + plotH - (value - min) / (max - min || 1) * plotH;
    const path = rows.map((row, index) => `${index ? "L" : "M"}${x(index)},${y(row.value)}`).join(" ");
    svg.setAttribute("viewBox", `0 0 ${width} ${height}`);
    svg.innerHTML = `${[0, 0.5, 1].map((step) => { const value = min + (max - min) * step; return `<line class="svg-grid" x1="${margin.left}" x2="${width - margin.right}" y1="${y(value)}" y2="${y(value)}"></line><text class="svg-label" x="${margin.left - 8}" y="${y(value) + 3}" text-anchor="end">${money(value)}</text>`; }).join("")}
      <path class="spend-line" d="${path}"></path>
      ${rows.map((row, index) => `<circle class="spend-point" cx="${x(index)}" cy="${y(row.value)}" r="3"><title>${esc(row.month)}: ${esc(money(row.value, false))}</title></circle>`).join("")}
      ${rows.filter((row, index) => index % 3 === 0 || index === rows.length - 1).map((row) => { const index = rows.indexOf(row); return `<text class="svg-label" x="${x(index)}" y="${height - 14}" text-anchor="middle">${esc(row.month)}</text>`; }).join("")}`;
  }

  function renderSupplierTable(rows) {
    el("supplier-body").innerHTML = rows.length ? rows.slice(0, 14).map((row) => `<tr><td><strong>${esc(row.name)}</strong><small>${esc(row.id)}</small></td><td class="number">${esc(money(row.positive))}</td><td class="number">${pct(row.share)}</td><td class="number">${fmt(row.categoryCount)}</td><td><span class="abc-class class-${row.abc.toLowerCase()}">${esc(row.abc)}</span></td></tr>`).join("") : '<tr><td colspan="5" class="empty-state">No suppliers match the active filters.</td></tr>';
  }

  function kraljicFor(row) {
    if (!Number.isFinite(row.impact) || !Number.isFinite(row.risk)) return "Unscored";
    if (row.impact >= 50 && row.risk >= 50) return "Strategic";
    if (row.impact >= 50) return "Leverage";
    if (row.risk >= 50) return "Bottleneck";
    return "Routine";
  }

  function categoryRiskLookup() {
    return new Map(sheet("Category_Risk").map((row) => [row.Category_L1, row]));
  }

  function renderPortfolio() {
    const gate = spendDecisionGate();
    if (gate.blocked) {
      el("strategic-spend").textContent = "-";
      el("bottleneck-spend").textContent = "-";
      el("single-source-count").textContent = "-";
      el("portfolio-headline").textContent = "Commercial portfolio scoring is blocked.";
      el("portfolio-subline").textContent = gate.reason;
      renderKraljic([]);
      renderPortfolioDetail(null);
      renderSegmentation();
      return;
    }
    const lines = filteredSpendLines();
    const categories = groupedSpend(lines, "category").map((row) => ({ ...row, quadrant: kraljicFor(row) }));
    if (!state.selectedCategory || !categories.some((row) => row.name === state.selectedCategory)) state.selectedCategory = categories[0]?.name || null;
    const strategic = categories.filter((row) => row.quadrant === "Strategic").reduce((sum, row) => sum + row.positive, 0);
    const bottleneck = categories.filter((row) => row.quadrant === "Bottleneck").reduce((sum, row) => sum + row.positive, 0);
    const allowedParts = new Set(lines.map((row) => row.Part_ID).filter(Boolean));
    const singleSource = sheet("Product_Master").filter((row) => allowedParts.has(row.SKU) && String(row.Single_Source) === "Yes").length;
    el("strategic-spend").textContent = money(strategic);
    el("bottleneck-spend").textContent = money(bottleneck);
    el("single-source-count").textContent = fmt(singleSource);
    const highRisk = categories.filter((row) => row.risk >= 50).length;
    el("portfolio-headline").textContent = highRisk ? `${highRisk} categories require continuity-led sourcing strategies.` : "No high-risk category remains in the selected scope.";
    el("portfolio-subline").textContent = `${categories.length} categories scored using explicit impact and supply-risk evidence; bubble size shows positive spend.`;
    renderKraljic(categories);
    renderPortfolioDetail(categories.find((row) => row.name === state.selectedCategory));
    renderSegmentation();
  }

  function renderKraljic(rows) {
    const svg = el("kraljic-chart");
    if (!rows.length) {
      svg.innerHTML = '<text x="20" y="40" class="svg-label">No scored category spend in the active filter.</text>';
      return;
    }
    const width = 760;
    const height = 430;
    const margin = { top: 24, right: 24, bottom: 52, left: 62 };
    const plotW = width - margin.left - margin.right;
    const plotH = height - margin.top - margin.bottom;
    const x = (value) => margin.left + clamp(value, 0, 100) / 100 * plotW;
    const y = (value) => margin.top + plotH - clamp(value, 0, 100) / 100 * plotH;
    const maxSpend = Math.max(...rows.map((row) => row.positive), 1);
    const radius = (value) => 7 + Math.sqrt(value / maxSpend) * 15;
    svg.setAttribute("viewBox", `0 0 ${width} ${height}`);
    svg.innerHTML = `<rect class="quadrant routine" x="${x(0)}" y="${y(50)}" width="${plotW / 2}" height="${plotH / 2}"></rect>
      <rect class="quadrant leverage" x="${x(50)}" y="${y(50)}" width="${plotW / 2}" height="${plotH / 2}"></rect>
      <rect class="quadrant bottleneck" x="${x(0)}" y="${y(100)}" width="${plotW / 2}" height="${plotH / 2}"></rect>
      <rect class="quadrant strategic" x="${x(50)}" y="${y(100)}" width="${plotW / 2}" height="${plotH / 2}"></rect>
      ${[0, 25, 50, 75, 100].map((tick) => `<line class="svg-grid" x1="${x(tick)}" x2="${x(tick)}" y1="${margin.top}" y2="${margin.top + plotH}"></line><line class="svg-grid" x1="${margin.left}" x2="${margin.left + plotW}" y1="${y(tick)}" y2="${y(tick)}"></line><text class="svg-label" x="${x(tick)}" y="${height - 27}" text-anchor="middle">${tick}</text><text class="svg-label" x="${margin.left - 9}" y="${y(tick) + 3}" text-anchor="end">${tick}</text>`).join("")}
      <line class="portfolio-boundary" x1="${x(50)}" x2="${x(50)}" y1="${margin.top}" y2="${margin.top + plotH}"></line><line class="portfolio-boundary" x1="${margin.left}" x2="${margin.left + plotW}" y1="${y(50)}" y2="${y(50)}"></line>
      <text class="quadrant-label" x="${x(25)}" y="${y(7)}" text-anchor="middle">Routine</text><text class="quadrant-label" x="${x(75)}" y="${y(7)}" text-anchor="middle">Leverage</text><text class="quadrant-label" x="${x(25)}" y="${y(94)}" text-anchor="middle">Bottleneck</text><text class="quadrant-label" x="${x(75)}" y="${y(94)}" text-anchor="middle">Strategic</text>
      <text class="axis-title" x="${margin.left + plotW / 2}" y="${height - 6}" text-anchor="middle">Business impact</text><text class="axis-title" x="15" y="${margin.top + plotH / 2}" text-anchor="middle" transform="rotate(-90 15 ${margin.top + plotH / 2})">Supply risk</text>
      ${rows.map((row) => `<g class="portfolio-point-group ${row.name === state.selectedCategory ? "selected" : ""}" data-category="${esc(row.name)}" tabindex="0"><circle class="portfolio-point ${row.quadrant.toLowerCase()}" cx="${x(row.impact)}" cy="${y(row.risk)}" r="${radius(row.positive)}" data-tip="${esc(`${row.name} | ${row.quadrant} | impact ${fmt(row.impact, 1)} | risk ${fmt(row.risk, 1)} | ${money(row.positive)}`)}"></circle><text class="portfolio-point-label" x="${x(row.impact)}" y="${y(row.risk) + 3}" text-anchor="middle">${esc((row.name.match(/\b[A-Z]/g) || []).join("").slice(0, 3) || row.name.slice(0, 3).toUpperCase())}</text></g>`).join("")}`;
    const tooltip = el("kraljic-tooltip");
    svg.querySelectorAll(".portfolio-point-group").forEach((group) => {
      const select = () => { state.selectedCategory = group.dataset.category; renderPortfolio(); };
      group.addEventListener("click", select);
      group.addEventListener("keydown", (event) => { if (event.key === "Enter" || event.key === " ") select(); });
      const point = group.querySelector("circle");
      point.addEventListener("pointerenter", (event) => {
        tooltip.textContent = point.dataset.tip;
        tooltip.style.display = "block";
        tooltip.style.left = `${Math.min(event.offsetX + 10, svg.clientWidth - 220)}px`;
        tooltip.style.top = `${Math.max(4, event.offsetY - 30)}px`;
      });
      point.addEventListener("pointerleave", () => { tooltip.style.display = "none"; });
    });
  }

  function renderPortfolioDetail(row) {
    if (!row) {
      el("portfolio-detail").innerHTML = '<p class="empty-state">Select a category to inspect its risk evidence and sourcing posture.</p>';
      return;
    }
    const risk = categoryRiskLookup().get(row.name) || {};
    const strategy = {
      Strategic: "Protect continuity, govern the relationship, share forward visibility, and qualify resilience options.",
      Leverage: "Use competitive tension and volume consolidation while protecting service and quality requirements.",
      Bottleneck: "Secure supply first, hold contingency, reduce switching barriers, and avoid savings actions that raise continuity risk.",
      Routine: "Standardize specifications, automate transactions, and reduce process cost before negotiating unit price.",
      Unscored: "Complete impact and risk evidence before assigning a sourcing posture.",
    }[row.quadrant];
    el("portfolio-detail").innerHTML = `<header class="detail-header"><p class="section-kicker">Category evidence</p><div class="detail-title-row"><h2>${esc(row.name)}</h2><span class="quadrant-tag ${row.quadrant.toLowerCase()}">${esc(row.quadrant)}</span></div><p>${money(row.positive)} positive spend | class ${esc(row.abc)}</p></header>
      <section class="detail-section"><h3>Recommended posture</h3><p class="detail-callout">${esc(strategy)}</p><dl class="formula-list"><div class="formula-row"><dt>Business impact</dt><dd>${fmt(row.impact, 1)} / 100</dd></div><div class="formula-row"><dt>Supply risk</dt><dd>${fmt(row.risk, 1)} / 100</dd></div><div class="formula-row"><dt>Supplier parents</dt><dd>${fmt(row.supplierCount)}</dd></div><div class="formula-row"><dt>On-contract share</dt><dd>${row.contractShare === null ? "Unknown" : pct(row.contractShare)}</dd></div></dl></section>
      <section class="detail-section"><h3>Impact evidence</h3><dl class="formula-list"><div class="formula-row"><dt>Spend impact</dt><dd>${fmt(risk.Spend_Impact_Score)} / 100</dd></div><div class="formula-row"><dt>Operational criticality</dt><dd>${fmt(risk.Operational_Criticality_Score)} / 100</dd></div><div class="formula-row"><dt>Revenue at risk</dt><dd>${fmt(risk.Revenue_at_Risk_Score)} / 100</dd></div><div class="formula-row"><dt>Quality / regulatory</dt><dd>${fmt(risk.Quality_Regulatory_Score)} / 100</dd></div></dl></section>
      <section class="detail-section"><h3>Supply-risk evidence</h3><dl class="formula-list"><div class="formula-row"><dt>Source scarcity</dt><dd>${fmt(risk.Source_Scarcity_Score)} / 100</dd></div><div class="formula-row"><dt>Switching time</dt><dd>${fmt(risk.Switching_Time_Days)} days</dd></div><div class="formula-row"><dt>Lead-time risk</dt><dd>${fmt(risk.Lead_Time_Risk_Score)} / 100</dd></div><div class="formula-row"><dt>Capacity risk</dt><dd>${fmt(risk.Capacity_Risk_Score)} / 100</dd></div></dl><p>${esc(risk.Evidence_Status || "Evidence status not supplied")}</p></section>`;
  }

  function filteredSegmentation() {
    const products = productLookup();
    const allowedParts = new Set(filteredSpendLines().map((row) => row.Part_ID).filter(Boolean));
    const query = state.commercialSearch.trim().toLowerCase();
    return sheet("Item_Segmentation").map((row) => ({ ...row, product: products.get(row.Part_ID) || {} })).filter((row) => {
      if (row.Segmentation_Status && row.Segmentation_Status !== "Valid") return false;
      if (!allowedParts.has(row.Part_ID)) return false;
      if (state.category !== "All" && row.product.Product_Group !== state.category) return false;
      return !query || `${row.Part_ID} ${row.product.SKU_Description || ""} ${row.product.Product_Group || ""}`.toLowerCase().includes(query);
    }).sort((a, b) => num(b.Annual_Usage_Value) - num(a.Annual_Usage_Value));
  }

  function renderSegmentation() {
    const rows = filteredSegmentation();
    const svg = el("segmentation-chart");
    el("segmentation-title").textContent = state.segmentationMode === "behavior" ? "Demand behavior beyond ABC/XYZ" : "Classic ABC/XYZ communication layer";
    el("segmentation-body").innerHTML = rows.length ? rows.slice(0, 24).map((row) => `<tr><td><strong>${esc(row.Part_ID)}</strong><small>${esc(row.product.Product_Group || "")}</small></td><td class="number">${money(row.Annual_Usage_Value)}</td><td><span class="abc-class class-${String(row.Usage_Value_ABC || "c").toLowerCase()}">${esc(row.ABC_XYZ || "-")}</span></td><td>${esc(row.Demand_Pattern || "-")}</td><td>${esc(row.Criticality || row.product.Operational_Criticality || "-")}</td><td class="number">${fmt(row.product.Qualified_Source_Count)}</td></tr>`).join("") : '<tr><td colspan="6" class="empty-state">No item segmentation matches the active commercial filters.</td></tr>';
    if (!rows.length) {
      svg.innerHTML = '<text x="20" y="40" class="svg-label">No valid item histories in the active filter.</text>';
      return;
    }
    const width = 760;
    const height = 360;
    const margin = { top: 20, right: 20, bottom: 50, left: 60 };
    const plotW = width - margin.left - margin.right;
    const plotH = height - margin.top - margin.bottom;
    let xValue;
    let yValue;
    let xDomain;
    let yMax;
    if (state.segmentationMode === "behavior") {
      xValue = (row) => num(row.ADI);
      yValue = (row) => num(row.CV2_Nonzero_Demand);
      xDomain = [Math.min(0.8, ...rows.map(xValue)), Math.max(1.55, ...rows.map(xValue))];
      yMax = Math.max(0.6, ...rows.map(yValue)) * 1.05;
    } else {
      xValue = (row) => Math.log10(Math.max(1, num(row.Annual_Usage_Value)));
      yValue = (row) => num(row.Demand_CV);
      xDomain = [Math.min(...rows.map(xValue)), Math.max(...rows.map(xValue))];
      yMax = Math.max(1.1, ...rows.map(yValue)) * 1.04;
    }
    const x = (value) => margin.left + (value - xDomain[0]) / (xDomain[1] - xDomain[0] || 1) * plotW;
    const y = (value) => margin.top + plotH - value / (yMax || 1) * plotH;
    const className = (row) => state.segmentationMode === "behavior" ? String(row.Demand_Pattern || "smooth").toLowerCase() : `class-${String(row.Usage_Value_ABC || "c").toLowerCase()}`;
    svg.setAttribute("viewBox", `0 0 ${width} ${height}`);
    const thresholds = state.segmentationMode === "behavior"
      ? `<line class="portfolio-boundary" x1="${x(1.32)}" x2="${x(1.32)}" y1="${margin.top}" y2="${margin.top + plotH}"></line><line class="portfolio-boundary" x1="${margin.left}" x2="${margin.left + plotW}" y1="${y(0.49)}" y2="${y(0.49)}"></line><text class="quadrant-label" x="${x(1.05)}" y="${y(0.05)}">Smooth</text><text class="quadrant-label" x="${x(1.05)}" y="${y(0.54)}">Erratic</text><text class="quadrant-label" x="${x(1.36)}" y="${y(0.05)}">Intermittent</text><text class="quadrant-label" x="${x(1.36)}" y="${y(0.54)}">Lumpy</text>`
      : `<line class="portfolio-boundary" x1="${margin.left}" x2="${margin.left + plotW}" y1="${y(0.5)}" y2="${y(0.5)}"></line><line class="portfolio-boundary" x1="${margin.left}" x2="${margin.left + plotW}" y1="${y(1.0)}" y2="${y(1.0)}"></line>`;
    svg.innerHTML = `${[0, 0.5, 1].map((step) => `<line class="svg-grid" x1="${margin.left}" x2="${margin.left + plotW}" y1="${margin.top + step * plotH}" y2="${margin.top + step * plotH}"></line>`).join("")}${thresholds}
      ${rows.map((row) => `<circle class="segmentation-point ${className(row)}" cx="${x(xValue(row))}" cy="${y(yValue(row))}" r="5" data-tip="${esc(`${row.Part_ID} | ${row.ABC_XYZ} | ${row.Demand_Pattern} | ${money(row.Annual_Usage_Value)}`)}"></circle>`).join("")}
      <text class="axis-title" x="${margin.left + plotW / 2}" y="${height - 7}" text-anchor="middle">${state.segmentationMode === "behavior" ? "Average demand interval (ADI)" : "Annual usage value (log scale)"}</text><text class="axis-title" x="15" y="${margin.top + plotH / 2}" text-anchor="middle" transform="rotate(-90 15 ${margin.top + plotH / 2})">${state.segmentationMode === "behavior" ? "CV2 of non-zero demand" : "Demand coefficient of variation"}</text>`;
    const tooltip = el("segmentation-tooltip");
    svg.querySelectorAll(".segmentation-point").forEach((point) => {
      point.addEventListener("pointerenter", (event) => {
        tooltip.textContent = point.dataset.tip;
        tooltip.style.display = "block";
        tooltip.style.left = `${Math.min(event.offsetX + 10, svg.clientWidth - 220)}px`;
        tooltip.style.top = `${Math.max(4, event.offsetY - 30)}px`;
      });
      point.addEventListener("pointerleave", () => { tooltip.style.display = "none"; });
    });
  }

  function filteredSupply() {
    const products = productLookup();
    const plan = new Map(allPositions("baseline").map((row) => [row.key, row]));
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
      return { ...row, product, position, days, risk, closed, eligible: truthy(row.Eligible_for_NFP) };
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
    const onhandRisk = filteredPositions("baseline").filter((row) => row.executionRisk).length;
    const total = supply.filter((row) => !row.closed && row.eligible).length;
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

  function spendDecisionGate() {
    const lines = sheet("Spend_Lines");
    if (!lines.length) return { blocked: true, reason: "No signed spend ledger is loaded." };
    const currencies = [...new Set(lines.map((row) => String(row.Base_Currency || "").trim()).filter(Boolean))];
    if (currencies.length !== 1 || currencies[0] !== "EUR") return { blocked: true, reason: `Expected one EUR base currency; found ${currencies.join(", ") || "none"}.` };
    const failedChecks = sheet("Spend_Data_Quality").filter((row) => String(row.Status).toUpperCase() === "FAIL");
    if (failedChecks.length) return { blocked: true, reason: `${failedChecks.length} packaged spend quality gate(s) failed.` };
    const control = sheet("Spend_Control")[0];
    if (!control || !hasNumber(control.Control_Net_Spend)) return { blocked: true, reason: "An independent signed net-spend control total is required." };
    const net = lines.reduce((sum, row) => sum + num(row.Spend_Base), 0);
    const difference = net - num(control.Control_Net_Spend);
    if (Math.abs(difference) > 0.01) return { blocked: true, reason: `Signed spend differs from the control total by ${money(difference, false)}.` };
    return { blocked: false, reason: `Signed spend reconciles within EUR 0.01 (${money(net, false)}).` };
  }

  function demandCompletenessAudit() {
    const rows = sheet("Demand_Data");
    if (!rows.length) return { pass: false, evidence: "No demand rows" };
    const keys = new Set();
    let duplicates = 0;
    let invalid = 0;
    const groups = new Map();
    rows.forEach((row) => {
      const month = isoDate(row.Month);
      const key = `${row.SKU}|${month}`;
      if (keys.has(key)) duplicates += 1;
      keys.add(key);
      if (!row.SKU || month === "-" || !hasNumber(row.Demand_Units) || num(row.Demand_Units) < 0) invalid += 1;
      if (!groups.has(row.SKU)) groups.set(row.SKU, new Set());
      groups.get(row.SKU).add(month);
    });
    const allMonths = [...new Set(rows.map((row) => isoDate(row.Month)).filter((value) => value !== "-"))].sort();
    const start = dateValue(allMonths[0]);
    const end = dateValue(allMonths.at(-1));
    const expected = start && end ? (end.getUTCFullYear() - start.getUTCFullYear()) * 12 + end.getUTCMonth() - start.getUTCMonth() + 1 : 0;
    const incomplete = [...groups.values()].filter((months) => months.size !== expected || [...months].some((month) => !allMonths.includes(month))).length;
    return {
      pass: !duplicates && !invalid && expected === allMonths.length && !incomplete,
      evidence: `${groups.size} SKUs x ${expected} contiguous months; ${duplicates} duplicates; ${incomplete} incomplete histories; ${invalid} invalid values`,
    };
  }

  function recommendationAudit() {
    const recommendations = new Map(sheet("DDMRP_Recommendations").map((row) => [keyFor(row), row]));
    const positions = allPositions("baseline");
    const invalid = positions.filter((row) => !row.calculationValid);
    let mismatches = 0;
    positions.filter((row) => row.calculationValid).forEach((row) => {
      const stored = recommendations.get(row.key);
      if (!stored) {
        mismatches += 1;
        return;
      }
      const checks = [
        [row.tor, stored.TOR], [row.toy, stored.TOY], [row.tog, stored.TOG],
        [row.nfp, stored.Net_Flow_Position], [row.recommended, stored.Recommended_Order_Units],
      ];
      if (checks.some(([calculated, saved]) => !hasNumber(saved) || Math.abs(calculated - num(saved)) > 0.01)) mismatches += 1;
    });
    return {
      pass: !invalid.length && !mismatches && positions.length > 0,
      evidence: `${positions.length} positions; ${invalid.length} blocked inputs; ${mismatches} stored-formula mismatches at EUR/unit tolerance 0.01`,
    };
  }

  function segmentationCaveat() {
    const bySku = new Map();
    sheet("Demand_Data").forEach((row) => {
      if (!bySku.has(row.SKU)) bySku.set(row.SKU, { profile: row.Demand_Profile, zeros: 0 });
      if (num(row.Demand_Units) === 0) bySku.get(row.SKU).zeros += 1;
    });
    const mismatched = [...bySku.values()].filter((row) => /intermittent/i.test(String(row.profile)) && row.zeros === 0).length;
    return { mismatched, evidence: `${mismatched} SKU profile labels say intermittent but contain no zero-demand months` };
  }

  function projectAuditRows() {
    const demand = demandCompletenessAudit();
    const ddmrp = recommendationAudit();
    const spend = spendDecisionGate();
    const scores = sheet("Model_Selection");
    const scaledComplete = scores.length > 0 && scores.every((row) => hasNumber(row.MASE) && hasNumber(row.RMSSE));
    const tirexRows = scores.filter((row) => row.Model === "TiRex2").length;
    const metadata = sheet("Forecast_Run_Metadata")[0] || {};
    const forecastPass = state.sourceMode === "packaged" && scaledComplete && tirexRows === 80 && metadata.Input_SHA256 === state.data.workbookSha256;
    const segmentation = segmentationCaveat();
    const lateEligible = sheet("Open_Supply").filter((row) => truthy(row.Eligible_for_NFP) && dateValue(row.Expected_Receipt_Date) < dateValue(snapshotDate())).length;
    const demandGroups = new Map();
    sheet("Demand_Data").forEach((row) => {
      if (!demandGroups.has(row.SKU)) demandGroups.set(row.SKU, []);
      demandGroups.get(row.SKU).push(num(row.Demand_Units));
    });
    const xyzBySku = new Map([...demandGroups].map(([sku, values]) => {
      const mean = values.reduce((sum, value) => sum + value, 0) / Math.max(1, values.length);
      const variance = values.reduce((sum, value) => sum + (value - mean) ** 2, 0) / Math.max(1, values.length);
      const cv = mean ? Math.sqrt(variance) / mean : Infinity;
      return [sku, cv < 0.5 ? "X" : cv < 1 ? "Y" : "Z"];
    }));
    const xyzMismatches = sheet("Product_Master").filter((row) => xyzBySku.has(row.SKU) && row.XYZ_Class !== xyzBySku.get(row.SKU)).length;
    const spendMonths = [...new Set(sheet("Spend_Lines").map((row) => isoDate(row.Transaction_Date).slice(0, 7)))].sort();
    const currentMonths = new Set(spendMonths.slice(-12));
    const annualOnContract = sheet("Spend_Lines").filter((row) => currentMonths.has(isoDate(row.Transaction_Date).slice(0, 7)) && truthy(row.On_Contract)).reduce((sum, row) => sum + Math.max(0, num(row.Spend_Base)), 0);
    const commitments = sheet("Contracts").reduce((sum, row) => sum + Math.max(0, num(row.Committed_Annual_Value_EUR)), 0);
    const commitmentRatio = commitments ? annualOnContract / commitments : null;
    const boundaryCategories = sheet("Category_Risk").filter((row) => Math.abs(num(row.Business_Impact_Score) - 50) <= 5 || Math.abs(num(row.Supply_Risk_Score) - 50) <= 5).length;
    const syntheticControl = /synthetic/i.test(String(state.data.sourceType)) || state.data.workbookSha256 === SAMPLE_DATA.workbookSha256;
    return [
      { control: "Demand calendar and keys", result: demand.pass ? "PASS" : "BLOCKED", evidence: demand.evidence, impact: demand.pass ? "Forecast evaluation may proceed." : "Forecast and segmentation calculations are blocked." },
      { control: "DDMRP formulas and inputs", result: ddmrp.pass ? "PASS" : "BLOCKED", evidence: ddmrp.evidence, impact: ddmrp.pass ? "Stored recommendations reproduce from atomic inputs." : "Order recommendations are blocked where evidence fails." },
      { control: "Signed spend reconciliation", result: spend.blocked ? "BLOCKED" : "PASS", evidence: spend.reason, impact: spend.blocked ? "Financial interpretation is suppressed." : "Spend totals may be analyzed." },
      { control: "Forecast artifact provenance", result: forecastPass ? "PASS" : "NOT RUN", evidence: forecastPass ? `Workbook SHA-256 matches; 7 candidates x 80 SKUs; TiRex2 revision ${metadata.TiRex2_Revision || "recorded"}.` : "No content-matched offline forecast package is attached.", impact: forecastPass ? "Serious model selection and STL are available." : "Only browser benchmarks are shown." },
      { control: "Seasonal history depth", result: "WARN", evidence: "36 monthly observations provide exactly three annual cycles.", impact: "Use STL components as diagnostics; seasonal evidence is not definitive." },
      { control: "Synthetic intermittency fixture", result: segmentation.mismatched ? "WARN" : "PASS", evidence: segmentation.evidence, impact: segmentation.mismatched ? "Do not claim intermittent-demand validation from this fixture." : "Observed zeros support demand-pattern classification." },
      { control: "Buffer variability evidence", result: xyzMismatches ? "WARN" : "PASS", evidence: `${xyzMismatches} approved XYZ classes differ from the observed 36-month coefficient-of-variation class.`, impact: xyzMismatches ? "Treat variability factors as master-data assumptions and review them before production use." : "Observed variability supports the stored class." },
      { control: "Contract commitment plausibility", result: commitmentRatio !== null && commitmentRatio > 1.25 ? "WARN" : "PASS", evidence: commitmentRatio === null ? "No annual commitment baseline." : `Current on-contract positive spend is ${fmt(commitmentRatio, 2)}x recorded annual commitments.`, impact: commitmentRatio !== null && commitmentRatio > 1.25 ? "Do not use commitment variance as a sourcing signal until contract baselines are corrected." : "Commitments are directionally aligned with current spend." },
      { control: "Kraljic threshold sensitivity", result: boundaryCategories ? "WARN" : "PASS", evidence: `${boundaryCategories} categories sit within five score points of a 50-point quadrant boundary.`, impact: boundaryCategories ? "Review factor evidence rather than treating the quadrant label as stable." : "No category is close to a quadrant boundary." },
      { control: "Control-total independence", result: syntheticControl ? "WARN" : "NOT TESTED", evidence: syntheticControl ? "The synthetic control total was generated from the same ledger." : "Source-system independence cannot be proven inside the workbook.", impact: "Reconciliation validates transformation arithmetic, not source extraction completeness." },
      { control: "Late supply eligibility", result: lateEligible ? "WARN" : "PASS", evidence: `${lateEligible} late open orders remain eligible for net-flow position.`, impact: lateEligible ? "Review eligibility before releasing replenishment." : "No late eligible supply found." },
    ];
  }

  function renderData() {
    const sheets = state.data.sheets || {};
    const requiredMissing = REQUIRED_SHEETS.filter((name) => !(sheets[name] || []).length);
    el("demand-record-count").textContent = fmt(sheet("Demand_Data").length);
    el("position-count").textContent = fmt(sheet("DDMRP_Positions").length);
    el("spend-line-count").textContent = fmt(sheet("Spend_Lines").length);
    const gates = projectAuditRows();
    const blocked = gates.filter((row) => row.result === "BLOCKED").length;
    const warnings = gates.filter((row) => row.result === "WARN").length;
    el("data-headline").textContent = requiredMissing.length ? `Workbook is missing ${requiredMissing.join(" and ")}.` : blocked ? `${blocked} decision gate${blocked === 1 ? " is" : "s are"} blocked.` : warnings ? `Workbook passes core calculations with ${warnings} declared warning${warnings === 1 ? "" : "s"}.` : "Workbook passes all active decision gates.";
    el("data-subline").textContent = "Demand, planning, procurement, model artifacts, and control totals remain separate and auditable.";
    el("quality-gate-body").innerHTML = gates.map((row) => `<tr><td><strong>${esc(row.control)}</strong></td><td><span class="gate-result gate-${row.result.toLowerCase().replace(" ", "-")}">${esc(row.result)}</span></td><td>${esc(row.evidence)}</td><td>${esc(row.impact)}</td></tr>`).join("");
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

  function populateCommercialFilters() {
    const categories = [...new Set(sheet("Spend_Lines").map((row) => row.Category_L1).filter(Boolean))].sort();
    const supplierMap = new Map();
    sheet("Spend_Lines").forEach((row) => { const parent = spendParent(row); supplierMap.set(parent.id, parent.name); });
    const suppliers = [...supplierMap].sort((a, b) => a[1].localeCompare(b[1]));
    el("category-filter").innerHTML = '<option value="All">All categories</option>' + categories.map((value) => `<option value="${esc(value)}">${esc(value)}</option>`).join("");
    el("supplier-filter").innerHTML = '<option value="All">All suppliers</option>' + suppliers.map(([id, name]) => `<option value="${esc(id)}">${esc(name)}</option>`).join("");
    if (!categories.includes(state.category)) state.category = "All";
    if (!supplierMap.has(state.supplier)) state.supplier = "All";
    el("category-filter").value = state.category;
    el("supplier-filter").value = state.supplier;
    el("spend-period-filter").value = state.spendPeriod;
    el("contract-filter").value = state.contract;
  }

  function renderAll() {
    populateFilters();
    populateCommercialFilters();
    renderPlanning();
    renderForecast();
    renderSpend();
    renderPortfolio();
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
      tab.tabIndex = active ? 0 : -1;
    });
    document.querySelectorAll(".view").forEach((panel) => {
      const active = panel.id === `view-${view}`;
      panel.classList.toggle("active", active);
      panel.hidden = !active;
    });
    const commercialView = view === "spend" || view === "portfolio";
    el("planning-filter-band").hidden = commercialView;
    el("commercial-filter-band").hidden = !commercialView;
    if (view === "forecast") requestAnimationFrame(renderForecast);
    if (view === "spend") requestAnimationFrame(renderSpend);
    if (view === "portfolio") requestAnimationFrame(renderPortfolio);
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
    downloadCsv("demand-genie-v4-recommendations.csv", filteredPositions().map((row) => ({
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
      Calculation_Status: row.calculationValid ? "Valid" : `Blocked: ${row.invalidFields.join("; ")}`,
      Planner_Decision: decisionLabel(row),
    })));
  }

  function exportForecast() {
    const bundle = getForecastBundle(state.selectedSku);
    downloadCsv(`demand-genie-v4-forecast-${state.selectedSku}.csv`, bundle.forecast.map((row) => ({
      SKU: state.selectedSku,
      Month: row.month,
      Model: bundle.selected,
      Forecast_Units: round(row.point, 2),
      PI80_Lower: row.lower80 === null ? "" : round(row.lower80, 2),
      PI80_Upper: row.upper80 === null ? "" : round(row.upper80, 2),
      PI95_Lower: row.lower95 === null ? "" : round(row.lower95, 2),
      PI95_Upper: row.upper95 === null ? "" : round(row.upper95, 2),
      Interval_Method: bundle.intervalMethod || "Not available",
      Engine: bundle.source === "packaged" ? state.data.forecastEngine : "Browser benchmarks",
    })));
  }

  function exportSpend() {
    const lines = filteredSpendLines();
    const categories = groupedSpend(lines, "category").map((row) => ({
      Scope_Type: "Category",
      Scope_ID: row.id,
      Scope: row.name,
      Positive_Spend_EUR: round(row.positive, 2),
      Net_Spend_EUR: round(row.net, 2),
      Positive_Spend_Share: round(row.share, 6),
      Cumulative_Share: round(row.cumulativeShare, 6),
      Spend_ABC: row.abc,
      Supplier_Count: row.supplierCount,
      On_Contract_Share: row.contractShare === null ? "" : round(row.contractShare, 6),
      Business_Impact_Score: round(row.impact, 1),
      Supply_Risk_Score: round(row.risk, 1),
      Kraljic_Quadrant: kraljicFor(row),
      Savings_Claim: "None - review prompt only",
    }));
    const suppliers = groupedSpend(lines, "supplier").map((row) => ({
      Scope_Type: "Supplier parent",
      Scope_ID: row.id,
      Scope: row.name,
      Positive_Spend_EUR: round(row.positive, 2),
      Net_Spend_EUR: round(row.net, 2),
      Positive_Spend_Share: round(row.share, 6),
      Cumulative_Share: round(row.cumulativeShare, 6),
      Spend_ABC: row.abc,
      Supplier_Count: "",
      On_Contract_Share: row.contractShare === null ? "" : round(row.contractShare, 6),
      Business_Impact_Score: "",
      Supply_Risk_Score: "",
      Kraljic_Quadrant: "",
      Savings_Claim: "None - review prompt only",
    }));
    downloadCsv("demand-genie-v4-spend-analysis.csv", [...categories, ...suppliers]);
  }

  function browserItemSegmentation(rows) {
    const groups = new Map();
    rows.forEach((row) => {
      const key = `${row.Part_ID}|${row.Location || ""}`;
      if (!groups.has(key)) groups.set(key, []);
      groups.get(key).push(row);
    });
    const globalPeriods = [...new Set(rows.map((row) => isoDate(row.Period)).filter((value) => value !== "-"))].sort();
    const output = [...groups.entries()].map(([key, members]) => {
      members.sort((a, b) => isoDate(a.Period).localeCompare(isoDate(b.Period)));
      const periods = members.map((row) => isoDate(row.Period));
      const complete = members.length === globalPeriods.length && new Set(periods).size === periods.length && globalPeriods.every((period) => periods.includes(period));
      const validValues = members.every((row) => hasNumber(row.Demand_Units) && num(row.Demand_Units) >= 0 && hasNumber(row.Unit_Cost_Base) && num(row.Unit_Cost_Base) >= 0);
      const values = members.map((row) => num(row.Demand_Units, NaN));
      if (!complete || !validValues) {
        return {
          Part_ID: key.split("|")[0], Location: key.split("|")[1], Annual_Usage_Value: 0,
          Complete_History: complete ? "Yes" : "No", Segmentation_Status: "Invalid",
        };
      }
      const mean = values.length ? values.reduce((a, b) => a + b, 0) / values.length : 0;
      const variance = values.length ? values.reduce((sum, value) => sum + (value - mean) ** 2, 0) / values.length : 0;
      const cv = mean ? Math.sqrt(variance) / mean : null;
      const nonzero = values.filter((value) => value > 0);
      const nzMean = nonzero.length ? nonzero.reduce((a, b) => a + b, 0) / nonzero.length : 0;
      const nzVariance = nonzero.length ? nonzero.reduce((sum, value) => sum + (value - nzMean) ** 2, 0) / nonzero.length : 0;
      const adi = nonzero.length ? values.length / nonzero.length : null;
      const cv2 = nzMean ? (Math.sqrt(nzVariance) / nzMean) ** 2 : null;
      const pattern = adi === null || cv2 === null ? "No demand" : adi < 1.32 && cv2 < 0.49 ? "Smooth" : adi < 1.32 ? "Erratic" : cv2 < 0.49 ? "Intermittent" : "Lumpy";
      const xyz = cv === null ? "No demand" : cv < 0.5 ? "X" : cv < 1 ? "Y" : "Z";
      const latest = members.at(-1) || {};
      return {
        Part_ID: key.split("|")[0], Location: key.split("|")[1], Mean_Demand: mean, Demand_CV: cv, XYZ: xyz,
        ADI: adi, CV2_Nonzero_Demand: cv2, Demand_Pattern: pattern, Approved_Unit_Cost_Base: num(latest.Unit_Cost_Base),
        Annualized_Usage_Units: mean * 12, Annual_Usage_Value: mean * 12 * num(latest.Unit_Cost_Base), Criticality: latest.Criticality || "",
        Lifecycle_Status: latest.Lifecycle_Status || "", Complete_History: "Yes", Segmentation_Status: "Valid",
      };
    }).filter((row) => row.Segmentation_Status !== "Valid" || row.Annual_Usage_Value > 0).sort((a, b) => b.Annual_Usage_Value - a.Annual_Usage_Value);
    const validOutput = output.filter((row) => row.Segmentation_Status === "Valid");
    const total = validOutput.reduce((sum, row) => sum + row.Annual_Usage_Value, 0);
    let cumulative = 0;
    validOutput.forEach((row, index) => {
      const prior = total ? cumulative / total : 0;
      cumulative += row.Annual_Usage_Value;
      row.Rank = index + 1;
      row.Usage_Value_ABC = prior < 0.8 ? "A" : prior < 0.95 ? "B" : "C";
      row.ABC_XYZ = `${row.Usage_Value_ABC}${row.XYZ}`;
      row.Annual_Usage_Value_Share = total ? row.Annual_Usage_Value / total : 0;
      row.Cumulative_Annual_Usage_Value_Share = total ? cumulative / total : 0;
    });
    return output;
  }

  async function sha256Buffer(buffer) {
    if (!window.crypto?.subtle) return null;
    const digest = await window.crypto.subtle.digest("SHA-256", buffer);
    return [...new Uint8Array(digest)].map((byte) => byte.toString(16).padStart(2, "0")).join("");
  }

  function updateForecastModelOptions() {
    const offlineOnly = new Set(["ARIMA", "TiRex2"]);
    [...el("forecast-model").options].forEach((option) => {
      option.disabled = state.sourceMode !== "packaged" && offlineOnly.has(option.value);
    });
    if (el("forecast-model").selectedOptions[0]?.disabled) {
      state.forecastModel = "Auto";
      el("forecast-model").value = "Auto";
    }
  }

  async function loadWorkbook(file) {
    try {
      updateDataAlert(`Loading ${file.name}...`);
      const bytes = await file.arrayBuffer();
      const workbookHash = await sha256Buffer(bytes);
      const workbook = XLSX.read(bytes, { type: "array", cellDates: true });
      const sheets = {};
      Object.keys(SHEET_PURPOSES).forEach((name) => {
        if (!workbook.SheetNames.includes(name)) return;
        sheets[name] = normalizeRows(XLSX.utils.sheet_to_json(workbook.Sheets[name], { defval: null, raw: true }));
      });
      const exactEmbeddedWorkbook = workbookHash && workbookHash === SAMPLE_DATA.workbookSha256;
      if (exactEmbeddedWorkbook) {
        ["Forecast_Results", "All_Model_Forecasts", "Model_Selection", "Model_Summary", "Forecast_Exceptions", "Forecast_Run_Metadata", "Decomposition", "Decomposition_Features"].forEach((name) => { sheets[name] = cloneData(SAMPLE_DATA.sheets[name] || []); });
        ["Spend_Summary", "Supplier_Summary", "Category_Summary", "Monthly_Spend", "Opportunity_Flags", "Spend_Data_Quality", "Item_Segmentation"].forEach((name) => { sheets[name] = cloneData(SAMPLE_DATA.sheets[name] || []); });
        state.sourceMode = "packaged";
      } else {
        state.sourceMode = "browser";
        if (!sheets.Item_Segmentation?.length && sheets.Demand_History?.length) sheets.Item_Segmentation = browserItemSegmentation(sheets.Demand_History);
      }
      state.data = {
        fileName: file.name,
        workbookSha256: workbookHash,
        sourceType: "Uploaded workbook",
        forecastEngine: state.sourceMode === "packaged" ? SAMPLE_DATA.forecastEngine : "Browser benchmarks only",
        forecastProtocol: state.sourceMode === "packaged" ? SAMPLE_DATA.forecastProtocol : "Single six-month benchmark holdout",
        spendEngine: "Local signed-spend analysis",
        sheets,
      };
      state.group = "All";
      state.location = "All";
      state.status = "All";
      state.search = "";
      state.selectedKey = null;
      state.selectedSku = null;
      state.selectedBufferKey = null;
      state.forecastModel = "Auto";
      state.forecastMode = "forecast";
      state.spendPeriod = "All";
      state.category = "All";
      state.supplier = "All";
      state.contract = "All";
      state.commercialSearch = "";
      state.selectedCategory = null;
      el("group-filter").value = "All";
      el("location-filter").value = "All";
      el("status-filter").value = "All";
      el("search-filter").value = "";
      el("commercial-search").value = "";
      el("forecast-model").value = "Auto";
      document.querySelectorAll("#forecast-mode-control button").forEach((button) => {
        const active = button.dataset.forecastMode === "forecast";
        button.classList.toggle("active", active);
        button.setAttribute("aria-pressed", String(active));
      });
      el("source-kind").textContent = "Uploaded workbook";
      el("source-file").textContent = file.name;
      el("source-freshness").textContent = `Snapshot ${snapshotDate()}`;
      updateForecastModelOptions();
      renderAll();
      const missing = REQUIRED_SHEETS.filter((name) => !sheet(name).length);
      updateDataAlert(missing.length ? "" : state.sourceMode === "browser" ? "No content-matched offline forecast package is attached. Serious model selection, TiRex2, native intervals, and robust STL are unavailable for this upload." : "");
      showToast(`${file.name} loaded: ${fmt(sheet("Demand_Data").length)} demand rows, ${fmt(sheet("DDMRP_Positions").length)} positions, and ${fmt(sheet("Spend_Lines").length)} spend lines.`);
    } catch (error) {
      console.error(error);
      updateDataAlert(`Could not load ${file.name}. Confirm that it is an unencrypted Excel workbook with named header rows.`);
    } finally {
      el("workbook-input").value = "";
    }
  }

  function bindEvents() {
    const tabs = [...document.querySelectorAll(".tab")];
    tabs.forEach((tab, index) => {
      tab.addEventListener("click", () => switchView(tab.dataset.view));
      tab.addEventListener("keydown", (event) => {
        if (!['ArrowLeft', 'ArrowRight', 'Home', 'End'].includes(event.key)) return;
        event.preventDefault();
        const nextIndex = event.key === 'Home' ? 0 : event.key === 'End' ? tabs.length - 1 : (index + (event.key === 'ArrowRight' ? 1 : -1) + tabs.length) % tabs.length;
        switchView(tabs[nextIndex].dataset.view);
        tabs[nextIndex].focus();
      });
    });
    el("upload-button").addEventListener("click", () => el("workbook-input").click());
    el("workbook-input").addEventListener("change", (event) => { if (event.target.files[0]) loadWorkbook(event.target.files[0]); });
    el("group-filter").addEventListener("change", (event) => { state.group = event.target.value; renderAll(); });
    el("location-filter").addEventListener("change", (event) => { state.location = event.target.value; renderAll(); });
    el("status-filter").addEventListener("change", (event) => { state.status = event.target.value; renderAll(); });
    el("search-filter").addEventListener("input", (event) => { state.search = event.target.value; renderPlanning(); renderExecution(); renderBuffers(); });
    el("spend-period-filter").addEventListener("change", (event) => { state.spendPeriod = event.target.value; state.selectedCategory = null; renderSpend(); renderPortfolio(); });
    el("category-filter").addEventListener("change", (event) => { state.category = event.target.value; state.selectedCategory = null; renderSpend(); renderPortfolio(); });
    el("supplier-filter").addEventListener("change", (event) => { state.supplier = event.target.value; state.selectedCategory = null; renderSpend(); renderPortfolio(); });
    el("contract-filter").addEventListener("change", (event) => { state.contract = event.target.value; state.selectedCategory = null; renderSpend(); renderPortfolio(); });
    el("commercial-search").addEventListener("input", (event) => { state.commercialSearch = event.target.value; state.selectedCategory = null; renderSpend(); renderPortfolio(); });
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
    document.querySelectorAll("#forecast-mode-control button").forEach((button) => button.addEventListener("click", () => {
      state.forecastMode = button.dataset.forecastMode;
      document.querySelectorAll("#forecast-mode-control button").forEach((item) => {
        const active = item === button;
        item.classList.toggle("active", active);
        item.setAttribute("aria-pressed", String(active));
      });
      renderForecast();
    }));
    document.querySelectorAll("#pareto-control button").forEach((button) => button.addEventListener("click", () => {
      state.paretoDimension = button.dataset.pareto;
      document.querySelectorAll("#pareto-control button").forEach((item) => {
        const active = item === button;
        item.classList.toggle("active", active);
        item.setAttribute("aria-pressed", String(active));
      });
      renderSpend();
    }));
    document.querySelectorAll("#segmentation-control button").forEach((button) => button.addEventListener("click", () => {
      state.segmentationMode = button.dataset.segmentation;
      document.querySelectorAll("#segmentation-control button").forEach((item) => {
        const active = item === button;
        item.classList.toggle("active", active);
        item.setAttribute("aria-pressed", String(active));
      });
      renderSegmentation();
    }));
    el("export-plan").addEventListener("click", exportPlan);
    el("export-forecast").addEventListener("click", exportForecast);
    el("export-spend").addEventListener("click", exportSpend);
    el("export-audit").addEventListener("click", () => downloadCsv("demand-genie-v4-planner-action-log.csv", state.actionLog));
    el("dialog-cancel").addEventListener("click", () => el("action-dialog").close());
    el("action-form").addEventListener("submit", confirmAction);
  }

  function initialize() {
    state.sourceMode = sheet("All_Model_Forecasts").length ? "packaged" : "browser";
    bindEvents();
    updateForecastModelOptions();
    renderAll();
    el("source-kind").textContent = state.data.sourceType;
    el("source-file").textContent = state.data.fileName;
    el("source-freshness").textContent = `Snapshot ${snapshotDate()}`;
  }

  initialize();
})();
