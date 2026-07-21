# Demand Genie

Demand Genie is an OpenAI Build Week Work & Productivity project: a standalone decision-support workbench for manufacturing planners. It combines forecast evaluation, time-series decomposition, DDMRP replenishment, procurement spend, supplier risk, and demand segmentation in one offline HTML dashboard.

The project also explores the democratization of data and code. Its premise is that planners should be able to turn their own operational knowledge into professional, maintainable tools instead of waiting months or years for an ERP provider to deliver a generic solution.

## Quick Start

### Run the dashboard

1. Clone or download this repository.
2. Open [`dashboard-v5.html`](dashboard-v5.html) in a current version of Chrome, Edge, Firefox, or Safari.
3. Explore the seven views: **Replenishment**, **Forecast**, **Spend**, **Portfolio**, **Execution**, **Buffers**, and **Data**.

That is the complete runtime setup. Version 5 is a self-contained HTML application with its JavaScript, styling, sample data, forecast results, and spreadsheet reader embedded. It needs no package installation, application server, account, API key, network connection, or access to OpenAI services.

If your browser restricts local files, serve the repository with Python and open <http://localhost:8001/dashboard-v5.html>:

```bash
python3 -m http.server 8001
```

### Exercise the Excel upload workflow

1. In the dashboard, select **Load workbook**.
2. Choose [`data/Demand_Genie_Synthetic_Portfolio_v5.xlsx`](data/Demand_Genie_Synthetic_Portfolio_v5.xlsx).
3. Confirm that the dashboard reports the uploaded workbook and keeps the full packaged ARIMA, ETS, and TiRex2 evidence available.

The bundled workbook is deterministic synthetic data; it contains no real products, suppliers, contracts, or transactions. Its SHA-256 matches the packaged analysis. A different workbook can also be loaded, but the dashboard intentionally falls back to clearly labeled browser benchmarks when no content-matched offline forecast package is available.

If the generated HTML is not present, install `openpyxl` and rebuild it from the committed source and data artifacts:

```bash
python3 -m pip install openpyxl
python3 scripts/build_dashboard_v5.py
```

The sections below document the sample data, full analytical rebuild, validation, audit, and optional browser-test workflows.

## Hackathon

- Event: OpenAI Build Week
- Deadline: July 21, 2026 at 5:00 PM PT
- Devpost: https://openai.devpost.com/
- Repository: https://github.com/grabowski187211-prog/Demand_Genie
- Category: Work & Productivity

## Built with Codex and GPT-5.6

Demand Genie was built as an iterative collaboration between a supply-chain practitioner and Codex, powered by GPT-5.6. It was not generated from a single prompt. I remained the product owner and domain reviewer: I defined the planner problem, selected the planning methods, set the usability and deployment constraints, reviewed each version, and challenged results that did not look credible. GPT-5.6 helped translate that direction into implementation plans, reason across forecasting and supply-chain concepts, and critique the work. Codex then carried that reasoning through the repository by researching, writing code, running analysis, testing in a browser, auditing calculations, and maintaining the Git history.

### How we divided the work

| Area | My product, engineering, or design decision | How GPT-5.6 and Codex accelerated the work |
| --- | --- | --- |
| Product direction | I framed Demand Genie around problems I have seen in manufacturing: noisy demand, disconnected spreadsheets, biased forecasts, and planners held accountable without usable decision support. I also made democratization central to the idea—supply-chain professionals should be able to build tools around their own workflows. | GPT-5.6 helped turn the domain narrative into a focused Work & Productivity product and a planner journey built around monitoring, exceptions, drill-down, and action. |
| Delivery model | I chose a self-contained HTML dashboard with Excel upload, no application server, and no runtime network dependency. I also required each major iteration to remain available instead of being overwritten. | Codex implemented the browser application, embedded its dependencies and validated artifacts, generated the upload workbooks, and preserved Versions 1-5 as an inspectable product history. |
| Information design | I required the dashboards to follow Stephen Few's information-dashboard principles and to avoid decorative or misleading analytics. I chose an exception-first workbench rather than a wall of KPI cards. | Codex researched the primary guidance, encoded it in the reusable `few-dashboard` skill, and applied a restrained hierarchy, compact comparisons, meaningful color, ranked queues, and responsive layouts. |
| Planning methods | I chose forecasting, DDMRP, spend analysis, supplier exposure, Kraljic strategy, and modern item segmentation as parts of one planning workflow. I explicitly asked that forecasts inform planning without becoming unexplained automatic decisions. | GPT-5.6 synthesized the disciplines while Codex created reusable `fpp3`, `ddmrp`, and `spend-analysis` workflows, generated a coherent 80-SKU dataset, and connected the calculations across the dashboard views. |
| Forecast engineering | I asked for real model comparison, TiRex2, decomposition, and decision-grade accuracy—not a visually convincing mockup. | Codex built R and Python pipelines for seven classical and foundation-model candidates, 23,520 rolling-origin predictions, training-only scaled errors, per-SKU selection, forecast intervals, and robust STL decomposition. It ran the official TiRex2 model locally and packaged the resulting artifacts for offline use. |
| Quality and realism | I requested a devil's-advocate review and authorized parallel specialist agents. Later, I noticed that the original ADI/CV2 view looked unrealistically uniform and asked whether the calculation or the synthetic data was wrong. | The review found material weaknesses, including stale packaged results being applied to modified uploads and missing inputs being treated as zero. Codex hardened V4 with exact SHA-256 provenance, explicit calculation gates, and independent controls. It then traced the segmentation issue to an overly smooth fixture—not a chart bug—and built V5 with realistic smooth, erratic, intermittent, and lumpy histories. |
| Human oversight | I kept the final say on scope, realism, and what the dashboard is allowed to claim. Commercial outputs remain review queues, and replenishment recommendations remain decision support. | GPT-5.6 and Codex made rapid exploration practical while also exposing assumptions, warnings, data provenance, backtest evidence, and audit results so that the human reviewer can challenge the output. |

### Where Codex changed the pace

Codex compressed a workflow that normally spans separate research, data-engineering, forecasting, front-end, QA, and documentation efforts into a tight build-review loop. Within the project thread it scaffolded and connected the GitHub repository, researched and created four reusable domain skills, generated and validated the Excel data model, built five dashboard iterations, ran R and Python forecasting pipelines, executed the local TiRex2 checkpoint, performed multi-agent adversarial reviews, tested desktop and mobile behavior with Playwright, and committed each milestone. When tests or reviews exposed a weakness, Codex followed the evidence through the data, calculation, interface, and documentation layers instead of only patching the visible symptom.

GPT-5.6 was especially valuable for long-horizon reasoning across those layers: keeping product intent, statistical validity, supply-chain rules, provenance, and interface behavior consistent while the implementation grew. Codex supplied the agentic execution environment around that reasoning—the ability to inspect files, use the shell and browser, edit and run code, coordinate specialist reviews, and verify the result. The combination let me work at the level of product direction and domain judgment while still remaining directly involved in every important decision.

The primary Codex project thread is session `019f83b6-8ade-7751-bd4e-69a13f3d6f36`. A timestamped technical record is available in [the build log](docs/BUILD_LOG.md), and the independent V5 calculation results are in [the machine-readable audit](data/audit-v5.json).

## Development Journey

### Overview

Demand Genie evolved through five evidence-driven iterations, moving from a focused demand-monitoring prototype to a decision-support workbench spanning forecasting, replenishment, procurement, supplier risk, execution, and buffer governance. Each version responded to a concrete planner need or a weakness exposed through calculation audits, adversarial review, browser testing, or hands-on product review.

The journey also shows how GPT-5.6 and Codex accelerated development without replacing human judgment. GPT-5.6 helped reason across product, forecasting, and supply-chain decisions; Codex implemented and tested those decisions throughout the repository. The supply-chain practitioner remained responsible for the problem definition, method selection, realism checks, automation boundaries, and final acceptance of each iteration.

- **Version 1 — Demand-monitoring MVP:** established the exception-first dashboard, demand history, forecast comparisons, accuracy and bias context, drill-down, and CSV export.
- **Version 2 — Forecasting and replenishment:** added Excel upload, multi-model forecasting, DDMRP buffers, net-flow positions, and order recommendations.
- **Version 3 — Connected planning workbench:** brought procurement spend, supplier exposure, Kraljic strategy, execution risks, and ABC/XYZ plus ADI/CV2 segmentation into the same seven-view workflow.
- **Version 4 — Decision-grade controls:** replaced the single holdout with rolling-origin evaluation, ran TiRex2 locally, added robust STL decomposition, enforced exact SHA-256 forecast provenance, and introduced explicit calculation gates and independent audits.
- **Version 5 — Realistic demand behavior:** traced an unrealistic segmentation result to overly smooth sample data, rebuilt the 80-SKU portfolio across smooth, erratic, intermittent, and lumpy patterns, and recalculated the complete planning and forecast package.

Each iteration was preserved rather than overwritten so judges can inspect how testing, planner review, and Codex-assisted analysis changed the product. See the [timestamped build log](docs/BUILD_LOG.md) for the detailed implementation record.

## Submission Requirements

The Devpost submission requires:

- A working project built with Codex and GPT-5.6.
- One selected category: Apps for Your Life, Work & Productivity, Developer Tools, or Education.
- A project description explaining features and functionality.
- A public YouTube demo video under 3 minutes with audio covering what was built and how Codex and GPT-5.6 were used.
- A code repository URL for judging and testing.
- Setup instructions, sample data if needed, and clear run instructions in this README.
- A `/feedback` Codex Session ID for the project thread where most core functionality was built.

## Dashboards

Each iteration remains available as a standalone HTML file:

- [Version 1](index.html): demand-monitoring MVP.
- [Version 2](dashboard-v2.html): forecasting and DDMRP planning workbench.
- [Version 3](dashboard-v3.html): forecasting, DDMRP, procurement spend, supplier exposure, Kraljic strategy, and modern item segmentation.
- [Version 4](dashboard-v4.html): independently audited forecasting and planning workbench with real TiRex2 inference, rolling-origin selection, robust STL decomposition, strict upload provenance, and visible decision gates.
- [Version 5](dashboard-v5.html): realistic item-level demand behavior across all four ADI/CV2 quadrants, recalculated DDMRP and forecasts, monotone TiRex2 quantiles, and an exception-oriented segmentation queue.

Open `dashboard-v5.html` directly in a modern browser. No application server or network connection is required. Load the sample Excel workbook from the dashboard to exercise the same upload path used for an ERP or procurement extract.

Version 5 includes seven operational views: Replenishment, Forecast, Spend, Portfolio, Execution, Buffers, and Data. Commercial exports are review queues and do not claim booked savings. Versions 1-4 remain unchanged and available for comparison.

Packaged forecast artifacts are accepted only when the uploaded workbook SHA-256 matches the analyzed sample workbook. Other uploads use clearly labeled browser benchmarks; they do not claim to have run ARIMA or TiRex2 in the browser.

## Sample Data

[Demand_Genie_Synthetic_Demand_History.xlsx](data/Demand_Genie_Synthetic_Demand_History.xlsx) contains deterministic synthetic planning and procurement data. No row represents a real product, supplier, contract, or transaction.

[Demand_Genie_Synthetic_Portfolio_v5.xlsx](data/Demand_Genie_Synthetic_Portfolio_v5.xlsx) is the current realistic behavior fixture. It preserves the same 80-SKU commercial scope while adding validated zero-demand occurrence patterns and nonzero-demand size variation: 30 smooth, 18 erratic, 17 intermittent, and 15 lumpy SKUs.

- 10 product groups with 8 SKUs each.
- 2,880 monthly demand records from July 2023 through June 2026.
- 80 complete DDMRP item-location positions with inventory, open supply, qualified demand, order rules, buffer settings, and recommendations.
- 1,964 signed procurement lines from July 2024 through June 2026, including visible credits.
- 38 normalized supplier parents, 20 active contracts, an approved category taxonomy, and category-level commercial risk evidence.
- A separate `Demand_History` consumption fact for usage-value ABC/XYZ and Syntetos-Boylan ADI/CV2 analysis. Purchase orders are not treated as demand.
- A `Spend_Control` sheet that reconciles signed net spend in EUR to the synthetic AP control total.

### Rebuild and Validate

The workbook requires Python and `openpyxl`. Regenerate and validate it with:

```bash
python3 scripts/generate_synthetic_demand_data.py
python3 scripts/analyze_spend.py data/Demand_Genie_Synthetic_Demand_History.xlsx data/spend-analysis --control-total 143686430.21
python3 scripts/validate_workbook_v3.py
```

Build and audit the self-contained V5 dashboard after changing source files or data:

```bash
python3 scripts/generate_realistic_demand_v5.py
python3 scripts/analyze_spend.py data/Demand_Genie_Synthetic_Portfolio_v5.xlsx data/spend-analysis-v5 --control-total 143686430.21
python3 scripts/test_spend_analysis.py
python3 scripts/build_dashboard_v5.py
python3 scripts/audit_project_v5.py
```

The independent audit recomputes demand patterns, DDMRP zones and orders, signed spend and Pareto outputs, ABC/XYZ and ADI/CV2 segmentation, 23,520 rolling-origin predictions, model selection, interval ordering, STL identities, artifact hashes, and the immutable hashes of dashboards V1-V4. Its machine-readable result is [data/audit-v5.json](data/audit-v5.json).

### Rebuild Forecasts

Classical forecasts require R with `fpp3`, `readxl`, and `readr`. TiRex2 0.1.1 requires Python 3.11-3.13 and first-download access to the gated `NX-AI/TiRex-2` Hugging Face model. The generated CSV artifacts are bundled into the standalone dashboard, so dashboard users do not need R, Python, model access, or a network connection.

```bash
python3.13 -m venv .venv
.venv/bin/pip install -r requirements-tirex2.txt
Rscript scripts/run_forecast_analysis_v4.R data/Demand_Genie_Synthetic_Portfolio_v5.xlsx data/forecast-v5
.venv/bin/python scripts/run_tirex2_forecasts.py data/Demand_Genie_Synthetic_Portfolio_v5.xlsx data/forecast-v5
.venv/bin/python scripts/merge_forecast_analysis_v4.py data/forecast-v5 --input data/Demand_Genie_Synthetic_Portfolio_v5.xlsx
python3 scripts/build_dashboard_v5.py
python3 scripts/audit_project_v5.py
```

V5 evaluates Mean, Naive, Drift, Seasonal Naive, ETS, ARIMA, and TiRex2 over seven expanding-window origins with six months per origin. It selects per-SKU RMSE over 42 out-of-sample points, preferring the least-complex candidate within 2% of the minimum RMSE. MASE and RMSSE scales use only each origin's training history. TiRex2 marginal quantiles are non-negative and monotonically rearranged per horizon; raw crossings and rearranged points remain recorded in the artifacts.

Optional browser verification requires Python Playwright with Chromium and a local server:

```bash
python3 -m http.server 8001
python3 scripts/test_dashboard_v5.py http://localhost:8001/dashboard-v5.html
```

### Dashboard Design

The dashboards follow the monitoring and visual-communication principles of Stephen Few's *Information Dashboard Design*:

- A single desktop-screen overview for rapid monitoring.
- Context-rich comparisons rather than isolated KPI tiles.
- Line charts for time, compact forecast comparisons, and a ranked exception queue for follow-up work.
- Flat, restrained visual treatment with color reserved for meaningful planning signals.
- Ranked spend bars with a Pareto line, a true Kraljic bubble scatter, explicit review queues, and visible data-reconciliation status.

The reusable Codex skills `few-dashboard`, `fpp3`, `ddmrp`, and `spend-analysis` contain the implementation workflows used for these choices.

## Build Log

See [docs/BUILD_LOG.md](docs/BUILD_LOG.md) for timestamped project notes and Codex collaboration history.

## License

This project is licensed under GPL-3.0. See [LICENSE](LICENSE).
