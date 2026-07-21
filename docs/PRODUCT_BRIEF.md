# Demand Genie Product Brief

## Product definition

Demand Genie is an offline decision-support workbench for manufacturing planners. It connects demand history, forecast evidence, DDMRP replenishment, procurement spend, supplier risk, execution exceptions, buffer governance, and item segmentation in one self-contained HTML dashboard.

The current judging release is **Version 5** in the **Work & Productivity** category.

## Target user

Manufacturing supply-chain, demand, inventory, materials, and procurement planners who work across fragmented ERP extracts, spreadsheet-based planning, changing sales signals, and high-pressure S&OP decisions.

## Problem

Planners are often accountable when service, inventory, or supply fails while the relevant evidence is split across forecasts, stock files, purchase orders, supplier reports, and spreadsheets. ERP systems may store the transactions without providing the context needed to assess forecast reliability, replenishment logic, demand behavior, supply exposure, and commercial concentration quickly.

## Product promise

Demand Genie helps a planner move from **monitoring**, to **exception**, to **evidence**, to **reviewed action** without hiding the calculation or silently automating a business decision.

## Key product decisions

- **Exception first:** prioritize items and commercial questions that need attention instead of presenting a wall of isolated KPI cards.
- **Human in control:** forecasts inform planning, but they do not silently alter qualified demand, approve buffer changes, release orders, or claim savings.
- **Visible evidence:** show forecast backtests, bias, intervals, decomposition limits, DDMRP inputs, spend reconciliation, provenance, and decision gates.
- **Offline delivery:** package the application, sample data, forecast results, and spreadsheet reader into a standalone HTML file with no runtime network dependency.
- **Honest model boundaries:** use full ARIMA, ETS, TiRex2, and robust STL evidence only for the exact analyzed workbook; label browser-only benchmarks for other uploads.
- **Synthetic demonstration:** provide realistic but fully synthetic data covering 80 SKUs and all smooth, erratic, intermittent, and lumpy demand patterns.
- **Inspectable evolution:** preserve Versions 1-4 and use Version 5 as the current judging release.

## Core workflow

1. Start in **Replenishment** and review ranked net-flow, order, and execution exceptions.
2. Select an SKU and inspect on-hand inventory, eligible open supply, qualified demand, buffer zones, order rules, and the recommended action.
3. Open **Forecast** to compare seven candidates on rolling-origin evidence, including bias, error, intervals, and the selected model.
4. Use decomposition evidence to understand trend, seasonality, and unexplained variation without overstating three years of history.
5. Review **Spend** and **Portfolio** for concentration, supplier exposure, Kraljic positioning, ABC/XYZ, and ADI/CV2 demand behavior.
6. Use **Execution** and **Buffers** to examine late supply, approved planning settings, ownership, and review dates.
7. Finish in **Data** to confirm workbook completeness, reconciliation, forecast provenance, calculation gates, and warnings.
8. Export a review queue or action log for the next planner conversation; keep the final business decision with the user.

## Version 5 capabilities

- Seven operational views: Replenishment, Forecast, Spend, Portfolio, Execution, Buffers, and Data.
- Excel upload using the same path a planner would use for an ERP or procurement extract.
- DDMRP zones, net-flow positions, order recommendations, demand-adjustment proposals, and explicit release/hold review actions.
- Seven-model forecast comparison over seven expanding-window origins and 42 out-of-sample points per SKU.
- Classical Mean, Naive, Drift, Seasonal Naive, ETS, and ARIMA candidates plus TiRex2 foundation-model forecasts.
- Forecast intervals, training-only scaled errors, per-SKU selection, robust STL decomposition, and visible evidence limits.
- Signed-spend reconciliation, category and supplier Pareto analysis, contract and concentration evidence, and commercial review queues.
- Kraljic portfolio analysis, ABC/XYZ segmentation, and Syntetos-Boylan ADI/CV2 patterns across all four quadrants.
- Strict missing-input gates and exact SHA-256 matching between the workbook and packaged forecast artifacts.
- CSV exports for replenishment, forecast, spend, segmentation, and planner action-log workflows.
- Deterministic synthetic demonstration data containing 80 SKUs, 2,880 monthly demand records, and 1,964 signed procurement lines.

## Delivery and validation

- Runs by opening `dashboard-v5.html` in a modern browser; a local server is optional.
- Requires no account, API key, application server, OpenAI connection, or runtime network access.
- Includes a V5 machine-readable audit covering 23 calculation and provenance controls.
- Includes repeatable workbook, spend, audit, desktop, and mobile browser tests.
- Preserves the earlier dashboard files as a visible build history.

## Non-goals

- Live ERP integration in the hackathon release.
- Automatic purchasing, inventory, or production decisions.
- A black-box forecast that hides accuracy, bias, or historical context.
- Claims of booked savings from analytical opportunity flags.
- Claims that ARIMA, TiRex2, or robust STL run in the browser for arbitrary uploads.
- Production deployment, identity management, workflow approvals, or source-system reconciliation beyond the documented synthetic demonstration.

## Hackathon category

Work & Productivity

Demand Genie improves a planner's daily work by reducing the time needed to assemble evidence, identify exceptions, challenge a forecast or recommendation, and prepare a traceable human decision.

## Build approach

The supply-chain practitioner remained the product owner and domain reviewer. GPT-5.6 supported cross-domain reasoning and critique; Codex accelerated repository research, implementation, data generation, forecasting pipelines, browser testing, independent audits, documentation, and iterative correction. The detailed division of work and timestamped evidence are recorded in `README.md` and `docs/BUILD_LOG.md`.
