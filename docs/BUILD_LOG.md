# Build Log

This file tracks timestamped project work for OpenAI Build Week.

## 2026-07-21 10:12 CEST

- Cloned `https://github.com/grabowski187211-prog/Demand_Genie` into the local workspace.
- Confirmed the repository uses the `main` branch and is connected to GitHub.
- Reviewed OpenAI Build Week rules and submission requirements through the Devpost Hackathons connector.
- Added initial repository documentation for hackathon tracking.

## 2026-07-21 13:22 CEST

- Researched Stephen Few's dashboard-design guidance from Perceptual Edge primary sources, including dashboard purpose, single-screen monitoring, chart selection, color discipline, sparklines, and bullet graphs.
- Created and validated the reusable global Codex skill `few-dashboard`, with a concise workflow and source reference notes.
- Built the first Demand Genie MVP as a dependency-free interactive `index.html` dashboard using deterministic sample manufacturing demand data.
- Implemented filters, actual-versus-forecast monitoring, model comparison, WAPE and bias context, a ranked exception queue, part-level drill-in, and CSV export.
- Verified JavaScript parsing and inspected the dashboard through local desktop and narrow-viewport browser screenshots.
- Used Few's principles to favor a one-screen monitoring hierarchy, flat visual treatment, direct comparisons, and color reserved for meaningful planning signals.

## 2026-07-21 13:33 CEST

- Added a deterministic synthetic demand-history generator and workbook for the planned Excel-upload workflow.
- Generated 36 complete months of data from July 2023 through June 2026 across 10 product groups and 80 SKUs.
- Included a tidy `Demand_Data` upload sheet with 2,880 monthly records, plus a business-friendly demand matrix, product master, and data dictionary.
- Validated product-group count, SKU count, monthly coverage, and positive demand values after workbook generation.

## 2026-07-21 15:55 CEST

- Preserved `index.html` and `dashboard-v2.html` as the first two working iterations.
- Extended the synthetic workbook with 1,964 signed procurement lines, 38 normalized supplier parents, 20 contracts, category taxonomy, explicit commercial-risk factors, an EUR control total, and a separate demand-history fact for item segmentation.
- Verified that the original `Demand_Data` and `DDMRP_Recommendations` outputs remained unchanged while adding the commercial data model.
- Ran the reusable `spend-analysis` workflow. The analysis reconciled exactly to EUR 143,686,430.21, passed all blocking data-quality gates, and produced supplier/category Pareto, contract, concentration, Kraljic, and ABC/XYZ plus ADI/CV2 outputs.
- Built `dashboard-v3.html` as a standalone seven-view supply-chain workbench with dynamic commercial filters, spend and supplier Pareto analysis, review prompts with no savings claim, a Kraljic bubble portfolio, and modern item segmentation.
- Added repeatable workbook and Playwright tests. Version 3 passed desktop and mobile browser verification with workbook re-upload, export, interaction, chart-rendering, console-error, and page-overflow checks.

## 2026-07-21 17:05 CEST

- Preserved the byte-level hashes of Versions 1-3 and built `dashboard-v4.html` as a separate iteration.
- Replaced the single holdout with seven expanding-window origins and 42 out-of-sample errors per SKU/model. Recomputed ME, RMSE, MAE, MASE, RMSSE, WAPE, bias, and empirical interval coverage from training-only scales.
- Ran the official TiRex2 0.1.1 package and `NX-AI/TiRex-2` checkpoint locally for all 80 SKUs. TiRex2 won raw RMSE for 18 SKUs and was selected for 15 after the declared 2% simplicity rule.
- Added robust monthly STL decomposition with aligned observed, trend, seasonal, and remainder small multiples plus explicit three-cycle evidence limits.
- Added exact workbook/artifact SHA-256 provenance, strict missing-input and EUR 0.01 decision gates, and visible warnings for synthetic-data and policy limitations.
- Added an independent audit that passed 22 calculation controls across demand, DDMRP, spend, segmentation, forecasts, intervals, selection, decomposition, and file provenance. Desktop/mobile Playwright tests also passed corrupted-upload, blocked-DDMRP, keyboard, rendering, and overflow cases.

## 2026-07-21 17:40 CEST

- Confirmed the original ADI/CV2 display was mathematically correct but exposed a weak synthetic fixture: all 80 histories had positive demand every month, forcing ADI to 1 and placing every item in the smooth quadrant.
- Preserved Dashboard V4 and created a separate V5 workbook with validated item-level occurrence and demand-size patterns: 30 smooth, 18 erratic, 17 intermittent, and 15 lumpy SKUs. ADI now spans 1.0-3.6 and CV2 spans 0.002-1.661.
- Rebuilt all DDMRP source facts, buffer settings, net-flow recommendations, spend segmentation, classical rolling-origin forecasts, TiRex2 forecasts, and robust STL components against the V5 histories.
- Detected 38 raw TiRex2 quantile crossings across 28 backtest points. Added declared non-negative clipping and monotone rearrangement, retained raw crossing evidence, and verified zero unresolved crossings.
- Built `dashboard-v5.html` with a visible pattern mix, all four ADI/CV2 quadrants, threshold labels, and an exception queue ranked lumpy, intermittent, erratic, then smooth.
- Passed 23 independent calculation controls plus desktop/mobile rendering, exact-hash upload, corrupted-upload, missing-DDMRP-input, export, keyboard, and overflow tests.

## 2026-07-21 22:44 CEST

- Finalized the project name as Demand Genie, the category as Work & Productivity, and Version 5 as the judging release.
- Added a judge-oriented README Quick Start covering direct offline launch, an optional local server, the exact sample workbook, upload behavior, and regeneration of the standalone HTML.
- Reconciled the hackathon checklist with decisions already evidenced in the repository and separated completed requirements from external submission actions.
- Updated the product brief from the original forecast-monitoring MVP to the implemented seven-view planning workbench.

## Codex and GPT-5.6 collaboration summary

- The supply-chain practitioner remained the product owner and domain reviewer: they defined the problem, selected the planning disciplines, set the offline and human-oversight constraints, reviewed each release, and challenged results that did not look credible.
- GPT-5.6 helped translate domain direction into implementation plans, reason across forecasting, DDMRP, procurement, segmentation, provenance, and information design, and critique whether the combined workflow remained coherent.
- Codex accelerated repository execution: it researched primary guidance, created reusable domain workflows, generated and validated synthetic workbooks, built five dashboard iterations, ran R and Python forecast pipelines, executed TiRex2 locally, tested browser behavior, and maintained the audit trail.
- The review loop changed implementation decisions. Adversarial checks led to exact-hash forecast provenance and strict missing-input gates in Version 4. A planner challenge to the uniform ADI/CV2 display identified an unrealistic fixture and led to the four-pattern Version 5 dataset.
- Human review remains explicit in the product. Order releases, forecast-driven buffer proposals, and commercial opportunities are review queues rather than autonomous or booked decisions.
- Supporting evidence is available in the repository: implementation history in this build log, setup and role attribution in `README.md`, and independent calculation results in `data/audit-v5.json`.
