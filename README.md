# Demand Genie

Demand Genie is an OpenAI Build Week hackathon project.

The working product scope is still being defined. This repository is now set up to track the implementation, submission requirements, and Codex-assisted build history in one place.

## Hackathon

- Event: OpenAI Build Week
- Deadline: July 21, 2026 at 5:00 PM PT
- Devpost: https://openai.devpost.com/
- Repository: https://github.com/grabowski187211-prog/Demand_Genie
- Expected category: TBD

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

Open `dashboard-v3.html` directly in a modern browser. No application server or network connection is required. Load the sample Excel workbook from the dashboard to exercise the same upload path used for an ERP or procurement extract.

Version 3 includes seven operational views: Replenishment, Forecast, Spend, Portfolio, Execution, Buffers, and Data. Commercial exports are review queues and do not claim booked savings.

## Sample Data

[Demand_Genie_Synthetic_Demand_History.xlsx](data/Demand_Genie_Synthetic_Demand_History.xlsx) contains deterministic synthetic planning and procurement data. No row represents a real product, supplier, contract, or transaction.

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

Build the self-contained dashboard after changing source files or data:

```bash
python3 scripts/build_dashboard_v3.py
```

Optional browser verification requires Python Playwright with Chromium and a local server:

```bash
python3 -m http.server 8001
python3 scripts/test_dashboard_v3.py http://localhost:8001/dashboard-v3.html
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
