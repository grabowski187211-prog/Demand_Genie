# OpenAI Build Week Checklist

Source of truth: https://openai.devpost.com/

## Required before submission

- [ ] Confirm the entrant type.
- [ ] Confirm country of residence and eligibility.
- [x] Select exactly one category:
  - Apps for Your Life
  - **Work & Productivity — selected**
  - Developer Tools
  - Education
- [x] Finish a working project that runs consistently on its target platform.
- [x] Add complete setup and run instructions to `README.md`.
- [x] Add sample data or test credentials if the app needs them.
- [x] Document how Codex accelerated the workflow.
- [x] Document where key product, engineering, and design decisions were made.
- [x] Document how GPT-5.6 and Codex contributed to the final result.
- [x] Record the `/feedback` Codex Session ID.
- [ ] Create a public YouTube demo video under 3 minutes.
- [ ] Make sure the video includes audio explaining the project and Codex/GPT-5.6 usage.
- [ ] Provide the repository URL in the Devpost submission.
- [ ] If the repository is private, share it with `testing@devpost.com` and `build-week-event@openai.com`.
- [x] Provide a working demo URL or judging instructions if available.
- [ ] Submit on Devpost before July 21, 2026 at 5:00 PM PT.

## Current project facts

- Project name: Demand Genie
- Repository: https://github.com/grabowski187211-prog/Demand_Genie
- Category: Work & Productivity
- Current release: Version 5
- Product scope: An offline manufacturing-planning workbench combining forecast evaluation, time-series decomposition, DDMRP replenishment, procurement spend, supplier risk, Kraljic analysis, execution exceptions, buffer governance, and demand segmentation.
- Runtime: Self-contained `dashboard-v5.html`; no server, account, API key, or network connection is required.
- Sample data: `data/Demand_Genie_Synthetic_Portfolio_v5.xlsx`
- Judging instructions: Follow the Quick Start in `README.md`; open `dashboard-v5.html` and optionally exercise the **Load workbook** flow with the V5 sample workbook.
- Hosted demo URL: Not recorded in the repository; the offline dashboard and judging instructions satisfy the runnable-demo path.
- Demo video URL: Not recorded in the repository; add the public YouTube URL before submission.
- Demo narration: `docs/Demand_Genie_V5_Video_Narration.html` and `docs/Demand_Genie_V5_Video_Narration.pdf`
- `/feedback` session ID: `019f83b6-8ade-7751-bd4e-69a13f3d6f36`

## Decisions already made

- Keep Demand Genie in the **Work & Productivity** category.
- Ship Version 5 as the judging version while preserving Versions 1-4 as implementation history.
- Use a self-contained offline HTML application rather than an application server or hosted dependency.
- Use synthetic data only; do not imply that the sample contains real suppliers, products, contracts, or transactions.
- Keep replenishment releases, forecast-driven buffer changes, and commercial opportunities as human review decisions.
- Accept packaged forecast evidence only when the uploaded workbook SHA-256 matches the analyzed workbook; otherwise show clearly labeled browser benchmarks.
- Do not claim booked savings, live ERP integration, automatic purchasing, or browser execution of ARIMA, TiRex2, or robust STL.

## Remaining submission actions

- Confirm entrant type, country of residence, and eligibility.
- Record and publish the narrated YouTube demo under three minutes.
- Add the final public video URL to this checklist and the Devpost submission.
- Add the repository URL to the Devpost submission.
- Perform a final clean-clone run-through using the README Quick Start.
- Submit on Devpost before the deadline.
