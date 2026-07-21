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

## Development

The current MVP is a dependency-free, interactive dashboard in [index.html](index.html). It runs locally without a build step or server:

1. Open `index.html` in a modern browser.
2. Use the product-family, part, and forecast-model controls to change the monitoring view.
3. Select a row in the exception queue to inspect that part's historical demand and forecast context.
4. Download the current exception queue as CSV when needed.

The dashboard currently uses deterministic sample manufacturing data embedded in the file. It is intentionally structured so an ERP or planning-system extract can replace the sample dataset in a later iteration.

### Dashboard Design

The dashboard follows the monitoring and visual-communication principles of Stephen Few's *Information Dashboard Design*:

- A single desktop-screen overview for rapid monitoring.
- Context-rich comparisons rather than isolated KPI tiles.
- Line charts for time, compact forecast comparisons, and a ranked exception queue for follow-up work.
- Flat, restrained visual treatment with color reserved for meaningful planning signals.

The reusable Codex skill `few-dashboard` contains the implementation workflow and primary-source research notes used for these choices.

## Build Log

See [docs/BUILD_LOG.md](docs/BUILD_LOG.md) for timestamped project notes and Codex collaboration history.

## License

This project is licensed under GPL-3.0. See [LICENSE](LICENSE).
