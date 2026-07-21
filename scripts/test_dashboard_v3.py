"""Browser smoke test for the standalone Demand Genie version 3 dashboard."""

from __future__ import annotations

import json
import sys
from pathlib import Path

from playwright.sync_api import sync_playwright


ROOT = Path(__file__).resolve().parents[1]
URL = sys.argv[1] if len(sys.argv) > 1 else "http://localhost:8001/dashboard-v3.html"
WORKBOOK = ROOT / "data" / "Demand_Genie_Synthetic_Demand_History.xlsx"


def verify_page(page, viewport_name: str) -> dict[str, object]:
    page.goto(URL, wait_until="load")
    page.wait_for_selector("#planning-body tr")
    assert page.locator(".tab").count() == 7
    assert page.locator("#planning-body tr").count() == 80
    assert page.locator("#source-file").inner_text() == WORKBOOK.name

    for tab_name, panel_id in (
        ("Forecast", "view-forecast"),
        ("Spend", "view-spend"),
        ("Portfolio", "view-portfolio"),
        ("Execution", "view-execution"),
        ("Buffers", "view-buffers"),
        ("Data", "view-data"),
        ("Replenishment", "view-planning"),
    ):
        page.get_by_role("tab", name=tab_name).click()
        assert page.locator(f"#{panel_id}").is_visible()

    page.get_by_role("tab", name="Spend").click()
    page.wait_for_selector("#pareto-chart .pareto-bar")
    assert page.locator("#pareto-chart .pareto-bar").count() == 10
    assert page.locator("#opportunity-body tr").count() >= 1
    initial_spend = page.locator("#net-spend").inner_text()
    page.select_option("#spend-period-filter", "Current")
    assert page.locator("#net-spend").inner_text() != initial_spend
    page.locator('#pareto-control button[data-pareto="supplier"]').click()
    assert "Supplier" in page.locator("#pareto-title").inner_text()
    with page.expect_download() as download_info:
        page.locator("#export-spend").click()
    assert download_info.value.suggested_filename == "demand-genie-v3-spend-analysis.csv"

    page.get_by_role("tab", name="Portfolio").click()
    page.wait_for_selector("#kraljic-chart .portfolio-point")
    assert page.locator("#kraljic-chart .portfolio-point").count() == 10
    assert page.locator("#segmentation-chart .segmentation-point").count() == 80
    page.locator('#segmentation-control button[data-segmentation="classic"]').click()
    assert "ABC/XYZ" in page.locator("#segmentation-title").inner_text()

    page.get_by_role("tab", name="Data").click()
    assert page.locator("#spend-line-count").inner_text() == "1,964"
    assert "complete" in page.locator("#data-headline").inner_text().lower()

    page.set_input_files("#workbook-input", str(WORKBOOK))
    page.wait_for_function("document.querySelector('#source-kind').textContent.trim() === 'Uploaded workbook'")
    page.wait_for_function("document.querySelector('#spend-line-count').textContent.trim() === '1,964'")
    assert page.locator("#source-kind").inner_text() == "Uploaded workbook"

    overflow = page.evaluate("document.documentElement.scrollWidth - document.documentElement.clientWidth")
    assert overflow <= 2, f"{viewport_name} horizontal overflow: {overflow}px"
    return {
        "viewport": viewport_name,
        "tabs": page.locator(".tab").count(),
        "spend_lines": page.locator("#spend-line-count").inner_text(),
        "horizontal_overflow_px": overflow,
    }


def main() -> None:
    errors: list[str] = []
    results = []
    with sync_playwright() as playwright:
        browser = playwright.chromium.launch()
        screenshots = []
        for name, viewport in (
            ("desktop", {"width": 1440, "height": 1000}),
            ("mobile", {"width": 390, "height": 844}),
        ):
            page = browser.new_page(viewport=viewport)
            page.on("console", lambda message: errors.append(f"console {message.type}: {message.text}") if message.type == "error" else None)
            page.on("pageerror", lambda error: errors.append(f"pageerror: {error}"))
            results.append(verify_page(page, name))
            page.get_by_role("tab", name="Spend").click()
            spend_screenshot = f"/tmp/demand-genie-v3-{name}-spend.png"
            page.screenshot(path=spend_screenshot, full_page=True)
            screenshots.append(spend_screenshot)
            page.get_by_role("tab", name="Portfolio").click()
            portfolio_screenshot = f"/tmp/demand-genie-v3-{name}-portfolio.png"
            page.screenshot(path=portfolio_screenshot, full_page=True)
            screenshots.append(portfolio_screenshot)
            page.close()
        browser.close()
    assert not errors, "\n".join(errors)
    print(json.dumps({"status": "PASS", "results": results, "screenshots": screenshots}, indent=2))


if __name__ == "__main__":
    main()
