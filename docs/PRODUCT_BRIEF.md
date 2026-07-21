# Demand Genie Product Brief

## Working idea

Demand Genie is a demand-planning workspace for manufacturing planners. It turns demand history, forecast performance, and planning exceptions into a clear monitoring view that supports material, purchasing, staffing, capacity, and supplier decisions.

## Target user

Manufacturing supply-chain planners and demand planners who work across incomplete ERP data, spreadsheet-based planning, changing sales signals, and high-pressure S&OP decisions.

## Problem

Planners are often accountable when service or inventory fails but are given fragmented, late, or politically inflated forecast information. ERP systems may store the data without providing the visual context needed to assess forecast reliability, bias, demand direction, and supply exposure quickly.

## Core workflow

1. Select a product family, part, and planning signal.
2. Compare historical demand with the selected forecast and its near-term horizon.
3. Review WAPE, forecast bias, and exception status with their defined tolerances.
4. Prioritize parts by forecast variance and supply exposure.
5. Drill into an exception, then export the ranked queue for the next planning conversation.

## MVP features

- Standalone HTML dashboard with no runtime dependencies.
- Product-family, part, and forecast-model filtering.
- Actual-versus-forecast time series with a visible forecast horizon.
- Forecast WAPE and bias with explicit tolerance context.
- Forecast-model comparison.
- Ranked planning exception queue with per-part demand patterns.
- CSV export of the displayed exception queue.

## Non-goals

- Direct ERP integration in the first release.
- Automatic purchasing, inventory, or production decisions.
- A black-box forecast that hides accuracy, bias, or historical context.

## Hackathon category

Work & Productivity
