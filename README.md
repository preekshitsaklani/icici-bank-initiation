# ICICI Bank: initiating coverage

**HOLD, 12-month target price ₹1,380** — 5% upside from ₹1,310.6 (NSE close, 1 Oct 2026). Bear case ₹900, bull case ₹1,740.

ICICI Bank is a best-in-class franchise — FY26 RoA of 2.2%, net NPA of 0.33%, CET-1 of 16.35% — but the current price already pays for that quality. RoE has drifted down from 18.7% (FY24) to 16.1% (FY26) as capital built up, and the stock trades at 2.0x Sep-27E standalone book after stripping out ₹173/share of listed subsidiaries, implying 9.7% long-run growth versus a 10.0% assumption in the model.

**Data as of 1 October 2026.** Prices, beta, peer multiples and subsidiary market caps are NSE close, 1 Oct 2026, via Yahoo Finance (`yfinance`). The RBI's next MPC decision was due 7 October 2026 — after this note's data date — so treat anything rate-sensitive as pre-decision.

## Method

Target price is a justified price-to-book on the standalone bank, plus listed subsidiaries (ICICI Prudential AMC, ICICI Lombard, ICICI Prudential Life) marked at their own market value, less a 20% holding-company discount. The standalone multiple is derived from a Gordon-growth justified P/B (RoE, cost of equity, long-term growth), not a peer-multiple read-across. Bear/base/bull scenarios flex RoE, credit cost and the growth assumption; the companion Excel model and its `RunScenarios` VBA macro recompute all three.

**ICICI Lombard's ~51% stake is labelled "approx." in the note and model — check it against the latest NSE shareholding pattern before relying on it.**

## Files

- `ICICI_Initiation_Note.pdf` — the full note: summary, thesis, financials, target-price build, downside/upside risks, methodology and sources.
- `ICICI_Bank_Initiation_Model.xlsx` — the live model (`Cover`, `Inputs`, `Historical`, `Forecast`, `Valuation`, `Peers`, `Sources` sheets), reconciled to ICICI Bank's audited standalone results for FY22–FY26.
- `ScenarioRunner.bas` — VBA macro that runs the Bear/Base/Bull scenarios and writes target price, upside and rating back to the `Cover` sheet.

## Running the scenario macro

1. Open `ICICI_Bank_Initiation_Model.xlsx` in Excel (Windows).
2. Alt+F11 → File → Import File → `ScenarioRunner.bas`.
3. Save the workbook as `.xlsm`.
4. Run `RunScenarios` — it cycles the named range `ScenarioSelector` through Bear/Base/Bull and records `TargetPrice`, `Upside`, `Rating`, `PAT_FY28E` and `RoE_FY28E` for each into `ScenarioOut`.

## Sources

ICICI Bank quarterly/annual results (Q4 FY23 through Q4/FY26), Q1 FY27 standalone results and earnings-call highlights, RBI MPC minutes and weekly statistical supplement, India 10-year G-sec yield (Trading Economics), ICICI Prudential Life/AMC stake disclosures (Upstox, Business Standard), and prices/peer data via Yahoo Finance (`yfinance`). Full citations with dates are in the note's Sources section.

## Disclaimer

This note is an independent student research project prepared for learning and portfolio purposes. It is not investment advice and not a recommendation to buy or sell any security. The author is not a SEBI-registered research analyst.
