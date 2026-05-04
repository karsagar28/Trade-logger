# Trade Logger

A lightweight trade journal built as a Google Apps Script web app backed by a Google Sheet. The UI is a single-page ledger for logging trades, reviewing R-multiple performance, and scanning notes or counterfactual analysis from trade history.

## Project Structure

- `Code.gs` contains the Google Apps Script backend. It reads and writes trade rows in the `Trades` sheet.
- `index.html` contains the full frontend: layout, styles, form logic, charts, and calls to the Apps Script web app URL.

## Trade Fields

Each trade records:

- Date
- Ticker
- Type: shares or options
- Strategy type: Breakout, Pullback-generic, Pullback-30min, or Other
- Risk: always saved as `1R`
- Stop type
- Result in R
- Close type and reason
- Compliance flag
- Analysis/comment text
- Alternate result for non-compliant trades

When `Compliant trade` is unchecked, the form shows an `Alternate result` input and treats the analysis box as a counterfactual note for what may have happened if the trade had followed the rules.

## Local Preview

Open `index.html` directly in a browser to preview the UI:

```text
file:///path/to/Trade-logger/index.html
```

The local page can connect to the live Google Sheet if the deployed Apps Script URL is saved in the settings panel.

## Google Apps Script Setup

1. Open the Google Sheet used for the journal.
2. Go to `Extensions` -> `Apps Script`.
3. Replace the Apps Script project's `Code.gs` contents with this repo's `Code.gs`.
4. If the Apps Script project has an `index.html` file, replace its contents with this repo's `index.html`.
5. Click `Deploy` -> `Manage deployments`.
6. Edit the existing web app deployment.
7. Select `New version`, then deploy.

Updating the existing deployment keeps the same web app URL. Creating a new deployment may require pasting the new URL into the app settings again.

## Using the App

1. Open the deployed web app URL or the local `index.html` file.
2. If prompted, paste the Apps Script deployment URL into settings.
3. Click `+` to add a trade.
4. Enter ticker, type, strategy type, stop type, optional result, and compliance details.
5. Use `Compliant trade` for rule-following trades.
6. Uncheck `Compliant trade` for rule breaks, then enter the alternate result and counterfactual analysis.
7. Hover over trade rows to quickly preview saved comments or counterfactual analysis.
8. Click a trade row to edit it.

## Updating the App

After changing `Code.gs` or `index.html` locally:

1. Copy the updated file contents into the matching Apps Script project files.
2. Deploy a new version from `Deploy` -> `Manage deployments`.
3. Refresh the app in the browser.

Google Apps Script does not automatically pull changes from this repository, so the copy/paste deployment step is required unless a separate Apps Script sync workflow is added.

## Data Notes

The app stores data in a sheet named `Trades`. If it does not exist, the backend creates it and adds the expected headers. Existing sheets are updated with newer header names such as `Compliant`, `Analysis`, `Alternate Result`, and `Strategy Type`.
