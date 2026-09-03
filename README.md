# Stock Review

Local web app for monthly Forecast Orders review — the same job as the National Stock Review **DETAILS** tab, not a Cursor canvas.

Reads FileMaker StockCalculator JSON from:

`\\fs\apps\SalesForecast\Data-DatabaseInfoSchema\StockCalculator\{STATE}-{YYYY-MM}.json`

## Run

```
npm install
npm start
```

Open http://localhost:5174

Product Profitability stays on 5173. This app uses **5174**.

## First screen

- State + month from the live calculator files
- On Hand, In Service, **Total Stock**
- Current / target utilisation, current orders, surplus / shortfall
- Order qty, interstate suggestions, system requested
- **Decision** — type a qty; cost is landed × decision (saved locally)
- Family totals + Excel export (DETAILS + TOTALS)

## Not this project

This is not Product Profitability and not the Sales forecast review chat workspace. Those stay where they are.
