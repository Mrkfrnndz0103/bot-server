# Dashboard

## Regional Validation Summary

The dashboard reads pivot data from a Google Sheet and renders it as a heatmap
table. Configure the pivot location using env vars:

- `PIVOT_SPREADSHEET_ID`
- `PIVOT_GID`
- `PIVOT_RANGE` (example: `A1:J9`)

The endpoint used by the UI:

- `GET /api/regional-validation`

The dashboard page:

- `GET /dashboard`

## Stuck Up Tagging Analysis

Configured via:

- `PIVOT_STUCKUP_RANGE` (example: `A12:E16`)

Endpoint:

- `GET /api/stuckup-analysis`

## Ageing Bucket Analysis

Configured via:

- `PIVOT_AGEING_RANGE`

Endpoint:

- `GET /api/ageing-bucket`

## Top Hubs

Configured via:

- `PIVOT_TOP_HUBS_RANGE`

Endpoint:

- `GET /api/top-hubs`

## 20hrs - 1d Validation Trend

Configured via:

- `PIVOT_VALIDATION_TREND_RANGE`

Endpoint:

- `GET /api/validation-trend`

## Stuck Up Tagging Trend

Configured via:

- `PIVOT_STUCKUP_TREND_RANGE`

Endpoint:

- `GET /api/stuckup-trend`
