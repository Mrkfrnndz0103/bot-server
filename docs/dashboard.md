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
