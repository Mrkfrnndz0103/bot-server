# API

## POST /import

Request body:

```json
{
  "source": {
    "spreadsheetId": "SOURCE_ID",
    "gid": 0,
    "range": "A:Z"
  },
  "destination": {
    "spreadsheetId": "DEST_ID",
    "gid": 0,
    "startCell": "A1"
  },
  "removeColumns": [2, 5],
  "keepColumns": [
    "Date",
    "spx_station_site",
    "shipment_id",
    "status_group",
    "status_desc",
    "status_timestamp",
    "hub_dest_station_name",
    "fms_last_update_time",
    "last_run_time",
    "cogs",
    "day",
    "Ageing bucket_",
    "operator"
  ],
  "headerRowIndex": 0,
  "clearDestination": true
}
```

Notes:

- `removeColumns` uses zero-based column indexes.
- `keepColumns` takes precedence and must match header names exactly.
- `gid` can be used instead of a sheet name; the server resolves it at runtime.

## GET /polling/status

Returns the current polling jobs and last run/update timestamps.
