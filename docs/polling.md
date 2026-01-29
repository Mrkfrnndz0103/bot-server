# Polling

Set `POLL_JOBS_JSON` to an array of jobs. Each job uses the same shape as
`POST /import`, plus:

- `pollIntervalMs` (optional) per-job interval
- `jobName` (optional) label for logs
- `source.ranges` (optional array) to watch multiple ranges
- `source.importRange` (optional) which watched range to import (defaults to the
  first range)

Example:

```json
[
  {
    "jobName": "source-to-dest",
    "source": {
      "spreadsheetId": "SOURCE_ID",
      "gid": 0,
      "ranges": ["A:Z", "AA:AF"],
      "importRange": "A:Z"
    },
    "destination": {
      "spreadsheetId": "DEST_ID",
      "gid": 0,
      "startCell": "A1"
    },
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
    "clearDestination": true,
    "pollIntervalMs": 60000
  }
]
```

Polling state is persisted to `data/poller-state.json` by default. Override via
`POLL_STATE_PATH`.
