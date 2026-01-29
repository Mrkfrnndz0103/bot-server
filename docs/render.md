# Render

This repo includes `render.yaml` to deploy a web service. It also provisions a
small persistent disk for polling state.

Required env vars to set in Render:

- `GOOGLE_APPLICATION_CREDENTIALS` (or `GOOGLE_SERVICE_ACCOUNT_JSON`)
- `POLL_JOBS_JSON`

Optional env vars:

- `POLL_INTERVAL_MS`
- `POLL_STATE_PATH`
- `PING_URL`
- `PING_INTERVAL_MS`
