# Configuration

## Environment

Copy `.env.example` to `.env`, then set:

- `GOOGLE_APPLICATION_CREDENTIALS` to a service account JSON path
- or `GOOGLE_SERVICE_ACCOUNT_JSON` to inline JSON

If the service account file is in the project root, use:
`GOOGLE_APPLICATION_CREDENTIALS=./service-account.json`

## Required sharing

Share both source and destination spreadsheets with the service account email.
