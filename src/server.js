const express = require("express");
const dotenv = require("dotenv");
const http = require("http");
const https = require("https");
const { URL } = require("url");
const { buildSheetsClient } = require("./sheets");
const { runImport } = require("./workflow");
const { startPolling, getPollingStatus } = require("./poller");

dotenv.config();

const app = express();
app.use(express.json({ limit: "2mb" }));

const port = process.env.PORT || 3000;
let pollingState = null;
let pingIntervalId = null;

app.get("/health", (_req, res) => {
  res.json({ ok: true });
});

app.get("/", (_req, res) => {
  res.json({ ok: true });
});

app.get("/polling/status", (_req, res) => {
  res.json({
    ok: true,
    jobs: getPollingStatus(pollingState),
    statePath: pollingState ? pollingState.statePath : null,
  });
});

app.post("/import", async (req, res) => {
  try {
    const {
      source,
      destination,
      removeColumns,
      keepColumns,
      headerRowIndex = 0,
      clearDestination = true,
    } = req.body || {};

    if (!source || !destination) {
      return res.status(400).json({
        error: "Missing source or destination object.",
      });
    }

    const {
      spreadsheetId: sourceSpreadsheetId,
      range: sourceRange,
      ranges: sourceRanges,
      gid: sourceGid,
      sheetName: sourceSheetName,
    } = source;
    const {
      spreadsheetId: destinationSpreadsheetId,
      sheetName: destinationSheetName,
      gid: destinationGid,
      startCell = "A1",
    } = destination;

    if (
      !sourceSpreadsheetId ||
      (!sourceRange &&
        (!Array.isArray(sourceRanges) || sourceRanges.length === 0))
    ) {
      return res.status(400).json({
        error:
          "source.spreadsheetId and source.range (or source.ranges) are required.",
      });
    }

    if (!destinationSpreadsheetId || (!destinationSheetName && !destinationGid)) {
      return res.status(400).json({
        error:
          "destination.spreadsheetId and destination.sheetName (or destination.gid) are required.",
      });
    }

    const sheets = await buildSheetsClient();
    const updated = await runImport({
      sheets,
      source: {
        spreadsheetId: sourceSpreadsheetId,
        range: sourceRange,
        ranges: sourceRanges,
        gid: sourceGid,
        sheetName: sourceSheetName,
      },
      destination: {
        spreadsheetId: destinationSpreadsheetId,
        sheetName: destinationSheetName,
        gid: destinationGid,
        startCell,
      },
      removeColumns,
      keepColumns,
      headerRowIndex,
      clearDestination,
    });

    return res.json({
      ok: true,
      updatedRange: updated.updatedRange,
      updatedRows: updated.updatedRows,
      updatedColumns: updated.updatedColumns,
      updatedCells: updated.updatedCells,
    });
  } catch (error) {
    const message =
      error && typeof error.message === "string"
        ? error.message
        : "Unknown error";
    return res.status(500).json({ error: message });
  }
});

async function bootstrap() {
  const sheets = await buildSheetsClient();
  app.locals.sheets = sheets;

  const jobsEnv = process.env.POLL_JOBS_JSON;
  if (jobsEnv) {
    try {
      const jobs = JSON.parse(jobsEnv);
      const intervalMs = Number.parseInt(
        process.env.POLL_INTERVAL_MS || "60000",
        10
      );
      pollingState = startPolling({
        sheets,
        jobs: Array.isArray(jobs) ? jobs : [],
        defaultIntervalMs: Number.isNaN(intervalMs) ? 60000 : intervalMs,
        statePath: process.env.POLL_STATE_PATH,
      });
      if (pollingState.count > 0) {
        console.log(`[poller] Started ${pollingState.count} polling job(s).`);
      }
    } catch (error) {
      const message =
        error && typeof error.message === "string"
          ? error.message
          : "Unknown error";
      console.error(`[poller] Failed to start: ${message}`);
    }
  }

  app.listen(port, () => {
    console.log(`Workflow server listening on port ${port}`);
    startPing();
  });
}

bootstrap();

function startPing() {
  if (pingIntervalId) {
    return;
  }

  const urlValue = process.env.PING_URL;
  if (!urlValue) {
    return;
  }

  const intervalMs = Number.parseInt(
    process.env.PING_INTERVAL_MS || "600000",
    10
  );

  if (!Number.isFinite(intervalMs) || intervalMs <= 0) {
    console.warn("[ping] Invalid PING_INTERVAL_MS; skipping pings.");
    return;
  }

  let pingUrl;
  try {
    pingUrl = new URL(urlValue);
  } catch (error) {
    console.warn("[ping] Invalid PING_URL; skipping pings.");
    return;
  }

  const client = pingUrl.protocol === "https:" ? https : http;
  pingIntervalId = setInterval(() => {
    const request = client.get(pingUrl, (response) => {
      response.resume();
    });
    request.on("error", (error) => {
      const message =
        error && typeof error.message === "string"
          ? error.message
          : "Unknown error";
      console.warn(`[ping] Failed: ${message}`);
    });
  }, intervalMs);

  console.log(`[ping] Enabled ${pingUrl.toString()} every ${intervalMs}ms.`);
}
