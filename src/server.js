const express = require("express");
const dotenv = require("dotenv");
const http = require("http");
const https = require("https");
const { URL } = require("url");
const path = require("path");
const { buildSheetsClient, getSheetTitleById, readValues } = require("./sheets");
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
  res.redirect("/dashboard");
});

app.get("/dashboard", (_req, res) => {
  res.sendFile(path.join(__dirname, "..", "public", "dashboard.html"));
});

app.use("/public", express.static(path.join(__dirname, "..", "public")));

app.get("/polling/status", (_req, res) => {
  res.json({
    ok: true,
    jobs: getPollingStatus(pollingState),
    statePath: pollingState ? pollingState.statePath : null,
  });
});

app.get("/api/regional-validation", async (_req, res) => {
  try {
    const sheets = app.locals.sheets || (await buildSheetsClient());

    const spreadsheetId = process.env.PIVOT_SPREADSHEET_ID;
    const gidValue = process.env.PIVOT_GID;
    const pivotRange = process.env.PIVOT_RANGE;

    if (!spreadsheetId || !gidValue || !pivotRange) {
      return res.status(400).json({
        error: "Missing PIVOT_SPREADSHEET_ID, PIVOT_GID, or PIVOT_RANGE.",
      });
    }

    const sheetTitle = await getSheetTitleById(
      sheets,
      spreadsheetId,
      Number(gidValue)
    );

    const range = pivotRange.includes("!")
      ? pivotRange
      : `${sheetTitle}!${pivotRange}`;

    const values = await readValues(sheets, spreadsheetId, range);
    const headers = values[0] || [];
    const rows = values.slice(1);

    return res.json({
      ok: true,
      headers,
      rows,
      range,
    });
  } catch (error) {
    const message =
      error && typeof error.message === "string"
        ? error.message
        : "Unknown error";
    return res.status(500).json({ error: message });
  }
});

app.get("/api/stuckup-analysis", async (_req, res) => {
  try {
    const sheets = app.locals.sheets || (await buildSheetsClient());

    const spreadsheetId = process.env.PIVOT_SPREADSHEET_ID;
    const gidValue = process.env.PIVOT_GID;
    const pivotRange = process.env.PIVOT_STUCKUP_RANGE;

    if (!spreadsheetId || !gidValue || !pivotRange) {
      return res.status(400).json({
        error:
          "Missing PIVOT_SPREADSHEET_ID, PIVOT_GID, or PIVOT_STUCKUP_RANGE.",
      });
    }

    const sheetTitle = await getSheetTitleById(
      sheets,
      spreadsheetId,
      Number(gidValue)
    );

    const range = pivotRange.includes("!")
      ? pivotRange
      : `${sheetTitle}!${pivotRange}`;

    const values = await readValues(sheets, spreadsheetId, range);
    const headers = values[0] || [];
    const rows = values.slice(1);

    return res.json({
      ok: true,
      headers,
      rows,
      range,
    });
  } catch (error) {
    const message =
      error && typeof error.message === "string"
        ? error.message
        : "Unknown error";
    return res.status(500).json({ error: message });
  }
});

function buildPivotHandler(envKey) {
  return async (_req, res) => {
    try {
      const sheets = app.locals.sheets || (await buildSheetsClient());

      const spreadsheetId = process.env.PIVOT_SPREADSHEET_ID;
      const gidValue = process.env.PIVOT_GID;
      const pivotRange = process.env[envKey];

      if (!spreadsheetId || !gidValue || !pivotRange) {
        return res.status(400).json({
          error: `Missing PIVOT_SPREADSHEET_ID, PIVOT_GID, or ${envKey}.`,
        });
      }

      const sheetTitle = await getSheetTitleById(
        sheets,
        spreadsheetId,
        Number(gidValue)
      );

      const range = pivotRange.includes("!")
        ? pivotRange
        : `${sheetTitle}!${pivotRange}`;

      const values = await readValues(sheets, spreadsheetId, range);
      const headers = values[0] || [];
      const rows = values.slice(1);

      return res.json({
        ok: true,
        headers,
        rows,
        range,
      });
    } catch (error) {
      const message =
        error && typeof error.message === "string"
          ? error.message
          : "Unknown error";
      return res.status(500).json({ error: message });
    }
  };
}

app.get("/api/ageing-bucket", buildPivotHandler("PIVOT_AGEING_RANGE"));
app.get("/api/top-hubs", buildPivotHandler("PIVOT_TOP_HUBS_RANGE"));
app.get("/api/validation-trend", buildPivotHandler("PIVOT_VALIDATION_TREND_RANGE"));
app.get("/api/stuckup-trend", buildPivotHandler("PIVOT_STUCKUP_TREND_RANGE"));

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
