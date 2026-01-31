const crypto = require("crypto");
const fs = require("fs");
const path = require("path");
const { readValues, batchReadValues } = require("./sheets");
const { importRows, resolveSheetName, resolveSourceRanges } = require("./workflow");

function hashRows(payload) {
  const json = JSON.stringify(payload || []);
  return crypto.createHash("sha256").update(json).digest("hex");
}

function buildStatePath(statePath) {
  if (statePath) {
    return statePath;
  }
  return path.join(process.cwd(), "data", "poller-state.json");
}

function loadState(statePath) {
  try {
    if (!fs.existsSync(statePath)) {
      return {};
    }
    const raw = fs.readFileSync(statePath, "utf8");
    return JSON.parse(raw);
  } catch (_error) {
    return {};
  }
}

function saveState(statePath, state) {
  try {
    const dir = path.dirname(statePath);
    fs.mkdirSync(dir, { recursive: true });
    fs.writeFileSync(statePath, JSON.stringify(state, null, 2));
  } catch (_error) {
    // Best-effort persistence.
  }
}

function normalizeRanges(source) {
  if (Array.isArray(source.ranges) && source.ranges.length > 0) {
    return source.ranges;
  }
  if (source.range) {
    return [source.range];
  }
  return [];
}

function resolveImportRange(source) {
  if (source.importRange) {
    return source.importRange;
  }
  const ranges = normalizeRanges(source);
  return ranges[0] || null;
}

function normalizeJob(job) {
  if (!job || typeof job !== "object") {
    return null;
  }

  const source = job.source || {};
  const destination = job.destination || {};
  const ranges = normalizeRanges(source);

  if (!source.spreadsheetId || ranges.length === 0) {
    return null;
  }

  if (
    !destination.spreadsheetId ||
    (!destination.sheetName && destination.gid === undefined)
  ) {
    return null;
  }

  const sourceKeySuffix =
    source.gid !== undefined && source.gid !== null
      ? `:${source.gid}`
      : source.sheetName
      ? `:${source.sheetName}`
      : "";
  const sourceKey = `${source.spreadsheetId}:${ranges.join("|")}${sourceKeySuffix}`;

  return {
    source: {
      spreadsheetId: source.spreadsheetId,
      range: source.range,
      ranges,
      importRange: source.importRange,
      gid: source.gid,
      sheetName: source.sheetName,
    },
    destination: {
      spreadsheetId: destination.spreadsheetId,
      sheetName: destination.sheetName,
      gid: destination.gid,
      startCell: destination.startCell || "A1",
      clearRange: destination.clearRange,
      dashboard: destination.dashboard,
      dashboardSheetName: destination.dashboardSheetName,
    },
    removeColumns: Array.isArray(job.removeColumns) ? job.removeColumns : [],
    keepColumns: Array.isArray(job.keepColumns) ? job.keepColumns : [],
    headerRowIndex: Number.isInteger(job.headerRowIndex)
      ? job.headerRowIndex
      : 0,
    clearDestination:
      typeof job.clearDestination === "boolean" ? job.clearDestination : true,
    pollIntervalMs:
      Number.isInteger(job.pollIntervalMs) && job.pollIntervalMs > 0
        ? job.pollIntervalMs
        : null,
    jobName: job.jobName || sourceKey,
    jobKey: job.jobKey || sourceKey,
  };
}

function startPolling({
  sheets,
  jobs,
  defaultIntervalMs = 60000,
  statePath,
}) {
  const normalizedJobs = (jobs || [])
    .map(normalizeJob)
    .filter(Boolean)
    .map((job) => ({
      ...job,
      lastHash: null,
      lastRunAt: null,
      lastUpdatedAt: null,
      lastError: null,
      inFlight: false,
    }));

  const resolvedStatePath = buildStatePath(statePath);
  const persistedState = loadState(resolvedStatePath);

  normalizedJobs.forEach((job) => {
    const persisted = persistedState[job.jobKey];
    if (persisted && typeof persisted.lastHash === "string") {
      job.lastHash = persisted.lastHash;
      job.lastRunAt = persisted.lastRunAt || null;
      job.lastUpdatedAt = persisted.lastUpdatedAt || null;
      job.lastError = persisted.lastError || null;
    }
  });

  function persistJobState(job) {
    const snapshot = loadState(resolvedStatePath);
    snapshot[job.jobKey] = {
      lastHash: job.lastHash,
      lastRunAt: job.lastRunAt,
      lastUpdatedAt: job.lastUpdatedAt,
      lastError: job.lastError,
    };
    saveState(resolvedStatePath, snapshot);
  }

  normalizedJobs.forEach((job) => {
    const interval = job.pollIntervalMs || defaultIntervalMs;
    setInterval(async () => {
      if (job.inFlight) {
        return;
      }
      job.inFlight = true;
      try {
        const resolved = await resolveSourceRanges(sheets, job.source);
        const ranges = resolved.ranges;
        let rangeValues = [];
        if (ranges.length === 1) {
          const rows = await readValues(
            sheets,
            job.source.spreadsheetId,
            ranges[0]
          );
          rangeValues = [rows];
        } else {
          rangeValues = await batchReadValues(
            sheets,
            job.source.spreadsheetId,
            ranges
          );
        }

        const nextHash = hashRows({
          ranges,
          values: rangeValues,
        });
        job.lastRunAt = new Date().toISOString();

        if (job.lastHash && job.lastHash === nextHash) {
          job.lastError = null;
          persistJobState(job);
          return;
        }

        const importRange = resolved.importRange || resolveImportRange(job.source);
        let importIndex = ranges.indexOf(importRange);
        if (importIndex < 0 && ranges.length > 0) {
          importIndex = 0;
          if (importRange) {
            console.warn(
              `[poller] ${job.jobName} importRange not found; using ${ranges[0]}`
            );
          }
        }
        const rows = rangeValues[importIndex] || [];

        const destinationSheetName = await resolveSheetName(
          sheets,
          job.destination.spreadsheetId,
          job.destination.sheetName,
          job.destination.gid
        );

        const updated = await importRows({
          sheets,
          rows,
          source: job.source,
          destination: {
            ...job.destination,
            sheetName: destinationSheetName,
          },
          removeColumns: job.removeColumns,
          keepColumns: job.keepColumns,
          headerRowIndex: job.headerRowIndex,
          clearDestination: job.clearDestination,
        });

        job.lastHash = nextHash;
        job.lastUpdatedAt = new Date().toISOString();
        job.lastError = null;
        persistJobState(job);

        const updatedRows =
          typeof updated.updatedRows === "number" ? updated.updatedRows : "?";
        console.log(
          `[poller] Updated ${job.jobName} (${updatedRows} rows) at ${new Date().toISOString()}`
        );
      } catch (error) {
        const message =
          error && typeof error.message === "string"
            ? error.message
            : "Unknown error";
        job.lastError = message;
        job.lastRunAt = new Date().toISOString();
        persistJobState(job);
        console.error(`[poller] ${job.jobName} failed: ${message}`);
      } finally {
        job.inFlight = false;
      }
    }, interval);
  });

  return {
    count: normalizedJobs.length,
    jobs: normalizedJobs,
    statePath: resolvedStatePath,
  };
}

function getPollingStatus(pollingState) {
  if (!pollingState || !Array.isArray(pollingState.jobs)) {
    return [];
  }

  return pollingState.jobs.map((job) => ({
    jobName: job.jobName,
    jobKey: job.jobKey,
    source: job.source,
    destination: job.destination,
    pollIntervalMs: job.pollIntervalMs,
    lastHash: job.lastHash,
    lastRunAt: job.lastRunAt,
    lastUpdatedAt: job.lastUpdatedAt,
    lastError: job.lastError,
  }));
}

module.exports = {
  startPolling,
  getPollingStatus,
};
