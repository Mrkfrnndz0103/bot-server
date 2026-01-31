const {
  ensureSheetExists,
  readValues,
  writeValues,
  clearRange,
  removeColumnsFromRows,
  keepColumnsFromRows,
  getSheetTitleById,
} = require("./sheets");
const { updateDashboard } = require("./dashboard");

function hasSheetPrefix(range) {
  return typeof range === "string" && range.includes("!");
}

async function resolveSheetName(sheets, spreadsheetId, sheetName, gid) {
  if (sheetName) {
    return sheetName;
  }
  if (gid !== undefined && gid !== null) {
    return getSheetTitleById(sheets, spreadsheetId, gid);
  }
  throw new Error("Missing sheetName or gid.");
}

function normalizeRange(range, sheetName) {
  if (hasSheetPrefix(range)) {
    return range;
  }
  return `${sheetName}!${range}`;
}

function columnLettersToIndex(letters) {
  if (!letters) {
    return null;
  }
  const normalized = letters.toUpperCase();
  let index = 0;
  for (let i = 0; i < normalized.length; i += 1) {
    const code = normalized.charCodeAt(i);
    if (code < 65 || code > 90) {
      return null;
    }
    index = index * 26 + (code - 64);
  }
  return index;
}

function columnIndexToLetters(index) {
  if (!Number.isInteger(index) || index <= 0) {
    return null;
  }
  let value = index;
  let letters = "";
  while (value > 0) {
    const remainder = (value - 1) % 26;
    letters = String.fromCharCode(65 + remainder) + letters;
    value = Math.floor((value - 1) / 26);
  }
  return letters;
}

function parseCellRef(cellRef) {
  const raw = String(cellRef || "A1");
  const withoutSheet = raw.includes("!") ? raw.split("!").pop() : raw;
  const match = withoutSheet.match(/^([A-Za-z]+)(\d+)?$/);
  if (!match) {
    return { colIndex: 1, rowIndex: 1 };
  }
  const colIndex = columnLettersToIndex(match[1]) || 1;
  const rowIndex = match[2] ? Number.parseInt(match[2], 10) : 1;
  return {
    colIndex,
    rowIndex: Number.isFinite(rowIndex) && rowIndex > 0 ? rowIndex : 1,
  };
}

function inferClearRange({
  sheetName,
  startCell,
  rows,
  keepColumns,
}) {
  const rowCount = Array.isArray(rows) ? rows.length : 0;
  let colCount = 0;
  if (Array.isArray(rows) && rows.length > 0) {
    colCount = rows.reduce((max, row) => {
      if (!Array.isArray(row)) {
        return max;
      }
      return Math.max(max, row.length);
    }, 0);
  } else if (Array.isArray(keepColumns) && keepColumns.length > 0) {
    colCount = keepColumns.length;
  }

  if (rowCount === 0 || colCount === 0) {
    return null;
  }

  const { colIndex, rowIndex } = parseCellRef(startCell || "A1");
  const endCol = columnIndexToLetters(colIndex + colCount - 1);
  const endRow = rowIndex + rowCount - 1;
  if (!endCol) {
    return null;
  }
  return `${sheetName}!${columnIndexToLetters(colIndex)}${rowIndex}:${endCol}${endRow}`;
}

function extractRanges(source) {
  if (Array.isArray(source.ranges) && source.ranges.length > 0) {
    return source.ranges;
  }
  if (source.range) {
    return [source.range];
  }
  return [];
}

async function resolveSourceRanges(sheets, source) {
  const ranges = extractRanges(source);
  if (ranges.length === 0) {
    throw new Error("Missing source range.");
  }

  if (ranges.every((range) => hasSheetPrefix(range))) {
    const importRange = source.importRange || ranges[0];
    if (hasSheetPrefix(importRange)) {
      return {
        ranges,
        importRange,
        sheetName: null,
      };
    }
    const fallbackSheet = String(ranges[0]).split("!")[0];
    return {
      ranges,
      importRange: normalizeRange(importRange, fallbackSheet),
      sheetName: null,
    };
  }

  const sheetName = await resolveSheetName(
    sheets,
    source.spreadsheetId,
    source.sheetName,
    source.gid
  );

  return {
    ranges: ranges.map((range) => normalizeRange(range, sheetName)),
    importRange: normalizeRange(source.importRange || ranges[0], sheetName),
    sheetName,
  };
}

async function importRows({
  sheets,
  rows,
  source,
  destination,
  removeColumns,
  keepColumns,
  headerRowIndex = 0,
  clearDestination = true,
}) {
  let filteredRows = rows;

  if (Array.isArray(keepColumns) && keepColumns.length > 0) {
    filteredRows = keepColumnsFromRows(rows, keepColumns, { headerRowIndex });
  } else {
    filteredRows = removeColumnsFromRows(
      rows,
      Array.isArray(removeColumns) ? removeColumns : []
    );
  }

  let destinationSheetName = destination.sheetName;
  if (!destinationSheetName) {
    destinationSheetName = await resolveSheetName(
      sheets,
      destination.spreadsheetId,
      destination.sheetName,
      destination.gid
    );
  }

  await ensureSheetExists(
    sheets,
    destination.spreadsheetId,
    destinationSheetName
  );

  if (clearDestination) {
    if (destination.clearRange) {
      const rangeToClear = normalizeRange(
        destination.clearRange,
        destinationSheetName
      );
      await clearRange(sheets, destination.spreadsheetId, rangeToClear);
    } else {
      const inferredRange = inferClearRange({
        sheetName: destinationSheetName,
        startCell: destination.startCell || "A1",
        rows: filteredRows,
        keepColumns,
      });
      if (inferredRange) {
        await clearRange(sheets, destination.spreadsheetId, inferredRange);
      } else {
        console.warn(
          "[import] clearDestination requested but range could not be inferred; skipping clear to preserve formulas."
        );
      }
    }
  }

  const writeResult = await writeValues(
    sheets,
    destination.spreadsheetId,
    destinationSheetName,
    destination.startCell || "A1",
    filteredRows
  );

  const dashboardConfig =
    destination && typeof destination.dashboard === "object"
      ? destination.dashboard
      : destination && destination.dashboard === true
      ? { sheetName: "Dashboard" }
      : null;
  const dashboardSheetName =
    (destination && destination.dashboardSheetName) ||
    (dashboardConfig && dashboardConfig.sheetName);
  const dashboardEnabled =
    dashboardSheetName && (dashboardConfig ? dashboardConfig.enabled !== false : true);

  if (dashboardEnabled) {
    try {
      await updateDashboard({
        sheets,
        spreadsheetId: destination.spreadsheetId,
        dataSheetName: destinationSheetName,
        dashboardSheetName,
        headerRowIndex,
      });
    } catch (error) {
      const message =
        error && typeof error.message === "string"
          ? error.message
          : "Unknown error";
      console.warn(`[dashboard] Failed to update: ${message}`);
    }
  }

  return writeResult;
}

async function runImport({
  sheets,
  source,
  destination,
  removeColumns,
  keepColumns,
  headerRowIndex = 0,
  clearDestination = true,
}) {
  const resolved = await resolveSourceRanges(sheets, source);
  const rows = await readValues(
    sheets,
    source.spreadsheetId,
    resolved.importRange
  );
  return importRows({
    sheets,
    rows,
    source,
    destination,
    removeColumns,
    keepColumns,
    headerRowIndex,
    clearDestination,
  });
}

module.exports = {
  importRows,
  runImport,
  resolveSheetName,
  resolveSourceRanges,
};
