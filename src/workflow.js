const {
  ensureSheetExists,
  readValues,
  writeValues,
  clearSheet,
  clearRange,
  removeColumnsFromRows,
  keepColumnsFromRows,
  getSheetTitleById,
} = require("./sheets");

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
      await clearSheet(
        sheets,
        destination.spreadsheetId,
        destinationSheetName
      );
    }
  }

  return writeValues(
    sheets,
    destination.spreadsheetId,
    destinationSheetName,
    destination.startCell || "A1",
    filteredRows
  );
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
