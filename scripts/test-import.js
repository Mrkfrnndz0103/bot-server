const dotenv = require("dotenv");
const { buildSheetsClient, readValues } = require("../src/sheets");
const { runImport, resolveSheetName } = require("../src/workflow");

dotenv.config();

function parseJsonEnv(name) {
  const raw = process.env[name];
  if (!raw) {
    return undefined;
  }
  try {
    return JSON.parse(raw);
  } catch (error) {
    throw new Error(`Invalid JSON in ${name}.`);
  }
}

function parseBoolEnv(name, fallback) {
  const raw = process.env[name];
  if (raw === undefined || raw === "") {
    return fallback;
  }
  if (/^(true|1|yes)$/i.test(raw)) {
    return true;
  }
  if (/^(false|0|no)$/i.test(raw)) {
    return false;
  }
  throw new Error(`Invalid boolean for ${name}: ${raw}`);
}

function rangeHasSheet(range) {
  return typeof range === "string" && range.includes("!");
}

function rangesHaveSheet(ranges) {
  return Array.isArray(ranges) && ranges.every((range) => rangeHasSheet(range));
}

async function main() {
  const missing = [];

  const sourceSpreadsheetId = process.env.TEST_SOURCE_SPREADSHEET_ID;
  if (!sourceSpreadsheetId) {
    missing.push("TEST_SOURCE_SPREADSHEET_ID");
  }

  const sourceRange = process.env.TEST_SOURCE_RANGE;
  const sourceRanges = parseJsonEnv("TEST_SOURCE_RANGES");
  if (!sourceRange && !Array.isArray(sourceRanges)) {
    missing.push("TEST_SOURCE_RANGE or TEST_SOURCE_RANGES");
  }

  const sourceSheetName = process.env.TEST_SOURCE_SHEET_NAME;
  const sourceGid = process.env.TEST_SOURCE_GID;
  const needsSourceSheet =
    (sourceRange && !rangeHasSheet(sourceRange)) ||
    (Array.isArray(sourceRanges) && sourceRanges.length > 0 && !rangesHaveSheet(sourceRanges));

  if (needsSourceSheet && !sourceSheetName && !sourceGid) {
    missing.push("TEST_SOURCE_SHEET_NAME or TEST_SOURCE_GID");
  }

  const destinationSpreadsheetId = process.env.TEST_DEST_SPREADSHEET_ID;
  if (!destinationSpreadsheetId) {
    missing.push("TEST_DEST_SPREADSHEET_ID");
  }

  const destinationSheetName = process.env.TEST_DEST_SHEET_NAME;
  const destinationGid = process.env.TEST_DEST_GID;
  if (!destinationSheetName && !destinationGid) {
    missing.push("TEST_DEST_SHEET_NAME or TEST_DEST_GID");
  }

  if (missing.length > 0) {
    console.error("Missing required test environment variables:");
    missing.forEach((name) => console.error(`- ${name}`));
    console.error("\nSet these in .env or your shell before running this test.");
    process.exit(1);
  }

  const removeColumns = parseJsonEnv("TEST_REMOVE_COLUMNS");
  const keepColumns = parseJsonEnv("TEST_KEEP_COLUMNS");
  const headerRowIndex = Number.parseInt(
    process.env.TEST_HEADER_ROW_INDEX || "0",
    10
  );
  const clearDestination = parseBoolEnv("TEST_CLEAR_DESTINATION", true);
  const previewRows = Number.parseInt(
    process.env.TEST_PREVIEW_ROWS || "5",
    10
  );

  const sheets = await buildSheetsClient();

  const result = await runImport({
    sheets,
    source: {
      spreadsheetId: sourceSpreadsheetId,
      range: sourceRange,
      ranges: sourceRanges,
      sheetName: sourceSheetName,
      gid: sourceGid,
    },
    destination: {
      spreadsheetId: destinationSpreadsheetId,
      sheetName: destinationSheetName,
      gid: destinationGid,
      startCell: process.env.TEST_DEST_START_CELL || "A1",
    },
    removeColumns,
    keepColumns,
    headerRowIndex: Number.isNaN(headerRowIndex) ? 0 : headerRowIndex,
    clearDestination,
  });

  const resolvedDestSheet = await resolveSheetName(
    sheets,
    destinationSpreadsheetId,
    destinationSheetName,
    destinationGid
  );
  const readRange = process.env.TEST_DEST_READ_RANGE || resolvedDestSheet;
  const values = await readValues(sheets, destinationSpreadsheetId, readRange);

  const rowCount = values.length;
  const colCount = values.reduce(
    (max, row) => Math.max(max, Array.isArray(row) ? row.length : 0),
    0
  );

  console.log("Import completed.");
  console.log(
    JSON.stringify(
      {
        updatedRange: result.updatedRange,
        updatedRows: result.updatedRows,
        updatedColumns: result.updatedColumns,
        updatedCells: result.updatedCells,
      },
      null,
      2
    )
  );
  console.log(`Destination read range: ${readRange}`);
  console.log(`Rows: ${rowCount}, Columns: ${colCount}`);

  if (previewRows > 0) {
    console.log("Preview:");
    console.log(JSON.stringify(values.slice(0, previewRows), null, 2));
  }
}

main().catch((error) => {
  const message = error && error.message ? error.message : String(error);
  console.error(`Test failed: ${message}`);
  process.exit(1);
});
