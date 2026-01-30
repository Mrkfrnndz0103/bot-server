const fs = require("fs");
const { google } = require("googleapis");

function loadServiceAccountJson() {
  const inlineJson = process.env.GOOGLE_SERVICE_ACCOUNT_JSON;
  if (inlineJson && inlineJson.trim()) {
    const normalized = inlineJson.replace(/\\n/g, "\n");
    return JSON.parse(normalized);
  }

  const jsonPath = process.env.GOOGLE_APPLICATION_CREDENTIALS;
  if (!jsonPath) {
    throw new Error(
      "Missing GOOGLE_SERVICE_ACCOUNT_JSON or GOOGLE_APPLICATION_CREDENTIALS."
    );
  }

  const fileContents = fs.readFileSync(jsonPath, "utf8");
  return JSON.parse(fileContents);
}

async function buildSheetsClient() {
  const credentials = loadServiceAccountJson();
  const auth = new google.auth.GoogleAuth({
    credentials: {
      client_email: credentials.client_email,
      private_key: credentials.private_key,
    },
    scopes: ["https://www.googleapis.com/auth/spreadsheets"],
  });
  return google.sheets({ version: "v4", auth });
}

async function ensureSheetExists(sheets, spreadsheetId, sheetName) {
  const meta = await sheets.spreadsheets.get({
    spreadsheetId,
  });

  const found = (meta.data.sheets || []).some(
    (sheet) => sheet.properties && sheet.properties.title === sheetName
  );

  if (found) {
    return;
  }

  await sheets.spreadsheets.batchUpdate({
    spreadsheetId,
    requestBody: {
      requests: [
        {
          addSheet: {
            properties: {
              title: sheetName,
            },
          },
        },
      ],
    },
  });
}

async function getSheetTitleById(sheets, spreadsheetId, sheetId) {
  const meta = await sheets.spreadsheets.get({
    spreadsheetId,
  });
  const match = (meta.data.sheets || []).find(
    (sheet) =>
      sheet.properties && Number(sheet.properties.sheetId) === Number(sheetId)
  );
  if (!match || !match.properties || !match.properties.title) {
    throw new Error(
      `Sheet ID ${sheetId} not found in spreadsheet ${spreadsheetId}.`
    );
  }
  return match.properties.title;
}

async function readValues(sheets, spreadsheetId, range) {
  const response = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range,
  });
  return response.data.values || [];
}

async function batchReadValues(sheets, spreadsheetId, ranges) {
  if (!Array.isArray(ranges) || ranges.length === 0) {
    return [];
  }

  const response = await sheets.spreadsheets.values.batchGet({
    spreadsheetId,
    ranges,
  });

  const valueRanges = response.data.valueRanges || [];
  return valueRanges.map((valueRange) => valueRange.values || []);
}

async function clearSheet(sheets, spreadsheetId, sheetName) {
  await sheets.spreadsheets.values.clear({
    spreadsheetId,
    range: sheetName,
  });
}

async function clearRange(sheets, spreadsheetId, range) {
  await sheets.spreadsheets.values.clear({
    spreadsheetId,
    range,
  });
}

async function writeValues(
  sheets,
  spreadsheetId,
  sheetName,
  startCell,
  values
) {
  if (!values || values.length === 0) {
    return {
      updatedRange: `${sheetName}!${startCell}`,
      updatedRows: 0,
      updatedColumns: 0,
      updatedCells: 0,
    };
  }

  const response = await sheets.spreadsheets.values.update({
    spreadsheetId,
    range: `${sheetName}!${startCell}`,
    valueInputOption: "RAW",
    requestBody: {
      values,
    },
  });

  return response.data || {};
}

function removeColumnsFromRows(rows, columnsToRemove) {
  if (!Array.isArray(rows) || rows.length === 0) {
    return [];
  }

  if (!Array.isArray(columnsToRemove) || columnsToRemove.length === 0) {
    return rows;
  }

  const removeSet = new Set(
    columnsToRemove
      .map((index) => Number(index))
      .filter((index) => Number.isInteger(index) && index >= 0)
  );

  if (removeSet.size === 0) {
    return rows;
  }

  return rows.map((row) => row.filter((_value, index) => !removeSet.has(index)));
}

function keepColumnsFromRows(
  rows,
  columnsToKeep,
  { headerRowIndex = 0, caseInsensitive = true, trim = true } = {}
) {
  if (!Array.isArray(rows) || rows.length === 0) {
    return [];
  }

  if (!Array.isArray(columnsToKeep) || columnsToKeep.length === 0) {
    return rows;
  }

  const headerRow = rows[headerRowIndex];
  if (!Array.isArray(headerRow)) {
    throw new Error("Header row not found for keepColumns.");
  }

  const normalize = (value) => {
    if (typeof value !== "string") {
      return "";
    }
    let next = value;
    if (trim) {
      next = next.trim();
    }
    if (caseInsensitive) {
      next = next.toLowerCase();
    }
    return next;
  };

  const headerIndexMap = new Map();
  headerRow.forEach((headerValue, index) => {
    const key = normalize(String(headerValue));
    if (!headerIndexMap.has(key)) {
      headerIndexMap.set(key, index);
    }
  });

  const requested = columnsToKeep.map((name) => normalize(String(name)));
  const keepIndexes = requested
    .map((name) => headerIndexMap.get(name))
    .filter((index) => Number.isInteger(index));

  const missing = requested.filter((name) => !headerIndexMap.has(name));
  if (missing.length > 0) {
    throw new Error(
      `Missing keepColumns in header row: ${missing.join(", ")}`
    );
  }

  return rows.map((row) => keepIndexes.map((index) => row[index]));
}

module.exports = {
  buildSheetsClient,
  ensureSheetExists,
  getSheetTitleById,
  readValues,
  batchReadValues,
  clearSheet,
  clearRange,
  writeValues,
  removeColumnsFromRows,
  keepColumnsFromRows,
};
