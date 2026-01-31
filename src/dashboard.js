const { ensureSheetExists, clearSheet } = require("./sheets");

const STUCK_STATUSES = [
  "Moving Parcel",
  "SOC_Packed",
  "Delivered Parcel",
  "Lost",
  "SOC_Packing",
  "Disposed",
  "SOC_Received",
];

const BUCKET_LABELS = ["l.15-20d+", "h.2d"];
const MONTHS = [
  "Jan",
  "Feb",
  "Mar",
  "Apr",
  "May",
  "Jun",
  "Jul",
  "Aug",
  "Sep",
  "Oct",
  "Nov",
  "Dec",
];
const DAY_MS = 24 * 60 * 60 * 1000;

function serialToDate(serial) {
  const epoch = Date.UTC(1899, 11, 30);
  return new Date(epoch + serial * DAY_MS);
}

function normalizeDate(value) {
  if (value === null || value === undefined || value === "") {
    return null;
  }

  let dateObj = null;
  if (typeof value === "number" && Number.isFinite(value)) {
    dateObj = serialToDate(value);
  } else {
    const parsed = new Date(value);
    if (!Number.isNaN(parsed.getTime())) {
      dateObj = parsed;
    }
  }

  if (dateObj) {
    const key = dateObj.toISOString().slice(0, 10);
    const label = `${MONTHS[dateObj.getMonth()]}-${dateObj.getDate()}`;
    return { key, label, dateObj };
  }

  const label = String(value).trim();
  if (!label) {
    return null;
  }
  return { key: label, label, dateObj: null };
}

function getText(value) {
  if (value === null || value === undefined) {
    return "";
  }
  return String(value).trim();
}

function incrementNested(map, key, subKey, amount = 1) {
  if (!map.has(key)) {
    map.set(key, new Map());
  }
  const nested = map.get(key);
  nested.set(subKey, (nested.get(subKey) || 0) + amount);
}

function increment(map, key, amount = 1) {
  map.set(key, (map.get(key) || 0) + amount);
}

function sortDateInfosDesc(dateInfos) {
  return [...dateInfos].sort((a, b) => {
    if (a.dateObj && b.dateObj) {
      return b.dateObj.getTime() - a.dateObj.getTime();
    }
    if (a.dateObj) {
      return -1;
    }
    if (b.dateObj) {
      return 1;
    }
    return String(b.key).localeCompare(String(a.key));
  });
}

function pickYesterdayKey(dateInfos, availableKeys) {
  const dated = dateInfos.filter((info) => info.dateObj);
  if (dated.length > 0) {
    const maxDate = new Date(
      Math.max(...dated.map((info) => info.dateObj.getTime()))
    );
    const yesterday = new Date(maxDate.getTime() - DAY_MS);
    const key = yesterday.toISOString().slice(0, 10);
    if (availableKeys.has(key)) {
      return key;
    }
  }

  const sorted = sortDateInfosDesc(dateInfos);
  if (sorted.length > 1) {
    return sorted[1].key;
  }
  return sorted.length > 0 ? sorted[0].key : null;
}

function buildCounts(rows, headerRowIndex) {
  const dateMap = new Map();
  const regionDateCounts = new Map();
  const statusDateCounts = new Map();
  const bucketTotals = new Map();
  const statusTotals = new Map();
  const hubDateCounts = new Map();

  const startIndex = Math.min(Math.max(headerRowIndex + 1, 0), rows.length);
  rows.slice(startIndex).forEach((row) => {
    const dateInfo = normalizeDate(row[0]);
    if (!dateInfo) {
      return;
    }

    dateMap.set(dateInfo.key, dateInfo);

    const hub = getText(row[6]);
    const bucket = getText(row[11]);
    const region = getText(row[13]);
    const status = getText(row[14]);

    if (region) {
      incrementNested(regionDateCounts, region, dateInfo.key);
    }
    if (status) {
      incrementNested(statusDateCounts, status, dateInfo.key);
      increment(statusTotals, status);
    }
    if (bucket) {
      increment(bucketTotals, bucket);
    }
    if (hub) {
      incrementNested(hubDateCounts, dateInfo.key, hub);
    }
  });

  return {
    dateInfos: Array.from(dateMap.values()),
    regionDateCounts,
    statusDateCounts,
    bucketTotals,
    statusTotals,
    hubDateCounts,
  };
}

function buildSummaryText({
  topRegion,
  topRegionAve,
  stuckAverage,
  topBucket,
}) {
  if (!topRegion) {
    return "No data available for the dashboard summary.";
  }
  const bucketLabel = topBucket || "N/A";
  return `20hrs - 1d Validation Summary: ${topRegion} shows highest stuckup orders (${topRegionAve} Ave L7D). 7-Day Average Stuck Up Tagging is ${stuckAverage} orders. ${bucketLabel} Ageing Bucket is top contributor.`;
}

function buildTableHeader(prefix, dateLabels) {
  return [...prefix, ...dateLabels];
}

async function updateDashboard({
  sheets,
  spreadsheetId,
  dataSheetName,
  dashboardSheetName = "Dashboard",
  headerRowIndex = 0,
}) {
  if (!dataSheetName) {
    throw new Error("Missing dataSheetName for dashboard generation.");
  }

  await ensureSheetExists(sheets, spreadsheetId, dashboardSheetName);
  await clearSheet(sheets, spreadsheetId, dashboardSheetName);

  const meta = await sheets.spreadsheets.get({
    spreadsheetId,
    includeGridData: false,
  });
  const dashboardSheetId = getSheetIdByTitle(
    meta.data,
    dashboardSheetName
  );

  const response = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: `${dataSheetName}!A:O`,
    valueRenderOption: "UNFORMATTED_VALUE",
  });

  const rows = response.data.values || [];
  const {
    dateInfos,
    regionDateCounts,
    statusDateCounts,
    bucketTotals,
    statusTotals,
    hubDateCounts,
  } = buildCounts(rows, headerRowIndex);

  const sortedDates = sortDateInfosDesc(dateInfos);
  const last7Dates = sortedDates.slice(0, 7);
  const last7Keys = last7Dates.map((info) => info.key);
  const dateLabels = last7Dates.map((info) => info.label);

  const regionTotals = new Map();
  const regionRows = Array.from(regionDateCounts.keys()).map((region) => {
    const counts = last7Keys.map(
      (key) => regionDateCounts.get(region).get(key) || 0
    );
    const total = counts.reduce((sum, value) => sum + value, 0);
    const average = counts.length > 0 ? Math.round(total / counts.length) : 0;
    regionTotals.set(region, total);
    return [region, average, total, ...counts];
  });

  regionRows.sort((a, b) => b[2] - a[2] || String(a[0]).localeCompare(String(b[0])));

  const totalCountsByDate = last7Keys.map((key) =>
    regionRows.reduce((sum, row) => sum + (row[3 + last7Keys.indexOf(key)] || 0), 0)
  );
  const totalL7D = totalCountsByDate.reduce((sum, value) => sum + value, 0);
  const totalAve = last7Keys.length > 0 ? Math.round(totalL7D / last7Keys.length) : 0;

  const regionalTable = [
    buildTableHeader(["Region", "Ave L7D", "Total L7D"], dateLabels),
    ...regionRows,
    ["Total", totalAve, totalL7D, ...totalCountsByDate],
  ];

  const topRegionRow = regionRows[0] || null;

  const statusRows = STUCK_STATUSES.map((status) => {
    const counts = last7Keys.map((key) => {
      const nested = statusDateCounts.get(status);
      return nested ? nested.get(key) || 0 : 0;
    });
    const total = counts.reduce((sum, value) => sum + value, 0);
    const average = counts.length > 0 ? Math.round(total / counts.length) : 0;
    return [status, average, total, ...counts];
  });

  const statusTotalsByDate = last7Keys.map((key) =>
    statusRows.reduce((sum, row) => sum + (row[3 + last7Keys.indexOf(key)] || 0), 0)
  );
  const statusTotalL7D = statusTotalsByDate.reduce(
    (sum, value) => sum + value,
    0
  );
  const statusAveL7D = last7Keys.length > 0 ? Math.round(statusTotalL7D / last7Keys.length) : 0;

  const stuckTable = [
    buildTableHeader(["Status", "Ave L7D", "Total L7D"], dateLabels),
    ...statusRows,
    ["Total", statusAveL7D, statusTotalL7D, ...statusTotalsByDate],
  ];

  const bucketTotal = Array.from(bucketTotals.values()).reduce(
    (sum, value) => sum + value,
    0
  );
  const bucketRows = BUCKET_LABELS.map((label) => {
    const count = bucketTotals.get(label) || 0;
    const percent = bucketTotal > 0 ? count / bucketTotal : 0;
    return [label, count, percent];
  });

  const bucketTable = [
    ["Ageing Bucket", "Volume", "Percentage"],
    ...bucketRows,
  ];

  const statusTotalAll = Array.from(statusTotals.values()).reduce(
    (sum, value) => sum + value,
    0
  );
  const statusVolumeRows = Array.from(statusTotals.entries())
    .sort((a, b) => b[1] - a[1])
    .slice(0, 7)
    .map(([status, count]) => [
      status,
      count,
      statusTotalAll > 0 ? count / statusTotalAll : 0,
    ]);
  const statusVolumeTable = [
    ["Status", "Volume", "Percentage"],
    ["Total", statusTotalAll, statusTotalAll > 0 ? 1 : 0],
    ...statusVolumeRows,
  ];

  const yesterdayKey = pickYesterdayKey(dateInfos, new Set(hubDateCounts.keys()));
  const hubCounts = yesterdayKey ? hubDateCounts.get(yesterdayKey) : null;
  const hubTotal = hubCounts
    ? Array.from(hubCounts.values()).reduce((sum, value) => sum + value, 0)
    : 0;
  const topHubs = hubCounts
    ? Array.from(hubCounts.entries())
        .sort((a, b) => b[1] - a[1])
        .slice(0, 5)
    : [];
  const topHubRows = topHubs.map(([hub, count]) => [
    hub,
    count,
    hubTotal > 0 ? count / hubTotal : 0,
  ]);
  const topHubTable = [
    ["Hub", "Volume", "Percentage"],
    ...topHubRows,
  ];

  const topBucket = bucketRows.reduce(
    (best, row) => (row[1] > best.count ? { label: row[0], count: row[1] } : best),
    { label: null, count: -1 }
  );
  const summaryText = buildSummaryText({
    topRegion: topRegionRow ? topRegionRow[0] : null,
    topRegionAve: topRegionRow ? topRegionRow[1] : 0,
    stuckAverage: statusAveL7D,
    topBucket: topBucket.label,
  });

  const topRegionsForChart = regionRows.slice(0, 6).map((row) => row[0]);
  const regionalChartData = [
    ["Date", ...topRegionsForChart],
    ...last7Keys.map((key, index) => {
      const label = dateLabels[index] || key;
      const values = topRegionsForChart.map((region) => {
        const nested = regionDateCounts.get(region);
        return nested ? nested.get(key) || 0 : 0;
      });
      return [label, ...values];
    }),
  ];

  const statusChartData = [
    ["Date", ...STUCK_STATUSES],
    ...last7Keys.map((key, index) => {
      const label = dateLabels[index] || key;
      const values = STUCK_STATUSES.map((status) => {
        const nested = statusDateCounts.get(status);
        return nested ? nested.get(key) || 0 : 0;
      });
      return [label, ...values];
    }),
  ];

  const regionalTableStartRow = 5;
  const regionalTableStartCol = 1;
  const bucketTableStartRow = 5;
  const bucketTableStartCol = 12;
  const statusVolumeStartRow = bucketTableStartRow + bucketTable.length + 2;
  const statusVolumeStartCol = bucketTableStartCol;
  const topHubTableStartRow = 5;
  const topHubTableStartCol = 19;
  const stuckTableTitleRow = regionalTableStartRow + regionalTable.length + 2;
  const stuckTableStartRow = stuckTableTitleRow + 1;
  const stuckTableStartCol = 1;
  const regionalChartStartRow = 1;
  const regionalChartStartCol = 27;
  const statusChartStartRow = regionalChartStartRow + regionalChartData.length + 3;
  const statusChartStartCol = regionalChartStartCol;

  const valueRequests = [
    {
      range: `${dashboardSheetName}!A1`,
      values: [["Daily Briefing"]],
    },
    {
      range: `${dashboardSheetName}!A2`,
      values: [[summaryText]],
    },
    {
      range: `${dashboardSheetName}!A4`,
      values: [["Regional Validation Summary"]],
    },
    {
      range: `${dashboardSheetName}!A${regionalTableStartRow}`,
      values: regionalTable,
    },
    {
      range: `${dashboardSheetName}!L4`,
      values: [["Ageing Bucket Analysis"]],
    },
    {
      range: `${dashboardSheetName}!L${bucketTableStartRow}`,
      values: bucketTable,
    },
    {
      range: `${dashboardSheetName}!L${statusVolumeStartRow - 1}`,
      values: [["Status Volume"]],
    },
    {
      range: `${dashboardSheetName}!L${statusVolumeStartRow}`,
      values: statusVolumeTable,
    },
    {
      range: `${dashboardSheetName}!S4`,
      values: [["Top Hubs"]],
    },
    {
      range: `${dashboardSheetName}!S${topHubTableStartRow}`,
      values: topHubTable,
    },
    {
      range: `${dashboardSheetName}!A${stuckTableTitleRow}`,
      values: [["Stuck Up Tagging Analysis"]],
    },
    {
      range: `${dashboardSheetName}!A${stuckTableStartRow}`,
      values: stuckTable,
    },
    {
      range: `${dashboardSheetName}!AA${regionalChartStartRow}`,
      values: regionalChartData,
    },
    {
      range: `${dashboardSheetName}!AA${statusChartStartRow}`,
      values: statusChartData,
    },
  ];

  await sheets.spreadsheets.values.batchUpdate({
    spreadsheetId,
    valueInputOption: "USER_ENTERED",
    requestBody: {
      data: valueRequests,
    },
  });

  const chartDeletes = [];
  (meta.data.sheets || []).forEach((sheet) => {
    (sheet.charts || []).forEach((chart) => {
      const anchorSheetId =
        chart.position &&
        chart.position.overlayPosition &&
        chart.position.overlayPosition.anchorCell &&
        chart.position.overlayPosition.anchorCell.sheetId;
      if (anchorSheetId === dashboardSheetId) {
        chartDeletes.push({
          deleteEmbeddedObject: {
            objectId: chart.chartId,
          },
        });
      }
    });
  });

  const headerFormat = {
    textFormat: { bold: true },
    backgroundColor: { red: 0.92, green: 0.94, blue: 0.96 },
  };

  const requests = [
    ...chartDeletes,
    {
      repeatCell: {
        range: {
          sheetId: dashboardSheetId,
          startRowIndex: 0,
          endRowIndex: 1,
          startColumnIndex: 0,
          endColumnIndex: 1,
        },
        cell: {
          userEnteredFormat: {
            textFormat: { bold: true, fontSize: 14 },
          },
        },
        fields: "userEnteredFormat.textFormat",
      },
    },
    {
      repeatCell: {
        range: {
          sheetId: dashboardSheetId,
          startRowIndex: regionalTableStartRow - 1,
          endRowIndex: regionalTableStartRow,
          startColumnIndex: 0,
          endColumnIndex: regionalTable[0].length,
        },
        cell: { userEnteredFormat: headerFormat },
        fields: "userEnteredFormat",
      },
    },
    {
      repeatCell: {
        range: {
          sheetId: dashboardSheetId,
          startRowIndex: bucketTableStartRow - 1,
          endRowIndex: bucketTableStartRow,
          startColumnIndex: bucketTableStartCol - 1,
          endColumnIndex: bucketTableStartCol - 1 + bucketTable[0].length,
        },
        cell: { userEnteredFormat: headerFormat },
        fields: "userEnteredFormat",
      },
    },
    {
      repeatCell: {
        range: {
          sheetId: dashboardSheetId,
          startRowIndex: statusVolumeStartRow - 1,
          endRowIndex: statusVolumeStartRow,
          startColumnIndex: statusVolumeStartCol - 1,
          endColumnIndex:
            statusVolumeStartCol - 1 + statusVolumeTable[0].length,
        },
        cell: { userEnteredFormat: headerFormat },
        fields: "userEnteredFormat",
      },
    },
    {
      repeatCell: {
        range: {
          sheetId: dashboardSheetId,
          startRowIndex: topHubTableStartRow - 1,
          endRowIndex: topHubTableStartRow,
          startColumnIndex: topHubTableStartCol - 1,
          endColumnIndex: topHubTableStartCol - 1 + topHubTable[0].length,
        },
        cell: { userEnteredFormat: headerFormat },
        fields: "userEnteredFormat",
      },
    },
    {
      repeatCell: {
        range: {
          sheetId: dashboardSheetId,
          startRowIndex: stuckTableStartRow - 1,
          endRowIndex: stuckTableStartRow,
          startColumnIndex: stuckTableStartCol - 1,
          endColumnIndex: stuckTableStartCol - 1 + stuckTable[0].length,
        },
        cell: { userEnteredFormat: headerFormat },
        fields: "userEnteredFormat",
      },
    },
  ];

  const percentFormat = {
    numberFormat: {
      type: "PERCENT",
      pattern: "0.0%",
    },
  };

  const percentRanges = [
    {
      startRow: bucketTableStartRow + 1,
      startCol: bucketTableStartCol + 2,
      rowCount: bucketTable.length - 1,
      colCount: 1,
    },
    {
      startRow: statusVolumeStartRow + 1,
      startCol: statusVolumeStartCol + 2,
      rowCount: statusVolumeTable.length - 1,
      colCount: 1,
    },
    {
      startRow: topHubTableStartRow + 1,
      startCol: topHubTableStartCol + 2,
      rowCount: topHubTable.length - 1,
      colCount: 1,
    },
  ];

  percentRanges.forEach((range) => {
    requests.push({
      repeatCell: {
        range: {
          sheetId: dashboardSheetId,
          startRowIndex: range.startRow - 1,
          endRowIndex: range.startRow - 1 + range.rowCount,
          startColumnIndex: range.startCol - 1,
          endColumnIndex: range.startCol - 1 + range.colCount,
        },
        cell: { userEnteredFormat: percentFormat },
        fields: "userEnteredFormat.numberFormat",
      },
    });
  });

  if (regionalChartData.length > 1) {
    const regionDomainRange = {
      sheetId: dashboardSheetId,
      startRowIndex: regionalChartStartRow - 1,
      endRowIndex: regionalChartStartRow - 1 + regionalChartData.length,
      startColumnIndex: regionalChartStartCol - 1,
      endColumnIndex: regionalChartStartCol,
    };
    const regionSeries = topRegionsForChart.map((_region, index) => ({
      series: {
        sourceRange: {
          sources: [
            {
              sheetId: dashboardSheetId,
              startRowIndex: regionalChartStartRow - 1,
              endRowIndex: regionalChartStartRow - 1 + regionalChartData.length,
              startColumnIndex: regionalChartStartCol + index,
              endColumnIndex: regionalChartStartCol + index + 1,
            },
          ],
        },
      },
      targetAxis: "LEFT_AXIS",
    }));

    requests.push({
      addChart: {
        chart: {
          spec: {
            title: "20hrs - 1d Validation Trend",
            basicChart: {
              chartType: "AREA",
              legendPosition: "TOP_LEGEND",
              headerCount: 1,
              domains: [
                {
                  domain: {
                    sourceRange: {
                      sources: [regionDomainRange],
                    },
                  },
                },
              ],
              series: regionSeries,
              axis: [
                { position: "BOTTOM_AXIS" },
                { position: "LEFT_AXIS" },
              ],
            },
          },
          position: {
            overlayPosition: {
              anchorCell: {
                sheetId: dashboardSheetId,
                rowIndex: stuckTableTitleRow - 2,
                columnIndex: 7,
              },
              widthPixels: 380,
              heightPixels: 220,
            },
          },
        },
      },
    });
  }

  if (statusChartData.length > 1) {
    const statusDomainRange = {
      sheetId: dashboardSheetId,
      startRowIndex: statusChartStartRow - 1,
      endRowIndex: statusChartStartRow - 1 + statusChartData.length,
      startColumnIndex: statusChartStartCol - 1,
      endColumnIndex: statusChartStartCol,
    };
    const statusSeries = STUCK_STATUSES.map((_status, index) => ({
      series: {
        sourceRange: {
          sources: [
            {
              sheetId: dashboardSheetId,
              startRowIndex: statusChartStartRow - 1,
              endRowIndex: statusChartStartRow - 1 + statusChartData.length,
              startColumnIndex: statusChartStartCol + index,
              endColumnIndex: statusChartStartCol + index + 1,
            },
          ],
        },
      },
      targetAxis: "LEFT_AXIS",
    }));

    requests.push({
      addChart: {
        chart: {
          spec: {
            title: "Stuck Up Tagging Trend",
            basicChart: {
              chartType: "AREA",
              legendPosition: "TOP_LEGEND",
              headerCount: 1,
              domains: [
                {
                  domain: {
                    sourceRange: {
                      sources: [statusDomainRange],
                    },
                  },
                },
              ],
              series: statusSeries,
              axis: [
                { position: "BOTTOM_AXIS" },
                { position: "LEFT_AXIS" },
              ],
            },
          },
          position: {
            overlayPosition: {
              anchorCell: {
                sheetId: dashboardSheetId,
                rowIndex: stuckTableTitleRow - 2,
                columnIndex: 14,
              },
              widthPixels: 380,
              heightPixels: 220,
            },
          },
        },
      },
    });
  }

  if (bucketTable.length > 1) {
    const bucketDomainRange = {
      sheetId: dashboardSheetId,
      startRowIndex: bucketTableStartRow,
      endRowIndex: bucketTableStartRow + bucketTable.length - 1,
      startColumnIndex: bucketTableStartCol - 1,
      endColumnIndex: bucketTableStartCol,
    };
    const bucketSeriesRange = {
      sheetId: dashboardSheetId,
      startRowIndex: bucketTableStartRow,
      endRowIndex: bucketTableStartRow + bucketTable.length - 1,
      startColumnIndex: bucketTableStartCol,
      endColumnIndex: bucketTableStartCol + 1,
    };

    requests.push({
      addChart: {
        chart: {
          spec: {
            title: "Ageing Bucket Analysis",
            basicChart: {
              chartType: "BAR",
              legendPosition: "NO_LEGEND",
              headerCount: 0,
              domains: [
                {
                  domain: {
                    sourceRange: {
                      sources: [bucketDomainRange],
                    },
                  },
                },
              ],
              series: [
                {
                  series: {
                    sourceRange: {
                      sources: [bucketSeriesRange],
                    },
                  },
                  targetAxis: "BOTTOM_AXIS",
                },
              ],
              axis: [
                { position: "BOTTOM_AXIS" },
                { position: "LEFT_AXIS" },
              ],
            },
          },
          position: {
            overlayPosition: {
              anchorCell: {
                sheetId: dashboardSheetId,
                rowIndex: bucketTableStartRow - 1,
                columnIndex: bucketTableStartCol + 3,
              },
              widthPixels: 320,
              heightPixels: 180,
            },
          },
        },
      },
    });
  }

  if (topHubTable.length > 1) {
    const hubDomainRange = {
      sheetId: dashboardSheetId,
      startRowIndex: topHubTableStartRow,
      endRowIndex: topHubTableStartRow + topHubTable.length - 1,
      startColumnIndex: topHubTableStartCol - 1,
      endColumnIndex: topHubTableStartCol,
    };
    const hubSeriesRange = {
      sheetId: dashboardSheetId,
      startRowIndex: topHubTableStartRow,
      endRowIndex: topHubTableStartRow + topHubTable.length - 1,
      startColumnIndex: topHubTableStartCol,
      endColumnIndex: topHubTableStartCol + 1,
    };

    requests.push({
      addChart: {
        chart: {
          spec: {
            title: "Top Hubs",
            pieChart: {
              legendPosition: "RIGHT_LEGEND",
              domain: {
                sourceRange: {
                  sources: [hubDomainRange],
                },
              },
              series: {
                sourceRange: {
                  sources: [hubSeriesRange],
                },
              },
              pieHole: 0.5,
            },
          },
          position: {
            overlayPosition: {
              anchorCell: {
                sheetId: dashboardSheetId,
                rowIndex: topHubTableStartRow - 1,
                columnIndex: topHubTableStartCol + 3,
              },
              widthPixels: 240,
              heightPixels: 200,
            },
          },
        },
      },
    });
  }

  if (requests.length > 0) {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: { requests },
    });
  }
}

function getSheetIdByTitle(spreadsheet, sheetName) {
  const match = (spreadsheet.sheets || []).find(
    (sheet) => sheet.properties && sheet.properties.title === sheetName
  );
  if (!match || !match.properties || match.properties.sheetId === undefined) {
    throw new Error(`Sheet ${sheetName} not found.`);
  }
  return match.properties.sheetId;
}

module.exports = {
  updateDashboard,
};
