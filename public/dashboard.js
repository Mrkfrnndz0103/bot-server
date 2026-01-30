const regionalTable = document.querySelector("#regional-table");
const regionalHead = regionalTable.querySelector("thead");
const regionalBody = regionalTable.querySelector("tbody");
const stuckupTable = document.querySelector("#stuckup-table");
const stuckupHead = stuckupTable.querySelector("thead");
const stuckupBody = stuckupTable.querySelector("tbody");
const ageingTable = document.querySelector("#ageing-table");
const ageingHead = ageingTable.querySelector("thead");
const ageingBody = ageingTable.querySelector("tbody");
const ageingChartEl = document.querySelector("#ageing-chart");
const topHubsChartEl = document.querySelector("#top-hubs-chart");
const validationTrendChartEl = document.querySelector("#validation-trend-chart");
const stuckupTrendChartEl = document.querySelector("#stuckup-trend-chart");
const statusEl = document.querySelector("#last-updated");

const charts = new Map();

function setChart(key, chart) {
  if (charts.has(key)) {
    charts.get(key).destroy();
  }
  charts.set(key, chart);
}

function parseNumber(value) {
  if (value === null || value === undefined || value === "") {
    return null;
  }
  const cleaned = String(value).replace(/,/g, "");
  const parsed = Number.parseFloat(cleaned);
  return Number.isFinite(parsed) ? parsed : null;
}

function buildRegionalTable(headers, rows) {
  regionalHead.innerHTML = "";
  regionalBody.innerHTML = "";

  const headerRow = document.createElement("tr");
  headers.forEach((label) => {
    const th = document.createElement("th");
    th.textContent = label;
    headerRow.appendChild(th);
  });
  regionalHead.appendChild(headerRow);

  const numericValues = [];
  rows.forEach((row) => {
    row.slice(1).forEach((cell) => {
      const num = parseNumber(cell);
      if (num !== null) {
        numericValues.push(num);
      }
    });
  });

  const min = Math.min(...numericValues);
  const max = Math.max(...numericValues);
  const span = max - min || 1;

  rows.forEach((row) => {
    const tr = document.createElement("tr");
    row.forEach((cell, index) => {
      if (index === 0) {
        const th = document.createElement("th");
        th.textContent = cell;
        tr.appendChild(th);
        return;
      }

      const td = document.createElement("td");
      const num = parseNumber(cell);
      td.textContent = cell;
      td.classList.add("heat-cell");

      if (num !== null) {
        const intensity = (num - min) / span;
        const alpha = 0.15 + intensity * 0.65;
        td.style.background = `rgba(55, 160, 199, ${alpha})`;
      }
      tr.appendChild(td);
    });
    regionalBody.appendChild(tr);
  });
}

function buildSimpleTable(headers, rows, headEl, bodyEl, { heatmap } = {}) {
  headEl.innerHTML = "";
  bodyEl.innerHTML = "";

  const headerRow = document.createElement("tr");
  headers.forEach((label) => {
    const th = document.createElement("th");
    th.textContent = label;
    headerRow.appendChild(th);
  });
  headEl.appendChild(headerRow);

  let min = 0;
  let max = 0;
  let span = 1;
  if (heatmap) {
    const numericValues = [];
    rows.forEach((row) => {
      row.slice(1).forEach((cell) => {
        const num = parseNumber(cell);
        if (num !== null) {
          numericValues.push(num);
        }
      });
    });
    if (numericValues.length > 0) {
      min = Math.min(...numericValues);
      max = Math.max(...numericValues);
      span = max - min || 1;
    }
  }

  rows.forEach((row) => {
    const tr = document.createElement("tr");
    row.forEach((cell, index) => {
      const td = document.createElement(index === 0 ? "th" : "td");
      td.textContent = cell;
      if (heatmap && index > 0) {
        const num = parseNumber(cell);
        td.classList.add("heat-cell");
        if (num !== null) {
          const intensity = (num - min) / span;
          const alpha = 0.12 + intensity * 0.7;
          td.style.background = `rgba(55, 160, 199, ${alpha})`;
        }
      }
      tr.appendChild(td);
    });
    bodyEl.appendChild(tr);
  });
}

function buildHorizontalBarChart(canvas, labels, values) {
  const ctx = canvas.getContext("2d");
  const chart = new Chart(ctx, {
    type: "bar",
    data: {
      labels,
      datasets: [
        {
          label: "Volume",
          data: values,
          backgroundColor: "rgba(55, 160, 199, 0.7)",
          borderRadius: 6,
        },
      ],
    },
    options: {
      indexAxis: "y",
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: { mode: "nearest" },
      },
      scales: {
        x: {
          grid: { color: "rgba(148, 163, 184, 0.2)" },
          ticks: { color: "#6b7280" },
        },
        y: {
          grid: { display: false },
          ticks: { color: "#0f172a" },
        },
      },
    },
  });
  setChart(canvas.id, chart);
}

function buildDoughnutChart(canvas, labels, values) {
  const palette = [
    "#0f5a7a",
    "#37a0c7",
    "#7ec7de",
    "#9fd5ea",
    "#b9e2f2",
    "#d5eff8",
  ];
  const ctx = canvas.getContext("2d");
  const chart = new Chart(ctx, {
    type: "doughnut",
    data: {
      labels,
      datasets: [
        {
          data: values,
          backgroundColor: labels.map((_, i) => palette[i % palette.length]),
          borderWidth: 0,
        },
      ],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: {
          position: "bottom",
          labels: {
            boxWidth: 10,
            color: "#0f172a",
          },
        },
      },
      cutout: "62%",
    },
  });
  setChart(canvas.id, chart);
}

function buildLineChart(canvas, headers, rows) {
  const labels = headers.slice(1);
  const palette = [
    "rgba(15, 90, 122, 0.6)",
    "rgba(55, 160, 199, 0.5)",
    "rgba(126, 199, 222, 0.5)",
    "rgba(159, 213, 234, 0.5)",
    "rgba(185, 226, 242, 0.5)",
    "rgba(97, 156, 185, 0.5)",
  ];

  const datasets = rows.map((row, index) => {
    const color = palette[index % palette.length];
    return {
      label: row[0],
      data: row.slice(1).map(parseNumber),
      borderColor: color.replace("0.5", "0.9"),
      backgroundColor: color,
      fill: true,
      tension: 0.35,
      borderWidth: 2,
      pointRadius: 2,
      pointHoverRadius: 4,
    };
  });

  const ctx = canvas.getContext("2d");
  const chart = new Chart(ctx, {
    type: "line",
    data: { labels, datasets },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: {
          position: "top",
          labels: { color: "#0f172a", boxWidth: 10 },
        },
      },
      scales: {
        x: {
          grid: { color: "rgba(148, 163, 184, 0.15)" },
          ticks: { color: "#6b7280" },
        },
        y: {
          grid: { color: "rgba(148, 163, 184, 0.2)" },
          ticks: { color: "#6b7280" },
        },
      },
    },
  });
  setChart(canvas.id, chart);
}

async function loadRegionalSummary() {
  statusEl.textContent = "Loading…";
  try {
    const response = await fetch("/api/regional-validation");
    const data = await response.json();
    if (!response.ok) {
      throw new Error(data.error || "Failed to load data.");
    }
    buildRegionalTable(data.headers || [], data.rows || []);
    statusEl.textContent = `Last sync ${new Date().toLocaleString()}`;
  } catch (error) {
    const message =
      error && typeof error.message === "string"
        ? error.message
        : "Failed to load.";
    statusEl.textContent = message;
  }
}

async function loadStuckupAnalysis() {
  try {
    const response = await fetch("/api/stuckup-analysis");
    const data = await response.json();
    if (!response.ok) {
      throw new Error(data.error || "Failed to load data.");
    }
    buildSimpleTable(
      data.headers || [],
      data.rows || [],
      stuckupHead,
      stuckupBody,
      { heatmap: true }
    );
  } catch (_error) {
    stuckupHead.innerHTML = "";
    stuckupBody.innerHTML =
      "<tr><td colspan=\"5\">No data available</td></tr>";
  }
}

async function loadAgeingBucket() {
  try {
    const response = await fetch("/api/ageing-bucket");
    const data = await response.json();
    if (!response.ok) {
      throw new Error(data.error || "Failed to load data.");
    }
    buildSimpleTable(data.headers || [], data.rows || [], ageingHead, ageingBody, {
      heatmap: true,
    });
    const labels = (data.rows || []).map((row) => row[0]);
    const values = (data.rows || []).map((row) => parseNumber(row[1]) || 0);
    buildHorizontalBarChart(ageingChartEl, labels, values);
  } catch (_error) {
    ageingHead.innerHTML = "";
    ageingBody.innerHTML =
      "<tr><td colspan=\"5\">No data available</td></tr>";
  }
}

async function loadTopHubs() {
  try {
    const response = await fetch("/api/top-hubs");
    const data = await response.json();
    if (!response.ok) {
      throw new Error(data.error || "Failed to load data.");
    }
    const labels = (data.rows || []).map((row) => row[0]);
    const values = (data.rows || []).map((row) => parseNumber(row[1]) || 0);
    buildDoughnutChart(topHubsChartEl, labels, values);
  } catch (_error) {
    topHubsChartEl.replaceWith(document.createTextNode("No data available"));
  }
}

async function loadValidationTrend() {
  try {
    const response = await fetch("/api/validation-trend");
    const data = await response.json();
    if (!response.ok) {
      throw new Error(data.error || "Failed to load data.");
    }
    buildLineChart(validationTrendChartEl, data.headers || [], data.rows || []);
  } catch (_error) {
    validationTrendChartEl.replaceWith(
      document.createTextNode("No data available")
    );
  }
}

async function loadStuckupTrend() {
  try {
    const response = await fetch("/api/stuckup-trend");
    const data = await response.json();
    if (!response.ok) {
      throw new Error(data.error || "Failed to load data.");
    }
    buildLineChart(stuckupTrendChartEl, data.headers || [], data.rows || []);
  } catch (_error) {
    stuckupTrendChartEl.replaceWith(
      document.createTextNode("No data available")
    );
  }
}

loadRegionalSummary();
loadStuckupAnalysis();
loadAgeingBucket();
loadTopHubs();
loadValidationTrend();
loadStuckupTrend();
