const table = document.querySelector("#regional-table");
const thead = table.querySelector("thead");
const tbody = table.querySelector("tbody");
const statusEl = document.querySelector("#last-updated");

function parseNumber(value) {
  if (value === null || value === undefined || value === "") {
    return null;
  }
  const cleaned = String(value).replace(/,/g, "");
  const parsed = Number.parseFloat(cleaned);
  return Number.isFinite(parsed) ? parsed : null;
}

function buildTable(headers, rows) {
  thead.innerHTML = "";
  tbody.innerHTML = "";

  const headerRow = document.createElement("tr");
  headers.forEach((label) => {
    const th = document.createElement("th");
    th.textContent = label;
    headerRow.appendChild(th);
  });
  thead.appendChild(headerRow);

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
    tbody.appendChild(tr);
  });
}

async function loadRegionalSummary() {
  statusEl.textContent = "Loading…";
  try {
    const response = await fetch("/api/regional-validation");
    const data = await response.json();
    if (!response.ok) {
      throw new Error(data.error || "Failed to load data.");
    }
    buildTable(data.headers || [], data.rows || []);
    statusEl.textContent = `Last sync ${new Date().toLocaleString()}`;
  } catch (error) {
    const message =
      error && typeof error.message === "string"
        ? error.message
        : "Failed to load.";
    statusEl.textContent = message;
  }
}

loadRegionalSummary();
