const STORAGE_KEY = "kpi-tracker-v1";
const SETTINGS_KEY = "kpi-tracker-sheet-settings";
const MONTHS = [
  { id: "2026-01", label: "Jan 2026", longLabel: "January 2026" },
  { id: "2026-02", label: "Feb 2026", longLabel: "February 2026" },
  { id: "2026-03", label: "Mar 2026", longLabel: "March 2026" },
  { id: "2026-04", label: "Apr 2026", longLabel: "April 2026" },
  { id: "2026-05", label: "May 2026", longLabel: "May 2026" },
  { id: "2026-06", label: "Jun 2026", longLabel: "June 2026" },
  { id: "2026-07", label: "Jul 2026", longLabel: "July 2026" },
  { id: "2026-08", label: "Aug 2026", longLabel: "August 2026" },
  { id: "2026-09", label: "Sep 2026", longLabel: "September 2026" },
  { id: "2026-10", label: "Oct 2026", longLabel: "October 2026" },
  { id: "2026-11", label: "Nov 2026", longLabel: "November 2026" },
  { id: "2026-12", label: "Dec 2026", longLabel: "December 2026" }
];

const CLUBS = [
  { id: "newcastle", name: "Newcastle", short: "Newy", color: "#a855f7" },
  { id: "kotara", name: "Kotara", short: "Kotara", color: "#8b5cf6" },
  { id: "edgeworth", name: "Edgeworth", short: "Edgie", color: "#d946ef" }
];

const DEFAULT_INPUTS = {
  sales: 0,
  leads: 0,
  appointmentsBooked: 0,
  appointmentsCompleted: 0,
  totalAppointments: 0,
  monthlySalesTarget: 0,
  revenue: 0
};

const DEFAULT_SHEET_SETTINGS = CLUBS.reduce((acc, club) => {
  acc[club.id] = {
    csvUrl: "",
    membersMapping: "sales",
    monthColumn: "",
    membersColumn: "",
    incomeColumn: ""
  };
  return acc;
}, {});

const appState = {
  activeView: "group",
  activeClub: CLUBS[0].id,
  activeClubSection: "totals",
  activeMonth: MONTHS[0].id,
  data: loadState(),
  sheetSettings: loadSheetSettings(),
  charts: {}
};

const clubTabsEl = document.getElementById("club-tabs");
const groupViewEl = document.getElementById("group-view");
const clubViewEl = document.getElementById("club-view");
const saveStatusEl = document.getElementById("save-status");
const settingsDrawerEl = document.getElementById("settings-drawer");
const settingsFormEl = document.getElementById("settings-form");

document.getElementById("open-settings").addEventListener("click", () => toggleSettings(true));
document.getElementById("close-settings").addEventListener("click", () => toggleSettings(false));
document.getElementById("sync-sheets").addEventListener("click", syncAllSheets);
settingsDrawerEl.addEventListener("click", (event) => {
  if (event.target === settingsDrawerEl) {
    toggleSettings(false);
  }
});

renderApp();

function loadState() {
  const saved = safeParse(localStorage.getItem(STORAGE_KEY), {});
  return CLUBS.reduce((clubsAcc, club) => {
    clubsAcc[club.id] = MONTHS.reduce((monthAcc, month) => {
      monthAcc[month.id] = { ...DEFAULT_INPUTS, ...(saved?.[club.id]?.[month.id] || {}) };
      return monthAcc;
    }, {});
    return clubsAcc;
  }, {});
}

function loadSheetSettings() {
  return { ...DEFAULT_SHEET_SETTINGS, ...safeParse(localStorage.getItem(SETTINGS_KEY), {}) };
}

function safeParse(value, fallback) {
  try {
    return value ? JSON.parse(value) : fallback;
  } catch (error) {
    return fallback;
  }
}

function persistState() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(appState.data));
  flashSaveStatus("Saved locally");
}

function persistSheetSettings() {
  localStorage.setItem(SETTINGS_KEY, JSON.stringify(appState.sheetSettings));
  flashSaveStatus("Settings saved");
}

function flashSaveStatus(message) {
  saveStatusEl.textContent = message;
  window.clearTimeout(flashSaveStatus.timer);
  flashSaveStatus.timer = window.setTimeout(() => {
    saveStatusEl.textContent = "Ready";
  }, 1800);
}

function renderApp() {
  renderTopTabs();
  renderGroupView();
  renderClubView();
  renderSettings();
}

function renderTopTabs() {
  const tabs = [
    { id: "group", label: "Group Totals" },
    ...CLUBS.map((club) => ({ id: club.id, label: club.name }))
  ];

  clubTabsEl.innerHTML = tabs.map((tab) => `
    <button
      type="button"
      class="tab-button ${appState.activeView === tab.id ? "active" : ""}"
      data-view="${tab.id}"
    >${tab.label}</button>
  `).join("");

  clubTabsEl.querySelectorAll("[data-view]").forEach((button) => {
    button.addEventListener("click", () => {
      const nextView = button.dataset.view;
      appState.activeView = nextView;
      if (nextView !== "group") {
        appState.activeClub = nextView;
      }
      renderApp();
    });
  });
}

function renderGroupView() {
  const monthData = getGroupMonthSummary(appState.activeMonth);
  const yearData = getGroupYearSummary();

  groupViewEl.classList.toggle("hidden", appState.activeView !== "group");
  groupViewEl.innerHTML = `
    <div class="panel-grid">
      <section class="hero-card">
        <div class="hero-grid">
          <div class="section-stack">
            <div class="card-header">
              <div>
                <p class="eyebrow">Group Overview</p>
                <h2 class="card-title">${getMonthLabel(appState.activeMonth)} Funnel Snapshot</h2>
              </div>
              <span class="subtle-pill">3 Clubs Combined</span>
            </div>
            <p class="card-copy">
              Total funnel performance across Newcastle, Kotara, and Edgeworth. Local entries and any synced sheet data are blended together.
            </p>
            <div class="month-strip">${renderMonthButtons(appState.activeMonth)}</div>
          </div>
          <div class="summary-box">
            <div class="summary-label">Combined Revenue</div>
            <div class="summary-value">${formatCurrency(monthData.revenue)}</div>
            <p class="helper-text">Income synced from Google Sheets lands here when available.</p>
          </div>
        </div>
      </section>

      <section class="card">
        <div class="metric-grid">
          ${renderSummaryMetric("Total Sales", monthData.sales)}
          ${renderSummaryMetric("Total Leads", monthData.leads)}
          ${renderSummaryMetric("Appointments Booked", monthData.appointmentsBooked)}
          ${renderSummaryMetric("Appointments Completed", monthData.appointmentsCompleted)}
          ${renderSummaryMetric("Lead to Sale %", formatPercent(monthData.rates.leadToSale))}
          ${renderSummaryMetric("Appointment to Sale %", formatPercent(monthData.rates.appointmentToSale))}
        </div>
      </section>

      <section class="card">
        <div class="card-header">
          <div>
            <h3 class="card-title">Club Comparison</h3>
            <p class="card-copy">Performance breakdown for the selected month.</p>
          </div>
        </div>
        <table class="comparison-table">
          <thead>
            <tr>
              <th>Club</th>
              <th>Sales</th>
              <th>Leads</th>
              <th>Lead to Sale</th>
              <th>Booked</th>
              <th>Completed</th>
              <th>Revenue</th>
            </tr>
          </thead>
          <tbody>
            ${CLUBS.map((club) => {
              const month = getClubMonthData(club.id, appState.activeMonth);
              const summary = calculateSummary(month);
              return `
                <tr>
                  <td>${club.name}</td>
                  <td>${summary.sales}</td>
                  <td>${summary.leads}</td>
                  <td>${formatPercent(summary.rates.leadToSale)}</td>
                  <td>${summary.appointmentsBooked}</td>
                  <td>${summary.appointmentsCompleted}</td>
                  <td>${formatCurrency(summary.revenue)}</td>
                </tr>
              `;
            }).join("")}
          </tbody>
        </table>
      </section>

      <section class="chart-grid">
        <article class="chart-card">
          <div class="card-header">
            <div>
              <h3 class="card-title">Group Funnel Performance</h3>
              <p class="card-copy">Sales funnel volume for ${getMonthLabel(appState.activeMonth)}.</p>
            </div>
          </div>
          <div class="chart-wrap"><canvas id="group-funnel-chart"></canvas></div>
        </article>
        <article class="chart-card">
          <div class="card-header">
            <div>
              <h3 class="card-title">Monthly Sales Trend</h3>
              <p class="card-copy">Total sales across the full 2026 view.</p>
            </div>
          </div>
          <div class="chart-wrap"><canvas id="group-trend-chart"></canvas></div>
        </article>
        <article class="chart-card">
          <div class="card-header">
            <div>
              <h3 class="card-title">Club Sales Comparison</h3>
              <p class="card-copy">Selected month sales by club.</p>
            </div>
          </div>
          <div class="chart-wrap"><canvas id="group-comparison-chart"></canvas></div>
        </article>
      </section>
    </div>
  `;

  attachMonthButtonEvents(groupViewEl);
  drawOrReplaceChart("group-funnel-chart", buildFunnelChartConfig(monthData, "Group Funnel"));
  drawOrReplaceChart("group-trend-chart", buildTrendChartConfig(yearData.months, "sales", "Group Sales Trend"));
  drawOrReplaceChart("group-comparison-chart", buildGroupComparisonConfig(appState.activeMonth));
}

function renderClubView() {
  const club = CLUBS.find((entry) => entry.id === appState.activeClub) || CLUBS[0];

  clubViewEl.classList.toggle("hidden", appState.activeView === "group");
  clubViewEl.innerHTML = `
    <div class="panel-grid">
      <section class="hero-card section-stack">
        <div class="toolbar">
          <div>
            <p class="eyebrow">${club.short}</p>
            <h2 class="card-title">${club.name} KPI Tracker</h2>
          </div>
          <span class="subtle-pill">2026 Tracking</span>
        </div>
        <div class="sub-tabs">
          <button type="button" class="sub-tab-button ${appState.activeClubSection === "totals" ? "active" : ""}" data-section="totals">Club Totals</button>
          <button type="button" class="sub-tab-button ${appState.activeClubSection === "months" ? "active" : ""}" data-section="months">Monthly Tabs</button>
        </div>
      </section>

      ${appState.activeClubSection === "totals" ? renderClubTotals(club) : renderClubMonth(club)}
    </div>
  `;

  clubViewEl.querySelectorAll("[data-section]").forEach((button) => {
    button.addEventListener("click", () => {
      appState.activeClubSection = button.dataset.section;
      renderClubView();
    });
  });

  if (appState.activeClubSection === "months") {
    wireInputs(club.id, appState.activeMonth);
    attachMonthButtonEvents(clubViewEl);
  }
}

function renderClubTotals(club) {
  const months = MONTHS.map((month) => ({
    monthLabel: month.label,
    ...calculateSummary(getClubMonthData(club.id, month.id))
  }));
  const annual = aggregateSummaries(months);

  window.setTimeout(() => {
    drawOrReplaceChart(`club-total-trend-${club.id}`, buildTrendChartConfig(months, "sales", `${club.name} Sales Trend`, club.color));
    drawOrReplaceChart(`club-total-conversion-${club.id}`, buildConversionChartConfig(annual, club.color));
  }, 0);

  return `
    <section class="card">
      <div class="club-total-grid">
        ${renderSummaryMetric("Annual Sales", annual.sales)}
        ${renderSummaryMetric("Annual Leads", annual.leads)}
        ${renderSummaryMetric("Annual Revenue", formatCurrency(annual.revenue))}
        ${renderSummaryMetric("Lead to Sale %", formatPercent(annual.rates.leadToSale))}
        ${renderSummaryMetric("Lead to Appointment %", formatPercent(annual.rates.leadToAppointment))}
        ${renderSummaryMetric("Appointment Show %", formatPercent(annual.rates.appointmentShow))}
      </div>
    </section>
    <section class="chart-grid">
      <article class="chart-card">
        <div class="card-header">
          <div>
            <h3 class="card-title">${club.name} Sales Trend</h3>
            <p class="card-copy">Year-to-date monthly sales performance.</p>
          </div>
        </div>
        <div class="chart-wrap"><canvas id="club-total-trend-${club.id}"></canvas></div>
      </article>
      <article class="chart-card">
        <div class="card-header">
          <div>
            <h3 class="card-title">${club.name} Conversion Mix</h3>
            <p class="card-copy">Annual average conversion profile.</p>
          </div>
        </div>
        <div class="chart-wrap"><canvas id="club-total-conversion-${club.id}"></canvas></div>
      </article>
    </section>
  `;
}

function renderClubMonth(club) {
  const monthData = getClubMonthData(club.id, appState.activeMonth);
  const summary = calculateSummary(monthData);
  const targets = calculateTargets(summary);

  window.setTimeout(() => {
    const trendRows = MONTHS.map((month) => ({
      monthLabel: month.label,
      ...calculateSummary(getClubMonthData(club.id, month.id))
    }));
    drawOrReplaceChart(`club-funnel-${club.id}`, buildFunnelChartConfig(summary, `${club.name} Funnel`, club.color));
    drawOrReplaceChart(`club-trend-${club.id}`, buildTrendChartConfig(trendRows, "sales", `${club.name} Sales Trend`, club.color));
    drawOrReplaceChart(`club-conversion-${club.id}`, buildConversionChartConfig(summary, club.color));
  }, 0);

  return `
    <section class="section-stack">
      <div class="month-strip">${renderMonthButtons(appState.activeMonth)}</div>

      <article class="card">
        <div class="card-header">
          <div>
            <h3 class="card-title">${club.name} Inputs</h3>
            <p class="card-copy">Editable KPI inputs for ${getMonthLabel(appState.activeMonth)}. Values auto-save per club and month.</p>
          </div>
        </div>
        <div class="input-grid">
          ${renderInputField("sales", "Sales", monthData.sales)}
          ${renderInputField("leads", "Leads", monthData.leads)}
          ${renderInputField("appointmentsBooked", "Appointments Booked", monthData.appointmentsBooked)}
          ${renderInputField("appointmentsCompleted", "Appointments Completed", monthData.appointmentsCompleted)}
          ${renderInputField("totalAppointments", "Total Appointments", monthData.totalAppointments)}
          ${renderInputField("monthlySalesTarget", "Monthly Sales Target", monthData.monthlySalesTarget)}
          ${renderInputField("revenue", "Revenue", monthData.revenue)}
        </div>
      </article>

      <article class="card">
        <div class="card-header">
          <div>
            <h3 class="card-title">Conversion Metrics</h3>
            <p class="card-copy">Live funnel math based on the month's saved inputs.</p>
          </div>
        </div>
        <div class="metric-grid">
          ${renderSummaryMetric("Lead to Sale %", formatPercent(summary.rates.leadToSale))}
          ${renderSummaryMetric("Lead to Appointment %", formatPercent(summary.rates.leadToAppointment))}
          ${renderSummaryMetric("Appointment Show %", formatPercent(summary.rates.appointmentShow))}
          ${renderSummaryMetric("Appointment to Sale %", formatPercent(summary.rates.appointmentToSale))}
        </div>
      </article>

      <article class="card">
        <div class="card-header">
          <div>
            <h3 class="card-title">Targets Breakdown</h3>
            <p class="card-copy">Required activity to hit the monthly sales target using current conversion rates with benchmark fallbacks.</p>
          </div>
        </div>
        <div class="target-grid">
          ${renderTargetMetric("Required Leads", targets.monthly.leads)}
          ${renderTargetMetric("Required Appointments", targets.monthly.appointmentsBooked)}
          ${renderTargetMetric("Required Presentations", targets.monthly.presentations)}
          ${renderTargetMetric("Weekly Sales Target", targets.weekly.sales)}
          ${renderTargetMetric("Daily Sales Target", targets.daily.sales)}
          ${renderTargetMetric("Weekly Leads", targets.weekly.leads)}
          ${renderTargetMetric("Daily Leads", targets.daily.leads)}
          ${renderTargetMetric("Weekly Appointments", targets.weekly.appointmentsBooked)}
          ${renderTargetMetric("Daily Presentations", targets.daily.presentations)}
        </div>
      </article>

      <section class="chart-grid">
        <article class="chart-card">
          <div class="card-header">
            <div>
              <h3 class="card-title">Funnel Performance</h3>
              <p class="card-copy">Lead to sale volume for ${getMonthLabel(appState.activeMonth)}.</p>
            </div>
          </div>
          <div class="chart-wrap"><canvas id="club-funnel-${club.id}"></canvas></div>
        </article>
        <article class="chart-card">
          <div class="card-header">
            <div>
              <h3 class="card-title">Monthly Trend</h3>
              <p class="card-copy">Sales across all 2026 tabs for ${club.name}.</p>
            </div>
          </div>
          <div class="chart-wrap"><canvas id="club-trend-${club.id}"></canvas></div>
        </article>
        <article class="chart-card">
          <div class="card-header">
            <div>
              <h3 class="card-title">Conversion Rates</h3>
              <p class="card-copy">A quick read on the current conversion profile.</p>
            </div>
          </div>
          <div class="chart-wrap"><canvas id="club-conversion-${club.id}"></canvas></div>
        </article>
      </section>
    </section>
  `;
}

function wireInputs(clubId, monthId) {
  clubViewEl.querySelectorAll("[data-input]").forEach((input) => {
    input.addEventListener("input", (event) => {
      const field = event.target.dataset.input;
      const nextValue = Number(event.target.value) || 0;
      appState.data[clubId][monthId][field] = nextValue;
      persistState();
    });
  });
}

function renderInputField(field, label, value) {
  return `
    <div class="field">
      <label for="${field}">${label}</label>
      <input id="${field}" type="number" min="0" step="1" value="${value}" data-input="${field}">
    </div>
  `;
}

function renderSummaryMetric(label, value) {
  return `
    <div class="metric-box">
      <div class="metric-label">${label}</div>
      <div class="metric-value">${value}</div>
    </div>
  `;
}

function renderTargetMetric(label, value) {
  return `
    <div class="target-box">
      <div class="target-label">${label}</div>
      <div class="target-value">${value}</div>
    </div>
  `;
}

function renderMonthButtons(activeMonth) {
  return MONTHS.map((month) => `
    <button type="button" class="month-button ${activeMonth === month.id ? "active" : ""}" data-month="${month.id}">
      ${month.label}
    </button>
  `).join("");
}

function attachMonthButtonEvents(scope) {
  scope.querySelectorAll("[data-month]").forEach((button) => {
    button.addEventListener("click", () => {
      appState.activeMonth = button.dataset.month;
      renderApp();
    });
  });
}

function getClubMonthData(clubId, monthId) {
  return appState.data[clubId][monthId];
}

function calculateSummary(monthData) {
  const safe = { ...DEFAULT_INPUTS, ...monthData };
  return {
    ...safe,
    rates: {
      leadToSale: divide(safe.sales, safe.leads),
      leadToAppointment: divide(safe.appointmentsBooked, safe.leads),
      appointmentShow: divide(safe.appointmentsCompleted, safe.appointmentsBooked),
      appointmentToSale: divide(safe.sales, safe.appointmentsCompleted)
    }
  };
}

function calculateTargets(summary) {
  const targetSales = summary.monthlySalesTarget || 0;
  const assumedLeadToAppointment = summary.rates.leadToAppointment || 0.65;
  const assumedAppointmentShow = summary.rates.appointmentShow || 0.7;
  const assumedAppointmentToSale = summary.rates.appointmentToSale || 0.45;

  const requiredPresentations = targetSales > 0 ? Math.ceil(targetSales / assumedAppointmentToSale) : 0;
  const requiredAppointmentsBooked = requiredPresentations > 0 ? Math.ceil(requiredPresentations / assumedAppointmentShow) : 0;
  const requiredLeads = requiredAppointmentsBooked > 0 ? Math.ceil(requiredAppointmentsBooked / assumedLeadToAppointment) : 0;

  return {
    monthly: {
      leads: requiredLeads,
      appointmentsBooked: requiredAppointmentsBooked,
      presentations: requiredPresentations,
      sales: targetSales
    },
    weekly: {
      leads: roundToSingleDecimal(requiredLeads / 4.33),
      appointmentsBooked: roundToSingleDecimal(requiredAppointmentsBooked / 4.33),
      presentations: roundToSingleDecimal(requiredPresentations / 4.33),
      sales: roundToSingleDecimal(targetSales / 4.33)
    },
    daily: {
      leads: roundToSingleDecimal(requiredLeads / 21.67),
      appointmentsBooked: roundToSingleDecimal(requiredAppointmentsBooked / 21.67),
      presentations: roundToSingleDecimal(requiredPresentations / 21.67),
      sales: roundToSingleDecimal(targetSales / 21.67)
    }
  };
}

function aggregateSummaries(summaries) {
  const totals = summaries.reduce((acc, item) => {
    acc.sales += item.sales;
    acc.leads += item.leads;
    acc.appointmentsBooked += item.appointmentsBooked;
    acc.appointmentsCompleted += item.appointmentsCompleted;
    acc.totalAppointments += item.totalAppointments;
    acc.monthlySalesTarget += item.monthlySalesTarget;
    acc.revenue += item.revenue;
    return acc;
  }, { ...DEFAULT_INPUTS });

  return calculateSummary(totals);
}

function getGroupMonthSummary(monthId) {
  return aggregateSummaries(CLUBS.map((club) => calculateSummary(getClubMonthData(club.id, monthId))));
}

function getGroupYearSummary() {
  return {
    months: MONTHS.map((month) => ({
      monthLabel: month.label,
      ...getGroupMonthSummary(month.id)
    }))
  };
}

function divide(numerator, denominator) {
  if (!denominator) {
    return 0;
  }
  return numerator / denominator;
}

function formatPercent(value) {
  return `${(value * 100).toFixed(1)}%`;
}

function formatCurrency(value) {
  return new Intl.NumberFormat("en-AU", {
    style: "currency",
    currency: "AUD",
    maximumFractionDigits: 0
  }).format(value || 0);
}

function roundToSingleDecimal(value) {
  return Number((value || 0).toFixed(1));
}

function getMonthLabel(monthId) {
  return MONTHS.find((month) => month.id === monthId)?.longLabel || monthId;
}

function renderSettings() {
  settingsFormEl.innerHTML = CLUBS.map((club) => {
    const config = appState.sheetSettings[club.id];
    return `
      <section class="settings-club">
        <h3>${club.name}</h3>
        <div class="field">
          <label for="csv-${club.id}">Published Sheet URL</label>
          <input id="csv-${club.id}" type="url" value="${config.csvUrl}" data-setting-club="${club.id}" data-setting-field="csvUrl" placeholder="https://docs.google.com/spreadsheets/d/e/.../pubhtml">
        </div>
        <div class="field">
          <label for="mapping-${club.id}">Members Mapping</label>
          <select id="mapping-${club.id}" data-setting-club="${club.id}" data-setting-field="membersMapping">
            <option value="sales" ${config.membersMapping === "sales" ? "selected" : ""}>Members to Sales</option>
            <option value="leads" ${config.membersMapping === "leads" ? "selected" : ""}>Members to Leads</option>
          </select>
        </div>
        <div class="field">
          <label for="month-${club.id}">Month Column Override</label>
          <input id="month-${club.id}" type="text" value="${config.monthColumn}" data-setting-club="${club.id}" data-setting-field="monthColumn" placeholder="Optional, e.g. Month">
        </div>
        <div class="field">
          <label for="members-${club.id}">Members Column Override</label>
          <input id="members-${club.id}" type="text" value="${config.membersColumn}" data-setting-club="${club.id}" data-setting-field="membersColumn" placeholder="Optional, e.g. Members">
        </div>
        <div class="field">
          <label for="income-${club.id}">Income Column Override</label>
          <input id="income-${club.id}" type="text" value="${config.incomeColumn}" data-setting-club="${club.id}" data-setting-field="incomeColumn" placeholder="Optional, e.g. Income">
        </div>
      </section>
    `;
  }).join("");

  settingsFormEl.querySelectorAll("[data-setting-field]").forEach((input) => {
    input.addEventListener("input", (event) => {
      const clubId = event.target.dataset.settingClub;
      const field = event.target.dataset.settingField;
      appState.sheetSettings[clubId][field] = event.target.value;
      persistSheetSettings();
    });
  });
}

function toggleSettings(open) {
  settingsDrawerEl.classList.toggle("hidden", !open);
  settingsDrawerEl.setAttribute("aria-hidden", String(!open));
}

async function syncAllSheets() {
  const button = document.getElementById("sync-sheets");
  button.disabled = true;
  button.textContent = "Syncing...";

  try {
    for (const club of CLUBS) {
      const config = appState.sheetSettings[club.id];
      if (!config.csvUrl) {
        continue;
      }
      const rows = await fetchCsvRows(config.csvUrl);
      mergeSheetRowsIntoState(club.id, rows, config);
    }
    persistState();
    renderApp();
    flashSaveStatus("Sheets synced");
  } catch (error) {
    flashSaveStatus("Sync failed");
    window.alert(`Sheet sync failed: ${error.message}`);
  } finally {
    button.disabled = false;
    button.textContent = "Sync From Sheets";
  }
}

async function fetchCsvRows(url) {
  // Normalise URL: accept both pubhtml and pub?output=csv, always fetch as HTML
  const htmlUrl = url
    .replace(/pub\?output=csv.*$/, "pubhtml")
    .replace(/pubhtml.*$/, "pubhtml");

  const response = await fetch(htmlUrl);
  if (!response.ok) {
    throw new Error(`Unable to fetch sheet (${response.status})`);
  }
  const text = await response.text();
  return parseSheetHtml(text);
}

function parseSheetHtml(html) {
  const parser = new DOMParser();
  const doc = parser.parseFromString(html, "text/html");
  const table = doc.querySelector("table");
  if (!table) {
    return [];
  }

  const rows = Array.from(table.querySelectorAll("tr"));
  if (rows.length < 2) {
    return [];
  }

  const headers = Array.from(rows[0].querySelectorAll("th, td"))
    .map((cell) => cell.textContent.trim());

  return rows.slice(1).map((row) => {
    const cells = Array.from(row.querySelectorAll("th, td"));
    return headers.reduce((obj, header, index) => {
      obj[header] = (cells[index] ? cells[index].textContent.trim() : "");
      return obj;
    }, {});
  }).filter((row) => Object.values(row).some((v) => v !== ""));
}

function mergeSheetRowsIntoState(clubId, rows, config) {
  if (!rows.length) {
    return;
  }

  const headerKeys = Object.keys(rows[0]);
  const monthKey = pickHeader(headerKeys, config.monthColumn, ["month", "period", "date"]);
  const membersKey = pickHeader(headerKeys, config.membersColumn, ["members", "member", "sales", "joins"]);
  const incomeKey = pickHeader(headerKeys, config.incomeColumn, ["income", "revenue", "amount", "takings"]);

  MONTHS.forEach((month) => {
    const matchingRow = rows.find((row) => normalizeMonth(row[monthKey]) === month.id);
    if (!matchingRow) {
      return;
    }

    const membersValue = parseNumber(matchingRow[membersKey]);
    const incomeValue = parseNumber(matchingRow[incomeKey]);

    if (config.membersMapping === "leads") {
      appState.data[clubId][month.id].leads = membersValue;
    } else {
      appState.data[clubId][month.id].sales = membersValue;
    }

    appState.data[clubId][month.id].revenue = incomeValue;
  });
}

function pickHeader(headers, override, candidates) {
  if (override && headers.includes(override)) {
    return override;
  }

  const match = headers.find((header) => candidates.some((candidate) => header.toLowerCase().includes(candidate)));
  return match || headers[0];
}

function parseNumber(value) {
  return Number(String(value || "").replace(/[^0-9.-]/g, "")) || 0;
}

function normalizeMonth(value) {
  const raw = String(value || "").trim();
  if (!raw) {
    return "";
  }

  const direct = MONTHS.find((month) => month.id === raw || month.label.toLowerCase() === raw.toLowerCase() || month.longLabel.toLowerCase() === raw.toLowerCase());
  if (direct) {
    return direct.id;
  }

  const parsed = new Date(raw);
  if (!Number.isNaN(parsed.getTime())) {
    const month = String(parsed.getMonth() + 1).padStart(2, "0");
    return `${parsed.getFullYear()}-${month}`;
  }

  return "";
}

function buildFunnelChartConfig(summary, label, color = "#a855f7") {
  return {
    type: "bar",
    data: {
      labels: ["Leads", "Booked", "Completed", "Sales"],
      datasets: [{
        label,
        data: [summary.leads, summary.appointmentsBooked, summary.appointmentsCompleted, summary.sales],
        backgroundColor: [withAlpha(color, 0.9), withAlpha(color, 0.7), withAlpha(color, 0.55), withAlpha(color, 1)],
        borderRadius: 14
      }]
    },
    options: sharedChartOptions()
  };
}

function buildTrendChartConfig(rows, field, label, color = "#a855f7") {
  return {
    type: "line",
    data: {
      labels: rows.map((row) => row.monthLabel || ""),
      datasets: [{
        label,
        data: rows.map((row) => row[field]),
        borderColor: color,
        backgroundColor: withAlpha(color, 0.2),
        fill: true,
        tension: 0.35,
        pointRadius: 4,
        pointBackgroundColor: color
      }]
    },
    options: sharedChartOptions()
  };
}

function buildConversionChartConfig(summary, color = "#a855f7") {
  return {
    type: "doughnut",
    data: {
      labels: ["Lead to Sale", "Lead to Appointment", "Appointment Show", "Appointment to Sale"],
      datasets: [{
        data: [
          percentageNumber(summary.rates.leadToSale),
          percentageNumber(summary.rates.leadToAppointment),
          percentageNumber(summary.rates.appointmentShow),
          percentageNumber(summary.rates.appointmentToSale)
        ],
        backgroundColor: [
          withAlpha(color, 1),
          withAlpha(color, 0.82),
          withAlpha(color, 0.64),
          withAlpha(color, 0.46)
        ],
        borderWidth: 0
      }]
    },
    options: {
      ...sharedChartOptions(),
      cutout: "64%"
    }
  };
}

function buildGroupComparisonConfig(monthId) {
  const rows = CLUBS.map((club) => ({
    club: club.name,
    color: club.color,
    sales: calculateSummary(getClubMonthData(club.id, monthId)).sales
  }));

  return {
    type: "bar",
    data: {
      labels: rows.map((row) => row.club),
      datasets: [{
        label: "Sales",
        data: rows.map((row) => row.sales),
        backgroundColor: rows.map((row) => withAlpha(row.color, 0.85)),
        borderRadius: 14
      }]
    },
    options: sharedChartOptions()
  };
}

function sharedChartOptions() {
  return {
    responsive: true,
    maintainAspectRatio: false,
    animation: false,
    plugins: {
      legend: {
        labels: {
          color: "#d8d8e1"
        }
      }
    },
    scales: {
      x: {
        ticks: { color: "#9d9daf" },
        grid: { color: "rgba(255,255,255,0.05)" }
      },
      y: {
        beginAtZero: true,
        ticks: { color: "#9d9daf" },
        grid: { color: "rgba(255,255,255,0.05)" }
      }
    }
  };
}

function drawOrReplaceChart(canvasId, config) {
  const canvas = document.getElementById(canvasId);
  if (!canvas) {
    return;
  }
  if (appState.charts[canvasId]) {
    appState.charts[canvasId].destroy();
  }
  appState.charts[canvasId] = new Chart(canvas, config);
}

function withAlpha(hex, alpha) {
  const parsed = hex.replace("#", "");
  const bigint = Number.parseInt(parsed, 16);
  const r = (bigint >> 16) & 255;
  const g = (bigint >> 8) & 255;
  const b = bigint & 255;
  return `rgba(${r}, ${g}, ${b}, ${alpha})`;
}

function percentageNumber(value) {
  return Number((value * 100).toFixed(1));
}
