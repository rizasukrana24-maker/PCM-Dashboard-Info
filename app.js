const monthSelect = document.getElementById("monthSelect");
const realtimeDateLabel = document.getElementById("realtimeDateLabel");

const kpiYearTotal = document.getElementById("kpiYearTotal");
const kpiYearTotalUtara = document.getElementById("kpiYearTotalUtara");
const kpiYearTotalSelatan = document.getElementById("kpiYearTotalSelatan");
const kpiYearDomesticInline = document.getElementById("kpiYearDomesticInline");
const kpiYearForeignInline = document.getElementById("kpiYearForeignInline");
const kpiYearLabel = document.getElementById("kpiYearLabel");

const kpiMonthTotal = document.getElementById("kpiMonthTotal");
const kpiMonthTotalUtara = document.getElementById("kpiMonthTotalUtara");
const kpiMonthTotalSelatan = document.getElementById("kpiMonthTotalSelatan");
const kpiMonthDomesticInline = document.getElementById("kpiMonthDomesticInline");
const kpiMonthForeignInline = document.getElementById("kpiMonthForeignInline");
const kpiMonthLabel = document.getElementById("kpiMonthLabel");

const kpiDayTotal = document.getElementById("kpiDayTotal");
const kpiDayTotalUtara = document.getElementById("kpiDayTotalUtara");
const kpiDayTotalSelatan = document.getElementById("kpiDayTotalSelatan");
const kpiDayDomesticInline = document.getElementById("kpiDayDomesticInline");
const kpiDayForeignInline = document.getElementById("kpiDayForeignInline");
const kpiAverageTotal = document.getElementById("kpiAverageTotal");
const kpiAverageYearTotal = document.getElementById("kpiAverageYearTotal");
const invoiceTotalValue = document.getElementById("invoiceTotalValue");
const invoiceTotalsStats = document.getElementById("invoiceTotalsStats");

const monthlyRecapTable = document.getElementById("monthlyRecapTable");
const externalRecapTables = document.getElementById("externalRecapTables");
const SHOW_COMBINED_RECAP_TOTAL_ROWS = false;

const manualInvoiceTotals = [
  {
    period: "2026-01",
    label: "Januari 2026",
    filledCount: 947,
    domesticCount: 769,
    foreignCount: 178,
    domesticPandu: 999437773,
    domesticTunda: 5787572079,
    foreignPandu: 895017683,
    foreignTunda: 9988608550,
    specialPandu: 0,
    specialTunda: 0,
    settledTotal: 10203515963,
    paidPandu: 1140357177,
    paidTunda: 9063158786,
    paidDomestic: 4679003215,
    paidForeign: 5524512748,
    paidSpecial: 0,
  },
  {
    period: "2026-02",
    label: "Februari 2026",
    filledCount: 1050,
    domesticCount: 887,
    foreignCount: 159,
    domesticPandu: 1127814838,
    domesticTunda: 6716009722,
    foreignPandu: 1050369482,
    foreignTunda: 9908708730,
    specialPandu: 0,
    specialTunda: 371271412,
    settledTotal: 8346889975,
    paidPandu: 995999161,
    paidTunda: 7350890814,
    paidDomestic: 4170131577,
    paidForeign: 3863233720,
    paidSpecial: 313524678,
  },
  {
    period: "2026-03",
    label: "Maret 2026",
    filledCount: 455,
    domesticCount: 401,
    foreignCount: 54,
    domesticPandu: 492105703,
    domesticTunda: 2994797485,
    foreignPandu: 424391006,
    foreignTunda: 3699733565,
    specialPandu: 0,
    specialTunda: 0,
    settledTotal: 1636952165,
    paidPandu: 220383001,
    paidTunda: 1416569164,
    paidDomestic: 1105418822,
    paidForeign: 531533343,
    paidSpecial: 0,
  },
];

const manualDashboardData = [
  { tanggal: "2026-01-01", kantor: "utara", idr: 12, usd: 0, total: 12 },
  { tanggal: "2026-01-01", kantor: "selatan", idr: 5, usd: 6, total: 11 },
  { tanggal: "2026-01-02", kantor: "utara", idr: 23, usd: 0, total: 23 },
  { tanggal: "2026-01-02", kantor: "selatan", idr: 8, usd: 7, total: 15 },
  { tanggal: "2026-01-03", kantor: "utara", idr: 21, usd: 0, total: 21 },
  { tanggal: "2026-01-03", kantor: "selatan", idr: 7, usd: 8, total: 15 },
  { tanggal: "2026-01-04", kantor: "utara", idr: 21, usd: 0, total: 21 },
  { tanggal: "2026-01-04", kantor: "selatan", idr: 6, usd: 7, total: 13 },
  { tanggal: "2026-01-05", kantor: "utara", idr: 23, usd: 0, total: 23 },
  { tanggal: "2026-01-05", kantor: "selatan", idr: 2, usd: 3, total: 5 },
  { tanggal: "2026-01-06", kantor: "utara", idr: 30, usd: 1, total: 31 },
  { tanggal: "2026-01-06", kantor: "selatan", idr: 3, usd: 6, total: 9 },
  { tanggal: "2026-01-07", kantor: "utara", idr: 22, usd: 0, total: 22 },
  { tanggal: "2026-01-07", kantor: "selatan", idr: 4, usd: 4, total: 8 },
  { tanggal: "2026-01-08", kantor: "utara", idr: 22, usd: 0, total: 22 },
  { tanggal: "2026-01-08", kantor: "selatan", idr: 4, usd: 5, total: 9 },
  { tanggal: "2026-01-09", kantor: "utara", idr: 20, usd: 4, total: 24 },
  { tanggal: "2026-01-09", kantor: "selatan", idr: 7, usd: 5, total: 12 },
  { tanggal: "2026-01-10", kantor: "utara", idr: 22, usd: 1, total: 23 },
  { tanggal: "2026-01-10", kantor: "selatan", idr: 4, usd: 5, total: 9 },
  { tanggal: "2026-01-11", kantor: "utara", idr: 19, usd: 0, total: 19 },
  { tanggal: "2026-01-11", kantor: "selatan", idr: 3, usd: 9, total: 12 },
  { tanggal: "2026-01-12", kantor: "utara", idr: 19, usd: 0, total: 19 },
  { tanggal: "2026-01-12", kantor: "selatan", idr: 3, usd: 1, total: 4 },
  { tanggal: "2026-01-13", kantor: "utara", idr: 13, usd: 0, total: 13 },
  { tanggal: "2026-01-13", kantor: "selatan", idr: 3, usd: 5, total: 8 },
  { tanggal: "2026-01-14", kantor: "utara", idr: 30, usd: 0, total: 30 },
  { tanggal: "2026-01-14", kantor: "selatan", idr: 6, usd: 3, total: 9 },
  { tanggal: "2026-01-15", kantor: "utara", idr: 11, usd: 0, total: 11 },
  { tanggal: "2026-01-15", kantor: "selatan", idr: 4, usd: 1, total: 5 },
  { tanggal: "2026-01-16", kantor: "utara", idr: 20, usd: 0, total: 20 },
  { tanggal: "2026-01-16", kantor: "selatan", idr: 2, usd: 6, total: 8 },
  { tanggal: "2026-01-17", kantor: "utara", idr: 30, usd: 1, total: 31 },
  { tanggal: "2026-01-17", kantor: "selatan", idr: 6, usd: 5, total: 11 },
  { tanggal: "2026-01-18", kantor: "utara", idr: 19, usd: 1, total: 20 },
  { tanggal: "2026-01-18", kantor: "selatan", idr: 4, usd: 6, total: 10 },
  { tanggal: "2026-01-19", kantor: "utara", idr: 17, usd: 1, total: 18 },
  { tanggal: "2026-01-19", kantor: "selatan", idr: 1, usd: 3, total: 4 },
  { tanggal: "2026-01-20", kantor: "utara", idr: 26, usd: 0, total: 26 },
  { tanggal: "2026-01-20", kantor: "selatan", idr: 4, usd: 3, total: 7 },
  { tanggal: "2026-01-21", kantor: "utara", idr: 28, usd: 0, total: 28 },
  { tanggal: "2026-01-21", kantor: "selatan", idr: 3, usd: 5, total: 8 },
  { tanggal: "2026-01-22", kantor: "utara", idr: 16, usd: 0, total: 16 },
  { tanggal: "2026-01-22", kantor: "selatan", idr: 4, usd: 5, total: 9 },
  { tanggal: "2026-01-23", kantor: "utara", idr: 20, usd: 0, total: 20 },
  { tanggal: "2026-01-23", kantor: "selatan", idr: 3, usd: 12, total: 15 },
  { tanggal: "2026-01-24", kantor: "utara", idr: 12, usd: 1, total: 13 },
  { tanggal: "2026-01-24", kantor: "selatan", idr: 3, usd: 6, total: 9 },
  { tanggal: "2026-01-25", kantor: "utara", idr: 15, usd: 1, total: 16 },
  { tanggal: "2026-01-25", kantor: "selatan", idr: 3, usd: 16, total: 19 },
  { tanggal: "2026-01-26", kantor: "utara", idr: 26, usd: 1, total: 27 },
  { tanggal: "2026-01-26", kantor: "selatan", idr: 5, usd: 7, total: 12 },
  { tanggal: "2026-01-27", kantor: "utara", idr: 26, usd: 0, total: 26 },
  { tanggal: "2026-01-27", kantor: "selatan", idr: 4, usd: 7, total: 11 },
  { tanggal: "2026-01-28", kantor: "utara", idr: 20, usd: 2, total: 22 },
  { tanggal: "2026-01-28", kantor: "selatan", idr: 8, usd: 7, total: 15 },
  { tanggal: "2026-01-29", kantor: "utara", idr: 26, usd: 1, total: 27 },
  { tanggal: "2026-01-29", kantor: "selatan", idr: 4, usd: 3, total: 7 },
  { tanggal: "2026-01-30", kantor: "utara", idr: 18, usd: 0, total: 18 },
  { tanggal: "2026-01-30", kantor: "selatan", idr: 6, usd: 3, total: 9 },
  { tanggal: "2026-01-31", kantor: "utara", idr: 33, usd: 0, total: 33 },
  { tanggal: "2026-01-31", kantor: "selatan", idr: 3, usd: 6, total: 9 },
  { tanggal: "2026-02-01", kantor: "utara", idr: 20, usd: 0, total: 20 },
  { tanggal: "2026-02-01", kantor: "selatan", idr: 6, usd: 8, total: 14 },
  { tanggal: "2026-02-02", kantor: "utara", idr: 23, usd: 2, total: 25 },
  { tanggal: "2026-02-02", kantor: "selatan", idr: 6, usd: 6, total: 12 },
  { tanggal: "2026-02-03", kantor: "utara", idr: 34, usd: 2, total: 36 },
  { tanggal: "2026-02-03", kantor: "selatan", idr: 2, usd: 9, total: 11 },
  { tanggal: "2026-02-04", kantor: "utara", idr: 33, usd: 1, total: 34 },
  { tanggal: "2026-02-04", kantor: "selatan", idr: 2, usd: 4, total: 6 },
  { tanggal: "2026-02-05", kantor: "utara", idr: 24, usd: 1, total: 25 },
  { tanggal: "2026-02-05", kantor: "selatan", idr: 5, usd: 8, total: 13 },
  { tanggal: "2026-02-06", kantor: "utara", idr: 22, usd: 0, total: 22 },
  { tanggal: "2026-02-06", kantor: "selatan", idr: 8, usd: 5, total: 13 },
  { tanggal: "2026-02-07", kantor: "utara", idr: 33, usd: 1, total: 34 },
  { tanggal: "2026-02-07", kantor: "selatan", idr: 4, usd: 4, total: 8 },
  { tanggal: "2026-02-08", kantor: "utara", idr: 22, usd: 1, total: 23 },
  { tanggal: "2026-02-08", kantor: "selatan", idr: 4, usd: 5, total: 9 },
  { tanggal: "2026-02-09", kantor: "utara", idr: 23, usd: 0, total: 23 },
  { tanggal: "2026-02-09", kantor: "selatan", idr: 4, usd: 7, total: 11 },
  { tanggal: "2026-02-10", kantor: "utara", idr: 20, usd: 1, total: 21 },
  { tanggal: "2026-02-10", kantor: "selatan", idr: 3, usd: 5, total: 8 },
  { tanggal: "2026-02-11", kantor: "utara", idr: 28, usd: 0, total: 28 },
  { tanggal: "2026-02-11", kantor: "selatan", idr: 10, usd: 8, total: 18 },
  { tanggal: "2026-02-12", kantor: "utara", idr: 29, usd: 0, total: 29 },
  { tanggal: "2026-02-12", kantor: "selatan", idr: 3, usd: 3, total: 6 },
  { tanggal: "2026-02-13", kantor: "utara", idr: 45, usd: 0, total: 45 },
  { tanggal: "2026-02-13", kantor: "selatan", idr: 4, usd: 4, total: 8 },
  { tanggal: "2026-02-14", kantor: "utara", idr: 31, usd: 0, total: 31 },
  { tanggal: "2026-02-14", kantor: "selatan", idr: 3, usd: 8, total: 11 },
  { tanggal: "2026-02-15", kantor: "utara", idr: 20, usd: 1, total: 21 },
  { tanggal: "2026-02-15", kantor: "selatan", idr: 5, usd: 6, total: 11 },
  { tanggal: "2026-02-16", kantor: "utara", idr: 34, usd: 1, total: 35 },
  { tanggal: "2026-02-16", kantor: "selatan", idr: 6, usd: 5, total: 11 },
  { tanggal: "2026-02-17", kantor: "utara", idr: 21, usd: 0, total: 21 },
  { tanggal: "2026-02-17", kantor: "selatan", idr: 1, usd: 4, total: 5 },
  { tanggal: "2026-02-18", kantor: "utara", idr: 24, usd: 0, total: 24 },
  { tanggal: "2026-02-18", kantor: "selatan", idr: 6, usd: 5, total: 11 },
  { tanggal: "2026-02-19", kantor: "utara", idr: 24, usd: 0, total: 24 },
  { tanggal: "2026-02-19", kantor: "selatan", idr: 2, usd: 4, total: 6 },
  { tanggal: "2026-02-20", kantor: "utara", idr: 20, usd: 0, total: 20 },
  { tanggal: "2026-02-20", kantor: "selatan", idr: 6, usd: 9, total: 15 },
  { tanggal: "2026-02-21", kantor: "utara", idr: 25, usd: 0, total: 25 },
  { tanggal: "2026-02-21", kantor: "selatan", idr: 4, usd: 10, total: 14 },
  { tanggal: "2026-02-22", kantor: "utara", idr: 16, usd: 0, total: 16 },
  { tanggal: "2026-02-22", kantor: "selatan", idr: 4, usd: 12, total: 16 },
  { tanggal: "2026-02-23", kantor: "utara", idr: 23, usd: 0, total: 23 },
  { tanggal: "2026-02-23", kantor: "selatan", idr: 8, usd: 4, total: 12 },
  { tanggal: "2026-02-24", kantor: "utara", idr: 31, usd: 1, total: 32 },
  { tanggal: "2026-02-24", kantor: "selatan", idr: 2, usd: 12, total: 14 },
  { tanggal: "2026-02-25", kantor: "utara", idr: 22, usd: 1, total: 23 },
  { tanggal: "2026-02-25", kantor: "selatan", idr: 7, usd: 6, total: 13 },
  { tanggal: "2026-02-26", kantor: "utara", idr: 28, usd: 0, total: 28 },
  { tanggal: "2026-02-26", kantor: "selatan", idr: 7, usd: 6, total: 13 },
  { tanggal: "2026-02-27", kantor: "utara", idr: 16, usd: 0, total: 16 },
  { tanggal: "2026-02-27", kantor: "selatan", idr: 4, usd: 1, total: 5 },
  { tanggal: "2026-02-28", kantor: "utara", idr: 33, usd: 1, total: 34 },
  { tanggal: "2026-02-28", kantor: "selatan", idr: 4, usd: 4, total: 8 },
  { tanggal: "2026-03-01", kantor: "utara", idr: 19, usd: 1, total: 20 },
  { tanggal: "2026-03-01", kantor: "selatan", idr: 9, usd: 5, total: 14 },
  { tanggal: "2026-03-02", kantor: "utara", idr: 28, usd: 0, total: 28 },
  { tanggal: "2026-03-02", kantor: "selatan", idr: 6, usd: 6, total: 12 },
  { tanggal: "2026-03-03", kantor: "utara", idr: 23, usd: 0, total: 23 },
  { tanggal: "2026-03-03", kantor: "selatan", idr: 6, usd: 6, total: 12 },
  { tanggal: "2026-03-04", kantor: "utara", idr: 28, usd: 1, total: 29 },
  { tanggal: "2026-03-04", kantor: "selatan", idr: 6, usd: 2, total: 8 },
  { tanggal: "2026-03-05", kantor: "utara", idr: 38, usd: 0, total: 38 },
  { tanggal: "2026-03-05", kantor: "selatan", idr: 8, usd: 4, total: 12 },
  { tanggal: "2026-03-06", kantor: "utara", idr: 34, usd: 0, total: 34 },
  { tanggal: "2026-03-06", kantor: "selatan", idr: 8, usd: 11, total: 19 },
  { tanggal: "2026-03-07", kantor: "utara", idr: 20, usd: 1, total: 21 },
  { tanggal: "2026-03-07", kantor: "selatan", idr: 5, usd: 3, total: 8 },
  { tanggal: "2026-03-08", kantor: "utara", idr: 18, usd: 0, total: 18 },
  { tanggal: "2026-03-08", kantor: "selatan", idr: 8, usd: 4, total: 12 },
  { tanggal: "2026-03-09", kantor: "utara", idr: 27, usd: 1, total: 28 },
  { tanggal: "2026-03-09", kantor: "selatan", idr: 6, usd: 3, total: 9 },
  { tanggal: "2026-03-10", kantor: "utara", idr: 19, usd: 0, total: 19 },
  { tanggal: "2026-03-10", kantor: "selatan", idr: 9, usd: 2, total: 11 },
  { tanggal: "2026-03-11", kantor: "utara", idr: 37, usd: 0, total: 37 },
  { tanggal: "2026-03-11", kantor: "selatan", idr: 5, usd: 5, total: 10 },
  { tanggal: "2026-03-12", kantor: "utara", idr: 20, usd: 0, total: 20 },
  { tanggal: "2026-03-12", kantor: "selatan", idr: 7, usd: 6, total: 13 },
  { tanggal: "2026-03-13", kantor: "utara", idr: 31, usd: 0, total: 31 },
  { tanggal: "2026-03-13", kantor: "selatan", idr: 11, usd: 6, total: 17 },
  { tanggal: "2026-03-14", kantor: "utara", idr: 35, usd: 0, total: 35 },
  { tanggal: "2026-03-14", kantor: "selatan", idr: 11, usd: 5, total: 16 },
  { tanggal: "2026-03-15", kantor: "utara", idr: 11, usd: 0, total: 11 },
  { tanggal: "2026-03-15", kantor: "selatan", idr: 3, usd: 0, total: 3 },
];

const fallbackDashboardData = manualDashboardData;
const fallbackRecapTablesData = normalizeRecapData(
  Array.isArray(window.excelRecapTablesData) ? window.excelRecapTablesData : []
);
const appConfig = typeof window.APP_CONFIG === "object" && window.APP_CONFIG !== null ? window.APP_CONFIG : {};
const dashboardApiUrl = typeof appConfig.dashboardApiUrl === "string" ? appConfig.dashboardApiUrl.trim() : "";
const dashboardApiToken = typeof appConfig.dashboardApiToken === "string" ? appConfig.dashboardApiToken.trim() : "";
const externalRecapApiUrl =
  typeof appConfig.externalRecapApiUrl === "string" ? appConfig.externalRecapApiUrl.trim() : "";
const externalRecapApiToken =
  typeof appConfig.externalRecapApiToken === "string" ? appConfig.externalRecapApiToken.trim() : "";

let dashboardData = fallbackDashboardData;
let recapTablesData = fallbackRecapTablesData;

let chartTotal;
let chartTotalYearly;
let chartInvoiceMonthly;
let chartDailyStacked;

function padNumber(value) {
  return String(value).padStart(2, "0");
}

function getDaysInMonth(period) {
  const [year, month] = period.split("-").map(Number);
  return new Date(year, month, 0).getDate();
}

function getTodayDateValue() {
  const today = new Date();
  return `${today.getFullYear()}-${padNumber(today.getMonth() + 1)}-${padNumber(today.getDate())}`;
}

function getTodayPeriodValue() {
  return getTodayDateValue().slice(0, 7);
}

function formatRealtimeDateLabel(value) {
  const [year, month, day] = value.split("-").map(Number);
  return new Intl.DateTimeFormat("id-ID", {
    weekday: "long",
    day: "2-digit",
    month: "long",
    year: "numeric",
  }).format(new Date(year, month - 1, day));
}

function formatRealtimeTimeLabel(date = new Date()) {
  return new Intl.DateTimeFormat("id-ID", {
    hour: "2-digit",
    minute: "2-digit",
    second: "2-digit",
  }).format(date);
}

function formatMonthLabel(period) {
  const [year, month] = period.split("-").map(Number);
  return new Intl.DateTimeFormat("id-ID", { month: "long", year: "numeric" }).format(
    new Date(year, month - 1, 1)
  );
}

function formatNumber(value) {
  return new Intl.NumberFormat("id-ID").format(value);
}

function formatAverage(value) {
  return new Intl.NumberFormat("id-ID", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  }).format(value);
}

function formatPercentage(value, total) {
  const percentage = total === 0 ? 0 : (value / total) * 100;
  return `${new Intl.NumberFormat("id-ID", {
    minimumFractionDigits: 1,
    maximumFractionDigits: 1,
  }).format(percentage)}%`;
}

function sumSeriesValues(primary, secondary) {
  const left = Array.isArray(primary) ? primary : [];
  const right = Array.isArray(secondary) ? secondary : [];
  const size = Math.max(left.length, right.length);
  const values = [];

  for (let index = 0; index < size; index += 1) {
    values.push((left[index] || 0) + (right[index] || 0));
  }

  return values;
}

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

function buildDummyData(year) {
  const data = [];
  for (let month = 1; month <= 12; month += 1) {
    const daysInMonth = new Date(year, month, 0).getDate();
    const monthWave = Math.sin((month / 12) * Math.PI * 2);

    for (let day = 1; day <= daysInMonth; day += 1) {
      const dayWave = Math.sin((day / daysInMonth) * Math.PI * 2);
      const utaraTotal = Math.max(8, Math.round(18 + month * 0.7 + monthWave * 2 + dayWave * 5 + (day % 5)));
      const selatanTotal = Math.max(
        6,
        Math.round(15 + month * 0.6 - monthWave * 1.6 + Math.cos((day / daysInMonth) * Math.PI * 2) * 4 + (day % 4))
      );

      const utaraUsd = Math.max(0, Math.round(utaraTotal * (0.2 + month * 0.002)));
      const selatanUsd = Math.max(0, Math.round(selatanTotal * (0.22 + month * 0.002)));

      const tanggal = `${year}-${padNumber(month)}-${padNumber(day)}`;

      data.push({
        tanggal,
        kantor: "utara",
        idr: utaraTotal - utaraUsd,
        usd: utaraUsd,
        total: utaraTotal,
      });
      data.push({
        tanggal,
        kantor: "selatan",
        idr: selatanTotal - selatanUsd,
        usd: selatanUsd,
        total: selatanTotal,
      });
    }
  }
  return data;
}

function parseNumberValue(value) {
  if (typeof value === "number" && Number.isFinite(value)) {
    return value;
  }

  const raw = String(value ?? "").trim();
  if (!raw) {
    return 0;
  }

  // Keep numeric separators, remove currency/text prefixes like "Rp."
  const cleaned = raw.replace(/[^0-9,.-]/g, "").replace(/^[.,]+/, "");
  if (!cleaned) {
    return 0;
  }

  if (/^-?\d{1,3}(\.\d{3})+(,\d+)?$/.test(cleaned)) {
    return Number(cleaned.replace(/\./g, "").replace(",", ".")) || 0;
  }

  if (/^-?\d{1,3}(,\d{3})+(\.\d+)?$/.test(cleaned)) {
    return Number(cleaned.replace(/,/g, "")) || 0;
  }

  if (/^-?\d+(,\d+)?$/.test(cleaned) && cleaned.includes(",")) {
    return Number(cleaned.replace(",", ".")) || 0;
  }

  if (/^-?\d+(\.\d+)?$/.test(cleaned)) {
    return Number(cleaned) || 0;
  }

  // Fallback for mixed formats: keep sign and digits only.
  const sanitized = cleaned.replace(/[^0-9-]/g, "");
  return Number(sanitized) || 0;
}

function normalizeOfficeName(value) {
  const office = String(value ?? "").trim().toLowerCase();

  if (["utara", "north", "u"].includes(office)) {
    return "utara";
  }

  if (["selatan", "south", "s"].includes(office)) {
    return "selatan";
  }

  return "";
}

function normalizeDateValue(value) {
  const raw = String(value ?? "").trim();
  if (!raw) {
    return "";
  }

  if (/^\d{4}-\d{2}-\d{2}/.test(raw)) {
    return raw.slice(0, 10);
  }

  const parsedDate = new Date(raw);
  if (Number.isNaN(parsedDate.getTime())) {
    return "";
  }

  return `${parsedDate.getFullYear()}-${padNumber(parsedDate.getMonth() + 1)}-${padNumber(parsedDate.getDate())}`;
}

function extractDashboardRows(payload) {
  if (Array.isArray(payload)) {
    return payload;
  }

  if (payload && typeof payload === "object") {
    if (Array.isArray(payload.rows)) {
      return payload.rows;
    }
    if (Array.isArray(payload.data)) {
      return payload.data;
    }
    if (Array.isArray(payload.items)) {
      return payload.items;
    }
    if (Array.isArray(payload.results)) {
      return payload.results;
    }
  }

  return [];
}

function normalizeDashboardRows(payload) {
  const rawRows = extractDashboardRows(payload);

  return rawRows
    .map((row) => {
      const tanggal = normalizeDateValue(row?.tanggal ?? row?.date ?? row?.tgl ?? row?.transaction_date);
      const kantor = normalizeOfficeName(row?.kantor ?? row?.office ?? row?.cabang ?? row?.branch);
      const idr = parseNumberValue(row?.idr ?? row?.domestik ?? row?.rupiah);
      const usd = parseNumberValue(row?.usd ?? row?.asing ?? row?.dollar);
      const totalValue = row?.total ?? row?.gerakan ?? row?.movements ?? row?.jumlah;
      const computedTotal = Math.max(0, Math.round(parseNumberValue(totalValue) || idr + usd));

      if (!tanggal || !kantor) {
        return null;
      }

      return {
        tanggal,
        kantor,
        idr: Math.max(0, Math.round(idr)),
        usd: Math.max(0, Math.round(usd)),
        total: computedTotal,
      };
    })
    .filter(Boolean)
    .sort((left, right) => {
      const dateComparison = left.tanggal.localeCompare(right.tanggal);
      if (dateComparison !== 0) {
        return dateComparison;
      }
      return left.kantor.localeCompare(right.kantor);
    });
}

async function loadDashboardData() {
  if (!dashboardApiUrl) {
    return fallbackDashboardData;
  }

  const headers = { Accept: "application/json" };
  if (dashboardApiToken) {
    headers.Authorization = `Bearer ${dashboardApiToken}`;
  }

  try {
    const response = await fetch(dashboardApiUrl, { method: "GET", headers });
    if (!response.ok) {
      throw new Error(`HTTP ${response.status}`);
    }

    const payload = await response.json();
    const normalized = normalizeDashboardRows(payload);
    if (!normalized.length) {
      throw new Error("Format payload dashboard tidak sesuai");
    }

    return normalized;
  } catch (error) {
    console.error("Gagal memuat data dashboard via API, fallback ke data dummy.", error);
    return fallbackDashboardData;
  }
}

function getAvailableMonths() {
  const dashboardMonths = dashboardData.map((row) => row.tanggal.slice(0, 7));
  const recapMonths = recapTablesData
    .map((group) => String(group?.period ?? "").trim())
    .filter((period) => /^\d{4}-\d{2}$/.test(period));
  return Array.from(new Set([...dashboardMonths, ...recapMonths])).sort();
}

function getRowsByDate(dateString) {
  return dashboardData.filter((row) => row.tanggal === dateString);
}

function getRowsByMonth(period) {
  return dashboardData.filter((row) => row.tanggal.startsWith(period));
}

function getRowsByYear(year) {
  return dashboardData.filter((row) => row.tanggal.startsWith(`${year}-`));
}

function aggregateRows(rows) {
  return rows.reduce(
    (acc, row) => {
      acc.idr += row.idr;
      acc.usd += row.usd;
      acc.total += row.total;
      if (row.kantor === "utara") {
        acc.totalUtara += row.total;
        acc.idrUtara += row.idr;
        acc.usdUtara += row.usd;
      }
      if (row.kantor === "selatan") {
        acc.totalSelatan += row.total;
        acc.idrSelatan += row.idr;
        acc.usdSelatan += row.usd;
      }
      return acc;
    },
    {
      idr: 0,
      usd: 0,
      total: 0,
      totalUtara: 0,
      totalSelatan: 0,
      idrUtara: 0,
      idrSelatan: 0,
      usdUtara: 0,
      usdSelatan: 0,
    }
  );
}

function getSeriesByMonth(period) {
  const daysInMonth = getDaysInMonth(period);
  const labels = Array.from({ length: daysInMonth }, (_, index) => index + 1);
  const utaraTotal = [];
  const utaraIdr = [];
  const utaraUsd = [];
  const selatanTotal = [];
  const selatanIdr = [];
  const selatanUsd = [];

  labels.forEach((day) => {
    const tanggal = `${period}-${padNumber(day)}`;
    const rows = getRowsByDate(tanggal);
    const utara = rows.find((row) => row.kantor === "utara") || { total: 0, idr: 0, usd: 0 };
    const selatan = rows.find((row) => row.kantor === "selatan") || { total: 0, idr: 0, usd: 0 };

    utaraTotal.push(utara.total);
    utaraIdr.push(utara.idr);
    utaraUsd.push(utara.usd);
    selatanTotal.push(selatan.total);
    selatanIdr.push(selatan.idr);
    selatanUsd.push(selatan.usd);
  });

  return {
    labels,
    utara: { total: utaraTotal, idr: utaraIdr, usd: utaraUsd },
    selatan: { total: selatanTotal, idr: selatanIdr, usd: selatanUsd },
  };
}

function getOfficeDailyValues(period, office) {
  const daysInMonth = getDaysInMonth(period);
  const idr = [];
  const usd = [];
  const total = [];

  for (let day = 1; day <= daysInMonth; day += 1) {
    const tanggal = `${period}-${padNumber(day)}`;
    const row =
      dashboardData.find((item) => item.tanggal === tanggal && item.kantor === office) || {
        idr: 0,
        usd: 0,
        total: 0,
      };
    idr.push(row.idr);
    usd.push(row.usd);
    total.push(row.total);
  }

  return { idr, usd, total };
}

function getCombinedDailyValues(period) {
  const daysInMonth = getDaysInMonth(period);
  const idr = [];
  const usd = [];
  const total = [];

  for (let day = 1; day <= daysInMonth; day += 1) {
    const tanggal = `${period}-${padNumber(day)}`;
    const dailyTotals = aggregateRows(getRowsByDate(tanggal));
    idr.push(dailyTotals.idr);
    usd.push(dailyTotals.usd);
    total.push(dailyTotals.total);
  }

  return { idr, usd, total };
}

function buildMonthlyRecapRow(label, values, daysInMonth) {
  const sum = values.reduce((total, current) => total + current, 0);
  const avg = sum / daysInMonth;
  const cells = values.map((value) => `<td>${formatNumber(value)}</td>`).join("");
  return `
    <tr>
      <th scope="row">${label}</th>
      ${cells}
      <td>${formatNumber(sum)}</td>
      <td>${formatAverage(avg)}</td>
    </tr>
  `;
}

function buildOfficeMonthlyRecapTable(period, office) {
  const daysInMonth = getDaysInMonth(period);
  const dayHeaders = Array.from({ length: daysInMonth }, (_, index) => `<th>${index + 1}</th>`).join("");
  const values = getOfficeDailyValues(period, office);
  const officeLabel = office.charAt(0).toUpperCase() + office.slice(1);

  return `
    <section class="monthly-recap__block">
      <span class="monthly-recap__label">${officeLabel}</span>
      <div class="monthly-recap__wrap">
        <table class="monthly-recap__table">
          <thead>
            <tr>
              <th>Tanggal</th>
              ${dayHeaders}
              <th>Total</th>
              <th>Rata-Rata</th>
            </tr>
          </thead>
          <tbody>
            ${buildMonthlyRecapRow("IDR", values.idr, daysInMonth)}
            ${buildMonthlyRecapRow("USD", values.usd, daysInMonth)}
            ${buildMonthlyRecapRow("Total", values.total, daysInMonth)}
          </tbody>
        </table>
      </div>
    </section>
  `;
}

function buildCombinedMonthlyRecapTable(period) {
  const daysInMonth = getDaysInMonth(period);
  const dayHeaders = Array.from({ length: daysInMonth }, (_, index) => `<th>${index + 1}</th>`).join("");
  const values = getCombinedDailyValues(period);

  return `
    <section class="monthly-recap__block">
      <span class="monthly-recap__label">Total (Utara + Selatan)</span>
      <div class="monthly-recap__wrap">
        <table class="monthly-recap__table">
          <thead>
            <tr>
              <th>Tanggal</th>
              ${dayHeaders}
              <th>Total</th>
              <th>Rata-Rata</th>
            </tr>
          </thead>
          <tbody>
            ${buildMonthlyRecapRow("IDR", values.idr, daysInMonth)}
            ${buildMonthlyRecapRow("USD", values.usd, daysInMonth)}
            ${buildMonthlyRecapRow("Total", values.total, daysInMonth)}
          </tbody>
        </table>
      </div>
    </section>
  `;
}

function getInvoiceNominalSummaryByPeriod(period) {
  const selectedPeriod = String(period ?? "").trim();
  const item = manualInvoiceTotals.find((entry) => entry.period === selectedPeriod);
  if (!item) {
    return {
      label: selectedPeriod,
      pandu: { domestic: 0, foreign: 0, special: 0 },
      tunda: { domestic: 0, foreign: 0, special: 0 },
      paid: { pandu: 0, tunda: 0, domestic: 0, foreign: 0, special: 0, total: 0 },
      settledTotal: 0,
    };
  }

  const paidPandu = item.paidPandu ?? 0;
  const paidTunda = item.paidTunda ?? 0;
  const paidDomestic = item.paidDomestic ?? 0;
  const paidForeign = item.paidForeign ?? 0;
  const paidSpecial = item.paidSpecial ?? 0;
  const paidTotal = item.settledTotal ?? paidPandu + paidTunda;

  return {
    label: item.label || selectedPeriod,
    pandu: {
      domestic: item.domesticPandu || 0,
      foreign: item.foreignPandu || 0,
      special: item.specialPandu || 0,
    },
    tunda: {
      domestic: item.domesticTunda || 0,
      foreign: item.foreignTunda || 0,
      special: item.specialTunda || 0,
    },
    paid: {
      pandu: paidPandu,
      tunda: paidTunda,
      domestic: paidDomestic,
      foreign: paidForeign,
      special: paidSpecial,
      total: paidTotal || 0,
    },
    settledTotal: paidTotal || 0,
  };
}

function buildMonthlyNominalNotes(period) {
  const summary = getInvoiceNominalSummaryByPeriod(period);
  const totalPandu = summary.pandu.domestic + summary.pandu.foreign + summary.pandu.special;
  const totalTunda = summary.tunda.domestic + summary.tunda.foreign + summary.tunda.special;
  const totalDomestic = summary.pandu.domestic + summary.tunda.domestic;
  const totalForeign = summary.pandu.foreign + summary.tunda.foreign;
  const totalSpecial = summary.pandu.special + summary.tunda.special;
  const paidPandu = summary.paid?.pandu || 0;
  const paidTunda = summary.paid?.tunda || 0;
  const paidDomestic = summary.paid?.domestic || 0;
  const paidForeign = summary.paid?.foreign || 0;
  const paidSpecial = summary.paid?.special || 0;
  const settledTotal = summary.paid?.total || summary.settledTotal || 0;
  const monthlyInvoiceTotal = totalPandu + totalTunda;

  return `
    <section class="monthly-recap__notes">
      <div class="monthly-recap__summary-wrap">
        <div class="monthly-recap__summary-grid">
          <p class="monthly-recap__summary-cell monthly-recap__summary-cell--pair monthly-recap__summary-cell--top">
            <span class="monthly-recap__summary-label">Total Invoice Terbit ${escapeHtml(
              summary.label || "Bulan"
            )}</span>
            <strong class="monthly-recap__summary-value">${escapeHtml(formatNumber(monthlyInvoiceTotal))}</strong>
          </p>
          <p class="monthly-recap__summary-cell monthly-recap__summary-cell--right monthly-recap__summary-cell--pair monthly-recap__summary-cell--top">
            <span class="monthly-recap__summary-label">Total Invoice Terbayar ${escapeHtml(
              summary.label || "Bulan"
            )}</span>
            <strong class="monthly-recap__summary-value">${escapeHtml(formatNumber(settledTotal))}</strong>
          </p>
          <p class="monthly-recap__summary-cell monthly-recap__summary-cell--pair">
            <span class="monthly-recap__summary-label">Nominal Pandu</span>
            <strong class="monthly-recap__summary-value">${escapeHtml(formatNumber(totalPandu))}</strong>
          </p>
          <p class="monthly-recap__summary-cell monthly-recap__summary-cell--right monthly-recap__summary-cell--pair">
            <span class="monthly-recap__summary-label">Nominal Pandu Terbayar</span>
            <strong class="monthly-recap__summary-value">${escapeHtml(formatNumber(paidPandu))}</strong>
          </p>
          <p class="monthly-recap__summary-cell monthly-recap__summary-cell--pair">
            <span class="monthly-recap__summary-label">Nominal Tunda</span>
            <strong class="monthly-recap__summary-value">${escapeHtml(formatNumber(totalTunda))}</strong>
          </p>
          <p class="monthly-recap__summary-cell monthly-recap__summary-cell--right monthly-recap__summary-cell--pair">
            <span class="monthly-recap__summary-label">Nominal Tunda Terbayar</span>
            <strong class="monthly-recap__summary-value">${escapeHtml(formatNumber(paidTunda))}</strong>
          </p>
          <p class="monthly-recap__summary-cell monthly-recap__summary-cell--pair">
            <span class="monthly-recap__summary-label">Total Domestik</span>
            <strong class="monthly-recap__summary-value">${escapeHtml(formatNumber(totalDomestic))}</strong>
          </p>
          <p class="monthly-recap__summary-cell monthly-recap__summary-cell--right monthly-recap__summary-cell--pair">
            <span class="monthly-recap__summary-label">Total Domestik Terbayar</span>
            <strong class="monthly-recap__summary-value">${escapeHtml(formatNumber(paidDomestic))}</strong>
          </p>
          <p class="monthly-recap__summary-cell monthly-recap__summary-cell--pair">
            <span class="monthly-recap__summary-label">Total Asing</span>
            <strong class="monthly-recap__summary-value">${escapeHtml(formatNumber(totalForeign))}</strong>
          </p>
          <p class="monthly-recap__summary-cell monthly-recap__summary-cell--right monthly-recap__summary-cell--pair">
            <span class="monthly-recap__summary-label">Total Asing Terbayar</span>
            <strong class="monthly-recap__summary-value">${escapeHtml(formatNumber(paidForeign))}</strong>
          </p>
          <p class="monthly-recap__summary-cell monthly-recap__summary-cell--pair">
            <span class="monthly-recap__summary-label">Total Khusus</span>
            <strong class="monthly-recap__summary-value">${escapeHtml(formatNumber(totalSpecial))}</strong>
          </p>
          <p class="monthly-recap__summary-cell monthly-recap__summary-cell--right monthly-recap__summary-cell--pair">
            <span class="monthly-recap__summary-label">Total Khusus Terbayar</span>
            <strong class="monthly-recap__summary-value">${escapeHtml(formatNumber(paidSpecial))}</strong>
          </p>
        </div>
      </div>
    </section>
  `;
}

function renderMonthlyRecapTable(period) {
  monthlyRecapTable.innerHTML = `
    ${buildCombinedMonthlyRecapTable(period)}
    ${buildOfficeMonthlyRecapTable(period, "selatan")}
    ${buildOfficeMonthlyRecapTable(period, "utara")}
    ${buildMonthlyNominalNotes(period)}
  `;
}

function normalizeRowCells(cells, headerLength) {
  const normalized = Array.isArray(cells) ? cells.map((cell) => String(cell ?? "").trim()) : [];
  while (normalized.length < headerLength) {
    normalized.push("");
  }
  return normalized.slice(0, headerLength);
}

function normalizeRecapSectionsArray(rawSections) {
  return rawSections
    .map((section) => {
      const title = String(section?.title ?? "").trim();
      const headers = Array.isArray(section?.headers)
        ? section.headers.map((item) => String(item ?? "").trim()).filter(Boolean)
        : [];
      const rows = Array.isArray(section?.rows)
        ? section.rows.map((row) => normalizeRowCells(row, headers.length))
        : [];

      if (!headers.length || !rows.length) {
        return null;
      }

      return { title, headers, rows };
    })
    .filter(Boolean);
}

function extractRecapPeriodGroups(payload) {
  if (Array.isArray(payload)) {
    const looksLikePeriodGroups = payload.every(
      (item) => item && typeof item === "object" && Array.isArray(item.sections)
    );
    return looksLikePeriodGroups ? payload : [{ period: "", sections: payload }];
  }

  if (payload && typeof payload === "object") {
    if (Array.isArray(payload.periods)) {
      return payload.periods;
    }
    if (Array.isArray(payload.sections)) {
      return [{ period: payload.period ?? "", sections: payload.sections }];
    }
    if (Array.isArray(payload.data)) {
      const looksLikePeriodGroups = payload.data.every(
        (item) => item && typeof item === "object" && Array.isArray(item.sections)
      );
      return looksLikePeriodGroups ? payload.data : [{ period: payload.period ?? "", sections: payload.data }];
    }
  }

  return [];
}

function normalizeRecapPeriodValue(value) {
  const raw = String(value ?? "").trim();
  return /^\d{4}-\d{2}$/.test(raw) ? raw : "";
}

function formatPeriodMonthName(period) {
  const normalized = normalizeRecapPeriodValue(period);
  if (!normalized) {
    return "";
  }

  const [year, month] = normalized.split("-").map(Number);
  return new Intl.DateTimeFormat("id-ID", { month: "long" }).format(new Date(year, month - 1, 1));
}

function normalizeRecapData(payload) {
  const rawGroups = extractRecapPeriodGroups(payload);
  return rawGroups
    .map((group) => {
      const sections = normalizeRecapSectionsArray(Array.isArray(group?.sections) ? group.sections : []);
      if (!sections.length) {
        return null;
      }
      return {
        period: normalizeRecapPeriodValue(group?.period),
        sections,
      };
    })
    .filter(Boolean);
}

function buildExternalRecapBlock(section) {
  const headers = Array.isArray(section?.headers) ? section.headers.map((item) => String(item ?? "").trim()) : [];
  const rows = Array.isArray(section?.rows) ? section.rows : [];
  const title = String(section?.title ?? "").trim();

  if (!headers.length || !rows.length) {
    return "";
  }

  const headerHtml = headers.map((header) => `<th>${escapeHtml(header)}</th>`).join("");
  const bodyHtml = rows
    .map((row) => {
      const cells = normalizeRowCells(row, headers.length).map((cell) => escapeHtml(cell));
      const firstCell = `<th scope="row">${cells[0] || "-"}</th>`;
      const restCells = cells
        .slice(1)
        .map((cell) => `<td>${cell || "-"}</td>`)
        .join("");
      return `<tr>${firstCell}${restCells}</tr>`;
    })
    .join("");

  return `
    <section class="external-recap__block">
      <span class="external-recap__title">${escapeHtml(title || "Tabel")}</span>
      <div class="monthly-recap__wrap">
        <table class="monthly-recap__table">
          <thead>
            <tr>${headerHtml}</tr>
          </thead>
          <tbody>
            ${bodyHtml}
          </tbody>
        </table>
      </div>
    </section>
  `;
}

function isAgentRecapSection(section, keyword) {
  const headers = Array.isArray(section?.headers) ? section.headers : [];
  const title = String(section?.title ?? "").toLowerCase();
  return headers.length >= 4 && String(headers[1] ?? "").trim().toLowerCase() === "agen" && title.includes(keyword);
}

function isJettyRecapSection(section, keyword) {
  const headers = Array.isArray(section?.headers) ? section.headers : [];
  const title = String(section?.title ?? "").toLowerCase();
  return headers.length >= 4 && String(headers[1] ?? "").trim().toLowerCase() === "jetty" && title.includes(keyword);
}

function normalizeRevenueDisplay(value) {
  const raw = String(value ?? "").trim();
  if (!raw || raw === "0" || raw.toLowerCase() === "rp. 0") {
    return "Rp. 0";
  }
  if (raw.toLowerCase().startsWith("rp")) {
    return raw;
  }
  const numeric = parseNumberValue(raw);
  return numeric > 0 ? `Rp. ${formatNumber(numeric)}` : "Rp. 0";
}

function formatRevenueValue(value) {
  const numeric = Math.max(0, Math.round(parseNumberValue(value)));
  return numeric > 0 ? `Rp. ${formatNumber(numeric)}` : "Rp. 0";
}

function normalizeMovementHeaderLabel(header, periodKey = "") {
  const raw = String(header ?? "").trim();
  const stripped = raw.replace(/^gerakan\s*-\s*/i, "").trim();
  if (stripped) {
    return stripped;
  }
  const monthName = formatPeriodMonthName(periodKey);
  return monthName || raw;
}

function normalizeRevenueHeaderLabel(header, periodKey = "") {
  const raw = String(header ?? "").trim();
  const monthName = formatPeriodMonthName(periodKey);

  if (!raw) {
    return monthName ? `Invoice Bulanan - ${monthName}` : "Invoice Bulanan";
  }

  if (/^invoice/i.test(raw)) {
    return raw;
  }

  const stripped = raw.replace(/^est\s*pendapatan\s*-\s*/i, "").trim();
  if (stripped) {
    return `Invoice Bulanan - ${stripped}`;
  }

  return raw;
}

function extractAgentSectionData(section) {
  const rows = Array.isArray(section?.rows) ? section.rows : [];
  const byAgent = new Map();
  let computedMovement = 0;
  let computedRevenue = 0;
  let totalMovementFromRow = null;
  let totalRevenueFromRow = null;

  rows.map((row) => normalizeRowCells(row, 4)).forEach((cells) => {
    const firstCell = String(cells[0] ?? "").trim().toUpperCase();
    if (firstCell === "TOTAL") {
      totalMovementFromRow = Math.max(0, Math.round(parseNumberValue(cells[2])));
      totalRevenueFromRow = Math.max(0, Math.round(parseNumberValue(cells[3])));
      return;
    }

    const agentName = String(cells[1] ?? "").trim();
    if (!agentName) {
      return;
    }

    const movement = Math.max(0, Math.round(parseNumberValue(cells[2])));
    const revenue = Math.max(0, Math.round(parseNumberValue(cells[3])));
    const key = agentName.toUpperCase();
    if (!byAgent.has(key)) {
      byAgent.set(key, { agent: agentName, movement: 0, revenue: 0 });
    }

    const target = byAgent.get(key);
    target.movement += movement;
    target.revenue += revenue;

    computedMovement += movement;
    computedRevenue += revenue;
  });

  return {
    byAgent,
    movementHeader: normalizeMovementHeaderLabel(section?.headers?.[2]),
    revenueHeader: String(section?.headers?.[3] ?? "").trim(),
    totalMovement: totalMovementFromRow ?? computedMovement,
    totalRevenue: totalRevenueFromRow ?? computedRevenue,
  };
}

function buildCombinedAgentRecapBlock(foreignEntries, domesticEntries, period) {
  const periodMap = new Map();
  const ensurePeriodBucket = (periodKey) => {
    if (!periodMap.has(periodKey)) {
      periodMap.set(periodKey, {
        period: periodKey,
        movementHeader: "",
        revenueHeader: "",
        foreign: null,
        domestic: null,
      });
    }
    return periodMap.get(periodKey);
  };

  foreignEntries.forEach((entry) => {
    const normalizedPeriod = normalizeRecapPeriodValue(entry?.period);
    if (!normalizedPeriod || !entry?.section) {
      return;
    }

    const bucket = ensurePeriodBucket(normalizedPeriod);
    const sectionData = extractAgentSectionData(entry.section);
    bucket.foreign = sectionData;
    if (!bucket.movementHeader) {
      bucket.movementHeader = sectionData.movementHeader;
    }
    if (!bucket.revenueHeader) {
      bucket.revenueHeader = sectionData.revenueHeader;
    }
  });

  domesticEntries.forEach((entry) => {
    const normalizedPeriod = normalizeRecapPeriodValue(entry?.period);
    if (!normalizedPeriod || !entry?.section) {
      return;
    }

    const bucket = ensurePeriodBucket(normalizedPeriod);
    const sectionData = extractAgentSectionData(entry.section);
    bucket.domestic = sectionData;
    if (!bucket.movementHeader) {
      bucket.movementHeader = sectionData.movementHeader;
    }
    if (!bucket.revenueHeader) {
      bucket.revenueHeader = sectionData.revenueHeader;
    }
  });

  const periodKeys = Array.from(periodMap.keys()).sort();
  if (!periodKeys.length) {
    return "";
  }

  const byAgent = new Map();
  const upsertAgent = (agent) => {
    const key = String(agent ?? "").trim().toUpperCase();
    if (!byAgent.has(key)) {
      byAgent.set(key, {
        agent: String(agent ?? "").trim(),
        foreignByPeriod: {},
        domesticByPeriod: {},
      });
    }
    return byAgent.get(key);
  };

  periodKeys.forEach((periodKey) => {
    const bucket = periodMap.get(periodKey);
    if (!bucket) {
      return;
    }

    bucket.foreign?.byAgent.forEach((item) => {
      const agent = upsertAgent(item.agent);
      agent.foreignByPeriod[periodKey] = {
        movement: item.movement,
        revenue: item.revenue,
      };
    });

    bucket.domestic?.byAgent.forEach((item) => {
      const agent = upsertAgent(item.agent);
      agent.domesticByPeriod[periodKey] = {
        movement: item.movement,
        revenue: item.revenue,
      };
    });
  });

  const activePeriod = normalizeRecapPeriodValue(period) || periodKeys[periodKeys.length - 1];
  const yearLabel = /^\d{4}-\d{2}$/.test(activePeriod) ? activePeriod.slice(0, 4) : "";
  const title = `Rekapitulasi Pelayanan Perusahaan Keagenan Kapal ${yearLabel}`.trim();

  const dynamicHeaders = periodKeys
    .map((periodKey) => {
      const bucket = periodMap.get(periodKey);
      const monthName = formatPeriodMonthName(periodKey);
      const movementHeader = normalizeMovementHeaderLabel(bucket?.movementHeader, periodKey) || monthName;
      const revenueHeader = normalizeRevenueHeaderLabel(bucket?.revenueHeader, periodKey);
      return `
        <th>${escapeHtml(movementHeader)}</th>
        <th>${escapeHtml(revenueHeader)}</th>
      `;
    })
    .join("");

  const bodyRows = [];
  Array.from(byAgent.values()).forEach((item, index) => {
    const rowGroupKey = `agent-${index + 1}`;
    const foreignCells = periodKeys
      .map((periodKey) => {
        const monthValue = item.foreignByPeriod[periodKey] || { movement: 0, revenue: 0 };
        return `
          <td>${escapeHtml(formatNumber(monthValue.movement))}</td>
          <td>${escapeHtml(formatRevenueValue(monthValue.revenue))}</td>
        `;
      })
      .join("");
    const domesticCells = periodKeys
      .map((periodKey) => {
        const monthValue = item.domesticByPeriod[periodKey] || { movement: 0, revenue: 0 };
        return `
          <td>${escapeHtml(formatNumber(monthValue.movement))}</td>
          <td>${escapeHtml(formatRevenueValue(monthValue.revenue))}</td>
        `;
      })
      .join("");

    bodyRows.push(`
      <tr class="external-recap__combined-row external-recap__combined-row--ln" data-row-group="${rowGroupKey}">
        <th scope="rowgroup" rowspan="2">${index + 1}</th>
        <th scope="rowgroup" rowspan="2" class="external-recap__agent">${escapeHtml(item.agent)}</th>
        <td class="external-recap__currency">$</td>
        ${foreignCells}
      </tr>
      <tr class="external-recap__combined-row external-recap__combined-row--dn" data-row-group="${rowGroupKey}">
        <td class="external-recap__currency">Rp</td>
        ${domesticCells}
      </tr>
    `);
  });

  const foreignTotalCells = periodKeys
    .map((periodKey) => {
      const bucket = periodMap.get(periodKey);
      const movement = bucket?.foreign?.totalMovement ?? 0;
      const revenue = bucket?.foreign?.totalRevenue ?? 0;
      return `
        <td>${escapeHtml(formatNumber(movement))}</td>
        <td>${escapeHtml(formatRevenueValue(revenue))}</td>
      `;
    })
    .join("");
  const domesticTotalCells = periodKeys
    .map((periodKey) => {
      const bucket = periodMap.get(periodKey);
      const movement = bucket?.domestic?.totalMovement ?? 0;
      const revenue = bucket?.domestic?.totalRevenue ?? 0;
      return `
        <td>${escapeHtml(formatNumber(movement))}</td>
        <td>${escapeHtml(formatRevenueValue(revenue))}</td>
      `;
    })
    .join("");

  if (SHOW_COMBINED_RECAP_TOTAL_ROWS) {
    bodyRows.push(`
      <tr class="external-recap__combined-row external-recap__combined-row--ln external-recap__combined-row--total">
        <th scope="rowgroup" rowspan="2" colspan="2">TOTAL</th>
        <td class="external-recap__currency">$</td>
        ${foreignTotalCells}
      </tr>
      <tr class="external-recap__combined-row external-recap__combined-row--dn external-recap__combined-row--total">
        <td class="external-recap__currency">Rp</td>
        ${domesticTotalCells}
      </tr>
    `);
  }

  return `
    <section class="external-recap__block">
      <span class="external-recap__title">${escapeHtml(title)}</span>
      <div class="monthly-recap__wrap">
        <table class="monthly-recap__table external-recap__combined-table">
          <thead>
            <tr>
              <th>No</th>
              <th>Agen</th>
              <th>Voyage</th>
              ${dynamicHeaders}
            </tr>
          </thead>
          <tbody>
            ${bodyRows.join("")}
          </tbody>
        </table>
      </div>
    </section>
  `;
}

function extractJettySectionData(section) {
  const rows = Array.isArray(section?.rows) ? section.rows : [];
  const byJetty = new Map();
  let computedTotal = 0;
  let totalFromRow = null;

  rows.map((row) => normalizeRowCells(row, 4)).forEach((cells) => {
    const firstCell = String(cells[0] ?? "").trim().toUpperCase();
    if (firstCell === "TOTAL") {
      totalFromRow = Math.max(0, Math.round(parseNumberValue(cells[3])));
      return;
    }

    const jetty = String(cells[1] ?? "").trim();
    if (!jetty) {
      return;
    }

    const region = String(cells[2] ?? "").trim() || "N/A";
    const movement = Math.max(0, Math.round(parseNumberValue(cells[3])));
    const key = `${jetty.toUpperCase()}__${region.toUpperCase()}`;

    if (!byJetty.has(key)) {
      byJetty.set(key, { jetty, region, movement: 0 });
    }
    byJetty.get(key).movement += movement;
    computedTotal += movement;
  });

  return {
    byJetty,
    movementHeader: String(section?.headers?.[3] ?? "").trim(),
    totalMovement: totalFromRow ?? computedTotal,
  };
}

function buildCombinedJettyRecapBlock(foreignEntries, domesticEntries, period) {
  const periodMap = new Map();
  const ensurePeriodBucket = (periodKey) => {
    if (!periodMap.has(periodKey)) {
      periodMap.set(periodKey, {
        period: periodKey,
        movementHeader: "",
        foreign: null,
        domestic: null,
      });
    }
    return periodMap.get(periodKey);
  };

  foreignEntries.forEach((entry) => {
    const normalizedPeriod = normalizeRecapPeriodValue(entry?.period);
    if (!normalizedPeriod || !entry?.section) {
      return;
    }

    const bucket = ensurePeriodBucket(normalizedPeriod);
    const sectionData = extractJettySectionData(entry.section);
    bucket.foreign = sectionData;
    if (!bucket.movementHeader) {
      bucket.movementHeader = sectionData.movementHeader;
    }
  });

  domesticEntries.forEach((entry) => {
    const normalizedPeriod = normalizeRecapPeriodValue(entry?.period);
    if (!normalizedPeriod || !entry?.section) {
      return;
    }

    const bucket = ensurePeriodBucket(normalizedPeriod);
    const sectionData = extractJettySectionData(entry.section);
    bucket.domestic = sectionData;
    if (!bucket.movementHeader) {
      bucket.movementHeader = sectionData.movementHeader;
    }
  });

  const periodKeys = Array.from(periodMap.keys()).sort();
  if (!periodKeys.length) {
    return "";
  }

  const byJetty = new Map();
  const upsertJetty = (jetty, region) => {
    const key = `${String(jetty ?? "").trim().toUpperCase()}__${String(region ?? "").trim().toUpperCase()}`;
    if (!byJetty.has(key)) {
      byJetty.set(key, {
        jetty: String(jetty ?? "").trim(),
        region: String(region ?? "").trim() || "N/A",
        foreignByPeriod: {},
        domesticByPeriod: {},
      });
    }
    return byJetty.get(key);
  };

  periodKeys.forEach((periodKey) => {
    const bucket = periodMap.get(periodKey);
    if (!bucket) {
      return;
    }

    bucket.foreign?.byJetty.forEach((item) => {
      const jetty = upsertJetty(item.jetty, item.region);
      jetty.foreignByPeriod[periodKey] = item.movement;
    });

    bucket.domestic?.byJetty.forEach((item) => {
      const jetty = upsertJetty(item.jetty, item.region);
      jetty.domesticByPeriod[periodKey] = item.movement;
    });
  });

  const activePeriod = normalizeRecapPeriodValue(period) || periodKeys[periodKeys.length - 1];
  const yearLabel = /^\d{4}-\d{2}$/.test(activePeriod) ? activePeriod.slice(0, 4) : "";
  const title = `Rekapitulasi Pelayanan TUKS ${yearLabel}`.trim();

  const dynamicHeaders = periodKeys
    .map((periodKey) => {
      const bucket = periodMap.get(periodKey);
      const monthHeader = bucket?.movementHeader || formatPeriodMonthName(periodKey);
      return `<th>${escapeHtml(monthHeader)}</th>`;
    })
    .join("");

  const bodyRows = [];
  Array.from(byJetty.values()).forEach((item, index) => {
    const rowGroupKey = `jetty-${index + 1}`;
    const foreignCells = periodKeys
      .map((periodKey) => `<td>${escapeHtml(formatNumber(item.foreignByPeriod[periodKey] || 0))}</td>`)
      .join("");
    const domesticCells = periodKeys
      .map((periodKey) => `<td>${escapeHtml(formatNumber(item.domesticByPeriod[periodKey] || 0))}</td>`)
      .join("");

    bodyRows.push(`
      <tr class="external-recap__combined-row external-recap__combined-row--ln" data-row-group="${rowGroupKey}">
        <th scope="rowgroup" rowspan="2">${index + 1}</th>
        <th scope="rowgroup" rowspan="2" class="external-recap__jetty">${escapeHtml(item.jetty)}</th>
        <th scope="rowgroup" rowspan="2" class="external-recap__region">${escapeHtml(item.region)}</th>
        <td class="external-recap__currency">$</td>
        ${foreignCells}
      </tr>
      <tr class="external-recap__combined-row external-recap__combined-row--dn" data-row-group="${rowGroupKey}">
        <td class="external-recap__currency">Rp</td>
        ${domesticCells}
      </tr>
    `);
  });

  const foreignTotalCells = periodKeys
    .map((periodKey) => {
      const bucket = periodMap.get(periodKey);
      const total = bucket?.foreign?.totalMovement ?? 0;
      return `<td>${escapeHtml(formatNumber(total))}</td>`;
    })
    .join("");
  const domesticTotalCells = periodKeys
    .map((periodKey) => {
      const bucket = periodMap.get(periodKey);
      const total = bucket?.domestic?.totalMovement ?? 0;
      return `<td>${escapeHtml(formatNumber(total))}</td>`;
    })
    .join("");

  if (SHOW_COMBINED_RECAP_TOTAL_ROWS) {
    bodyRows.push(`
      <tr class="external-recap__combined-row external-recap__combined-row--ln external-recap__combined-row--total">
        <th scope="rowgroup" rowspan="2" colspan="3">TOTAL</th>
        <td class="external-recap__currency">$</td>
        ${foreignTotalCells}
      </tr>
      <tr class="external-recap__combined-row external-recap__combined-row--dn external-recap__combined-row--total">
        <td class="external-recap__currency">Rp</td>
        ${domesticTotalCells}
      </tr>
    `);
  }

  return `
    <section class="external-recap__block">
      <span class="external-recap__title">${escapeHtml(title)}</span>
      <div class="monthly-recap__wrap">
        <table class="monthly-recap__table external-recap__combined-table external-recap__combined-jetty-table">
          <thead>
            <tr>
              <th>No</th>
              <th>Jetty</th>
              <th>Wilayah</th>
              <th>Voyage</th>
              ${dynamicHeaders}
            </tr>
          </thead>
          <tbody>
            ${bodyRows.join("")}
          </tbody>
        </table>
      </div>
    </section>
  `;
}

function renderInvoiceTotalsWidget() {
  if (!invoiceTotalValue || !invoiceTotalsStats) {
    return;
  }

  const domesticTotal = manualInvoiceTotals.reduce(
    (sum, item) => sum + (item.domesticPandu || 0) + (item.domesticTunda || 0),
    0
  );
  const foreignTotal = manualInvoiceTotals.reduce(
    (sum, item) => sum + (item.foreignPandu || 0) + (item.foreignTunda || 0),
    0
  );
  const specialTotal = manualInvoiceTotals.reduce(
    (sum, item) => sum + (item.specialPandu || 0) + (item.specialTunda || 0),
    0
  );
  const total = domesticTotal + foreignTotal + specialTotal;

  invoiceTotalValue.textContent = formatNumber(total);
  const summaryRows = [
    `
      <div>
        <dt>Domestik</dt>
        <dd>${escapeHtml(formatNumber(domesticTotal))}</dd>
      </div>
    `,
    `
      <div>
        <dt>Asing</dt>
        <dd>${escapeHtml(formatNumber(foreignTotal))}</dd>
      </div>
    `,
  ];

  if (specialTotal > 0) {
    summaryRows.push(
      `
        <div>
          <dt>Khusus</dt>
          <dd>${escapeHtml(formatNumber(specialTotal))}</dd>
        </div>
      `
    );
  }

  invoiceTotalsStats.innerHTML = summaryRows.join("");
}

function renderExternalRecapLoading() {
  if (!externalRecapTables) {
    return;
  }
  externalRecapTables.innerHTML = `<p class="external-recap__empty">Memuat data lampiran...</p>`;
}

function getRecapSectionsForPeriod(period) {
  if (!recapTablesData.length) {
    return [];
  }

  const selectedPeriod = String(period ?? "").trim();
  const exactMatch = recapTablesData.find((group) => group.period === selectedPeriod);
  if (exactMatch) {
    return exactMatch.sections;
  }

  const fallbackGroup = recapTablesData.find((group) => !group.period);
  if (fallbackGroup) {
    return fallbackGroup.sections;
  }

  return recapTablesData[0]?.sections || [];
}

function getRecapGroupsUntilPeriod(period) {
  if (!recapTablesData.length) {
    return [];
  }

  const selectedPeriod = normalizeRecapPeriodValue(period);
  if (selectedPeriod) {
    const selectedYear = selectedPeriod.slice(0, 4);
    const withinYear = recapTablesData
      .filter((group) => normalizeRecapPeriodValue(group?.period).startsWith(`${selectedYear}-`))
      .filter((group) => normalizeRecapPeriodValue(group?.period) <= selectedPeriod)
      .sort((left, right) => String(left.period).localeCompare(String(right.period)));

    if (withinYear.length) {
      return withinYear;
    }

    const exactMatch = recapTablesData.find((group) => group.period === selectedPeriod);
    if (exactMatch) {
      return [exactMatch];
    }
  }

  const fallbackGroup = recapTablesData.find((group) => !normalizeRecapPeriodValue(group?.period));
  if (fallbackGroup) {
    return [fallbackGroup];
  }

  return recapTablesData.slice(0, 1);
}

function collectRecapSectionEntries(groups, matcher) {
  return groups
    .map((group) => {
      const period = normalizeRecapPeriodValue(group?.period);
      if (!period || !Array.isArray(group?.sections)) {
        return null;
      }

      const section = group.sections.find((item) => matcher(item));
      if (!section) {
        return null;
      }

      return { period, section };
    })
    .filter(Boolean);
}

function renderExternalRecapTables(period) {
  if (!externalRecapTables) {
    return;
  }

  const sections = getRecapSectionsForPeriod(period);
  const cumulativeGroups = getRecapGroupsUntilPeriod(period);

  if (!sections.length) {
    externalRecapTables.innerHTML = `<p class="external-recap__empty">Data lampiran belum tersedia.</p>`;
    return;
  }

  const insertionBlocks = new Map();
  const skipIndexes = new Set();

  const pushInsertionBlock = (index, block) => {
    if (!insertionBlocks.has(index)) {
      insertionBlocks.set(index, []);
    }
    insertionBlocks.get(index).push(block);
  };

  const agentForeignIndex = sections.findIndex((section) => isAgentRecapSection(section, "luar negeri"));
  const agentDomesticIndex = sections.findIndex((section) => isAgentRecapSection(section, "dalam negeri"));
  if (agentForeignIndex !== -1 && agentDomesticIndex !== -1) {
    const foreignEntries = collectRecapSectionEntries(cumulativeGroups, (section) =>
      isAgentRecapSection(section, "luar negeri")
    );
    const domesticEntries = collectRecapSectionEntries(cumulativeGroups, (section) =>
      isAgentRecapSection(section, "dalam negeri")
    );
    const block = buildCombinedAgentRecapBlock(foreignEntries, domesticEntries, period);
    if (block) {
      pushInsertionBlock(Math.min(agentForeignIndex, agentDomesticIndex), block);
      skipIndexes.add(agentForeignIndex);
      skipIndexes.add(agentDomesticIndex);
    }
  }

  const jettyForeignIndex = sections.findIndex((section) => isJettyRecapSection(section, "luar negeri"));
  const jettyDomesticIndex = sections.findIndex((section) => isJettyRecapSection(section, "dalam negeri"));
  if (jettyForeignIndex !== -1 && jettyDomesticIndex !== -1) {
    const foreignEntries = collectRecapSectionEntries(cumulativeGroups, (section) =>
      isJettyRecapSection(section, "luar negeri")
    );
    const domesticEntries = collectRecapSectionEntries(cumulativeGroups, (section) =>
      isJettyRecapSection(section, "dalam negeri")
    );
    const block = buildCombinedJettyRecapBlock(foreignEntries, domesticEntries, period);
    if (block) {
      pushInsertionBlock(Math.min(jettyForeignIndex, jettyDomesticIndex), block);
      skipIndexes.add(jettyForeignIndex);
      skipIndexes.add(jettyDomesticIndex);
    }
  }

  const blocks = [];
  sections.forEach((section, index) => {
    if (insertionBlocks.has(index)) {
      blocks.push(...insertionBlocks.get(index));
    }
    if (skipIndexes.has(index)) {
      return;
    }
    blocks.push(buildExternalRecapBlock(section));
  });

  const html = blocks.join("");
  externalRecapTables.innerHTML =
    html || `<p class="external-recap__empty">Data lampiran belum tersedia.</p>`;
}

function setupCombinedRecapRowToggle() {
  if (!externalRecapTables) {
    return;
  }

  externalRecapTables.addEventListener("click", (event) => {
    const clickedRow = event.target.closest("tr.external-recap__combined-row[data-row-group]");
    if (!clickedRow || !externalRecapTables.contains(clickedRow)) {
      return;
    }

    const table = clickedRow.closest("table.external-recap__combined-table");
    const rowGroup = clickedRow.getAttribute("data-row-group");
    if (!table || !rowGroup) {
      return;
    }

    const shouldSelect = !clickedRow.classList.contains("external-recap__combined-row--selected");
    table.querySelectorAll(`tr.external-recap__combined-row[data-row-group="${rowGroup}"]`).forEach((row) => {
      row.classList.toggle("external-recap__combined-row--selected", shouldSelect);
    });
  });
}

async function loadRecapTablesData() {
  if (!externalRecapApiUrl) {
    return fallbackRecapTablesData;
  }

  const headers = { Accept: "application/json" };
  if (externalRecapApiToken) {
    headers.Authorization = `Bearer ${externalRecapApiToken}`;
  }

  try {
    const response = await fetch(externalRecapApiUrl, { method: "GET", headers });
    if (!response.ok) {
      throw new Error(`HTTP ${response.status}`);
    }

    const payload = await response.json();
    const normalized = normalizeRecapData(payload);
    if (!normalized.length) {
      throw new Error("Format payload tidak sesuai");
    }

    return normalized;
  } catch (error) {
    console.error("Gagal memuat data lampiran via API, fallback ke file lokal.", error);
    return fallbackRecapTablesData;
  }
}

async function initRecapTablesData() {
  renderExternalRecapLoading();
  recapTablesData = await loadRecapTablesData();
}

function getSeriesByYear(year) {
  const labels = [];
  const utaraTotal = [];
  const selatanTotal = [];
  const utaraIdr = [];
  const selatanIdr = [];
  const utaraUsd = [];
  const selatanUsd = [];

  for (let month = 1; month <= 12; month += 1) {
    const period = `${year}-${padNumber(month)}`;
    const totals = aggregateRows(getRowsByMonth(period));
    labels.push(
      new Intl.DateTimeFormat("id-ID", { month: "short" }).format(new Date(Number(year), month - 1, 1))
    );
    utaraTotal.push(totals.totalUtara);
    selatanTotal.push(totals.totalSelatan);
    utaraIdr.push(totals.idrUtara);
    selatanIdr.push(totals.idrSelatan);
    utaraUsd.push(totals.usdUtara);
    selatanUsd.push(totals.usdSelatan);
  }

  return { labels, utaraTotal, selatanTotal, utaraIdr, selatanIdr, utaraUsd, selatanUsd };
}

function getInvoiceNominalSeriesByYear(year) {
  const selectedYear = String(year ?? "").trim();
  const entries = manualInvoiceTotals
    .filter((entry) => String(entry?.period ?? "").startsWith(`${selectedYear}-`))
    .sort((left, right) => String(left.period).localeCompare(String(right.period)));

  const labels = entries.map((entry) => {
    const [entryYear, entryMonth] = String(entry.period).split("-").map(Number);
    return new Intl.DateTimeFormat("id-ID", { month: "short" }).format(new Date(entryYear, entryMonth - 1, 1));
  });
  const totals = entries.map((entry) => {
    const totalPandu = (entry.domesticPandu || 0) + (entry.foreignPandu || 0) + (entry.specialPandu || 0);
    const totalTunda = (entry.domesticTunda || 0) + (entry.foreignTunda || 0) + (entry.specialTunda || 0);
    return Math.max(0, Math.round(totalPandu + totalTunda));
  });

  return { labels, totals };
}

function createGradient(ctx, colorTop) {
  const gradient = ctx.createLinearGradient(0, 0, 0, 240);
  gradient.addColorStop(0, colorTop);
  gradient.addColorStop(1, "rgba(9, 29, 40, 0.05)");
  return gradient;
}

function populateMonthOptions() {
  const months = getAvailableMonths();
  monthSelect.innerHTML = months
    .map((period) => `<option value="${period}">${formatMonthLabel(period)}</option>`)
    .join("");
  if (!months.length) {
    monthSelect.value = "";
    return [];
  }

  if (!months.includes(monthSelect.value)) {
    monthSelect.value = months[0];
  }

  return months;
}

function setDefaultMonthFilter() {
  const months = populateMonthOptions();
  if (!months.length) {
    return;
  }

  const todayPeriod = getTodayPeriodValue();
  const defaultPeriod = months.includes(todayPeriod) ? todayPeriod : months[months.length - 1];

  monthSelect.value = defaultPeriod;
}

function updateRealtimeDisplay() {
  const todayDate = getTodayDateValue();

  if (realtimeDateLabel) {
    realtimeDateLabel.textContent = formatRealtimeDateLabel(todayDate);
  }
}

function updateRealtimeWidgets() {
  const dateString = getTodayDateValue();
  const period = dateString.slice(0, 7);
  const year = period.slice(0, 4);
  const yearRows = getRowsByYear(year);
  const yearTotals = aggregateRows(yearRows);
  const monthTotals = aggregateRows(getRowsByMonth(period));
  const dayTotals = aggregateRows(getRowsByDate(dateString));
  const yearActiveMonths = new Set(yearRows.map((row) => row.tanggal.slice(0, 7))).size || 1;
  const avgPerDay = Math.round((monthTotals.total / getDaysInMonth(period)) * 10) / 10;
  const avgPerYearMonth = Math.round((yearTotals.total / yearActiveMonths) * 10) / 10;
  const [periodYear, periodMonth] = period.split("-").map(Number);
  const monthLabel = new Intl.DateTimeFormat("id-ID", { month: "long" }).format(
    new Date(periodYear, periodMonth - 1, 1)
  );

  if (kpiYearLabel) {
    kpiYearLabel.textContent = `Gerakan Tahun ${year}`;
  }
  if (kpiMonthLabel) {
    kpiMonthLabel.textContent = `Gerakan Bulan ${monthLabel}`;
  }

  kpiYearTotal.textContent = formatNumber(yearTotals.total);
  kpiYearTotalUtara.textContent = formatNumber(yearTotals.totalUtara);
  kpiYearTotalSelatan.textContent = formatNumber(yearTotals.totalSelatan);
  kpiYearDomesticInline.textContent = formatNumber(yearTotals.idr);
  kpiYearForeignInline.textContent = formatNumber(yearTotals.usd);

  kpiMonthTotal.textContent = formatNumber(monthTotals.total);
  kpiMonthTotalUtara.textContent = formatNumber(monthTotals.totalUtara);
  kpiMonthTotalSelatan.textContent = formatNumber(monthTotals.totalSelatan);
  kpiMonthDomesticInline.textContent = formatNumber(monthTotals.idr);
  kpiMonthForeignInline.textContent = formatNumber(monthTotals.usd);

  kpiDayTotal.textContent = formatNumber(dayTotals.total);
  kpiDayTotalUtara.textContent = formatNumber(dayTotals.totalUtara);
  kpiDayTotalSelatan.textContent = formatNumber(dayTotals.totalSelatan);
  kpiDayDomesticInline.textContent = formatNumber(dayTotals.idr);
  kpiDayForeignInline.textContent = formatNumber(dayTotals.usd);
  kpiAverageTotal.textContent = formatNumber(avgPerDay);
  kpiAverageYearTotal.textContent = formatNumber(avgPerYearMonth);
}

function startRealtimeWidgets() {
  updateRealtimeDisplay();
  updateRealtimeWidgets();
  window.setInterval(() => {
    updateRealtimeDisplay();
    updateRealtimeWidgets();
  }, 1000);
}

const baseChartOptions = {
  responsive: true,
  maintainAspectRatio: false,
  plugins: {
    legend: {
      labels: {
        color: "#21586f",
        font: { family: "Avenir, Avenir Next, Segoe UI, sans-serif", size: 12 },
      },
    },
    tooltip: {
      backgroundColor: "rgba(255, 255, 255, 0.98)",
      borderColor: "rgba(29, 183, 196, 0.4)",
      borderWidth: 1,
      titleColor: "#073347",
      bodyColor: "#21586f",
      padding: 10,
    },
  },
  scales: {
    x: {
      grid: { color: "rgba(29, 183, 196, 0.18)" },
      ticks: { color: "#4b7387" },
    },
    y: {
      grid: { color: "rgba(29, 183, 196, 0.18)" },
      ticks: { color: "#4b7387" },
    },
  },
};

function createCharts(period) {
  const series = getSeriesByMonth(period);
  const yearSeries = getSeriesByYear(period.slice(0, 4));
  const invoiceSeries = getInvoiceNominalSeriesByYear(period.slice(0, 4));
  const monthlyCombinedTotal = sumSeriesValues(series.utara.total, series.selatan.total);
  const yearlyCombinedTotal = sumSeriesValues(yearSeries.utaraTotal, yearSeries.selatanTotal);

  const chartTotalCtx = document.getElementById("chartTotal").getContext("2d");
  const chartTotalYearlyCtx = document.getElementById("chartTotalYearly").getContext("2d");
  const chartInvoiceMonthlyCtx = document.getElementById("chartInvoiceMonthly").getContext("2d");
  const chartDailyStackedCtx = document.getElementById("chartDailyStacked").getContext("2d");

  chartTotal = new Chart(chartTotalCtx, {
    type: "line",
    data: {
      labels: series.labels,
      datasets: [
        {
          label: "Total",
          data: monthlyCombinedTotal,
          borderColor: "#42d67f",
          fill: false,
          tension: 0.35,
          borderWidth: 2.4,
          pointRadius: 3,
        },
        {
          label: "Utara",
          data: series.utara.total,
          borderColor: "#25c3c8",
          fill: false,
          tension: 0.35,
          borderWidth: 2.2,
          pointRadius: 3,
        },
        {
          label: "Selatan",
          data: series.selatan.total,
          borderColor: "#ff9d4d",
          fill: false,
          tension: 0.35,
          borderWidth: 2.2,
          pointRadius: 3,
        },
        {
          label: "IDR - Utara",
          data: series.utara.idr,
          borderColor: "#25c3c8",
          fill: false,
          tension: 0.35,
          borderWidth: 2,
          pointRadius: 2,
          borderDash: [8, 6],
        },
        {
          label: "IDR - Selatan",
          data: series.selatan.idr,
          borderColor: "#ff9d4d",
          fill: false,
          tension: 0.35,
          borderWidth: 2,
          pointRadius: 2,
          borderDash: [8, 6],
        },
        {
          label: "USD - Utara",
          data: series.utara.usd,
          borderColor: "#0e7e89",
          fill: false,
          tension: 0.35,
          borderWidth: 1.8,
          pointRadius: 2,
          borderDash: [2, 6],
        },
        {
          label: "USD - Selatan",
          data: series.selatan.usd,
          borderColor: "#a34f18",
          fill: false,
          tension: 0.35,
          borderWidth: 1.8,
          pointRadius: 2,
          borderDash: [2, 6],
        },
      ],
    },
    options: baseChartOptions,
  });

  chartTotalYearly = new Chart(chartTotalYearlyCtx, {
    type: "line",
    data: {
      labels: yearSeries.labels,
      datasets: [
        {
          label: "Total",
          data: yearlyCombinedTotal,
          borderColor: "#42d67f",
          fill: false,
          tension: 0.35,
          borderWidth: 2.4,
          pointRadius: 3,
        },
        {
          label: "Utara",
          data: yearSeries.utaraTotal,
          borderColor: "#25c3c8",
          fill: false,
          tension: 0.35,
          borderWidth: 2.2,
          pointRadius: 3,
        },
        {
          label: "Selatan",
          data: yearSeries.selatanTotal,
          borderColor: "#ff9d4d",
          fill: false,
          tension: 0.35,
          borderWidth: 2.2,
          pointRadius: 3,
        },
        {
          label: "IDR - Utara",
          data: yearSeries.utaraIdr,
          borderColor: "#25c3c8",
          fill: false,
          tension: 0.35,
          borderWidth: 2,
          pointRadius: 2,
          borderDash: [8, 6],
        },
        {
          label: "IDR - Selatan",
          data: yearSeries.selatanIdr,
          borderColor: "#ff9d4d",
          fill: false,
          tension: 0.35,
          borderWidth: 2,
          pointRadius: 2,
          borderDash: [8, 6],
        },
        {
          label: "USD - Utara",
          data: yearSeries.utaraUsd,
          borderColor: "#0e7e89",
          fill: false,
          tension: 0.35,
          borderWidth: 1.8,
          pointRadius: 2,
          borderDash: [2, 6],
        },
        {
          label: "USD - Selatan",
          data: yearSeries.selatanUsd,
          borderColor: "#a34f18",
          fill: false,
          tension: 0.35,
          borderWidth: 1.8,
          pointRadius: 2,
          borderDash: [2, 6],
        },
      ],
    },
    options: baseChartOptions,
  });

  chartInvoiceMonthly = new Chart(chartInvoiceMonthlyCtx, {
    type: "line",
    data: {
      labels: invoiceSeries.labels,
      datasets: [
        {
          label: "Total Invoice Bulanan",
          data: invoiceSeries.totals,
          borderColor: "#25c3c8",
          backgroundColor: "rgba(37, 195, 200, 0.16)",
          fill: true,
          tension: 0.35,
          borderWidth: 2.2,
          pointRadius: 3,
        },
      ],
    },
    options: {
      ...baseChartOptions,
      scales: {
        ...baseChartOptions.scales,
        y: {
          ...baseChartOptions.scales.y,
          beginAtZero: true,
        },
      },
    },
  });

  const stackedChartOptions = {
    ...baseChartOptions,
    scales: {
      x: { ...baseChartOptions.scales.x, stacked: true },
      y: { ...baseChartOptions.scales.y, stacked: true },
    },
  };

  chartDailyStacked = new Chart(chartDailyStackedCtx, {
    type: "bar",
    data: {
      labels: series.labels,
      datasets: [
        {
          label: "Utara",
          data: series.utara.total,
          backgroundColor: "rgba(37, 195, 200, 0.65)",
          borderRadius: 6,
          stack: "daily",
        },
        {
          label: "Selatan",
          data: series.selatan.total,
          backgroundColor: "rgba(255, 157, 77, 0.6)",
          borderRadius: 6,
          stack: "daily",
        },
      ],
    },
    options: stackedChartOptions,
  });

}

function updateCharts(period) {
  const series = getSeriesByMonth(period);
  const yearSeries = getSeriesByYear(period.slice(0, 4));
  const invoiceSeries = getInvoiceNominalSeriesByYear(period.slice(0, 4));
  const monthlyCombinedTotal = sumSeriesValues(series.utara.total, series.selatan.total);
  const yearlyCombinedTotal = sumSeriesValues(yearSeries.utaraTotal, yearSeries.selatanTotal);

  chartTotal.data.labels = series.labels;
  chartTotal.data.datasets[0].data = monthlyCombinedTotal;
  chartTotal.data.datasets[1].data = series.utara.total;
  chartTotal.data.datasets[2].data = series.selatan.total;
  chartTotal.data.datasets[3].data = series.utara.idr;
  chartTotal.data.datasets[4].data = series.selatan.idr;
  chartTotal.data.datasets[5].data = series.utara.usd;
  chartTotal.data.datasets[6].data = series.selatan.usd;
  chartTotal.update();

  chartTotalYearly.data.labels = yearSeries.labels;
  chartTotalYearly.data.datasets[0].data = yearlyCombinedTotal;
  chartTotalYearly.data.datasets[1].data = yearSeries.utaraTotal;
  chartTotalYearly.data.datasets[2].data = yearSeries.selatanTotal;
  chartTotalYearly.data.datasets[3].data = yearSeries.utaraIdr;
  chartTotalYearly.data.datasets[4].data = yearSeries.selatanIdr;
  chartTotalYearly.data.datasets[5].data = yearSeries.utaraUsd;
  chartTotalYearly.data.datasets[6].data = yearSeries.selatanUsd;
  chartTotalYearly.update();

  chartInvoiceMonthly.data.labels = invoiceSeries.labels;
  chartInvoiceMonthly.data.datasets[0].data = invoiceSeries.totals;
  chartInvoiceMonthly.update();

  chartDailyStacked.data.labels = series.labels;
  chartDailyStacked.data.datasets[0].data = series.utara.total;
  chartDailyStacked.data.datasets[1].data = series.selatan.total;
  chartDailyStacked.update();
}

function refreshDashboard() {
  const period = monthSelect.value;
  renderMonthlyRecapTable(period);
  renderExternalRecapTables(period);
  updateCharts(period);
}

async function initApp() {
  renderExternalRecapLoading();
  dashboardData = await loadDashboardData();
  await initRecapTablesData();
  setDefaultMonthFilter();
  renderInvoiceTotalsWidget();
  createCharts(monthSelect.value);
  startRealtimeWidgets();
  renderMonthlyRecapTable(monthSelect.value);
  renderExternalRecapTables(monthSelect.value);
  setupCombinedRecapRowToggle();

  monthSelect.addEventListener("change", () => {
    refreshDashboard();
  });
}

initApp();
