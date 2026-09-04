"use strict";

const path = require("path");
const fs = require("fs");
const os = require("os");
const express = require("express");
const ExcelJS = require("exceljs");
const { loadCompanyPos, attachPos } = require("./erp-pos");

const PORT = Number(process.env.PORT || 5174);
const STOCK_DIR =
  process.env.STOCK_CALCULATOR_DIR ||
  "\\\\fs\\apps\\SalesForecast\\Data-DatabaseInfoSchema\\StockCalculator";
const FORECAST_DIR =
  process.env.SALES_FORECAST_DIR ||
  "\\\\fs\\apps\\SalesForecast\\Data-DatabaseInfoSchema\\SalesForecast";
const DATA_DIR = path.join(__dirname, "data");
const DECISIONS_FILE = path.join(DATA_DIR, "decisions.json");
const STATES = ["QLD", "NSW", "VIC", "SA"];

function num(v) {
  if (v == null || v === "" || v === "?") return 0;
  const n = Number(String(v).replace(/,/g, ""));
  return Number.isFinite(n) ? n : 0;
}

function unwrapItems(parsed) {
  let items = parsed;
  while (Array.isArray(items) && items.length === 1 && Array.isArray(items[0])) {
    items = items[0];
  }
  if (
    Array.isArray(items) &&
    items.length === 1 &&
    items[0] &&
    Array.isArray(items[0].InventoryCode)
  ) {
    const col = items[0];
    const keys = Object.keys(col);
    const n = col.InventoryCode.length;
    const rows = [];
    for (let i = 0; i < n; i++) {
      const row = {};
      for (const k of keys) {
        const v = col[k];
        row[k] = Array.isArray(v) ? v[i] : v;
      }
      rows.push(row);
    }
    return rows;
  }
  return Array.isArray(items) ? items : [];
}

function listFiles() {
  if (!fs.existsSync(STOCK_DIR)) return [];
  return fs
    .readdirSync(STOCK_DIR)
    .filter((f) => /^[A-Z]{2,3}-\d{4}-\d{2}\.json$/i.test(f))
    .map((f) => {
      const m = f.match(/^([A-Z]{2,3})-(\d{4}-\d{2})\.json$/i);
      const full = path.join(STOCK_DIR, f);
      const st = fs.statSync(full);
      return {
        file: f,
        state: m[1].toUpperCase(),
        month: m[2],
        modified: st.mtime.toISOString(),
      };
    })
    .sort((a, b) => (a.month === b.month ? a.state.localeCompare(b.state) : b.month.localeCompare(a.month)));
}

function loadCalculator(state, month) {
  const file = `${state}-${month}.json`;
  const full = path.join(STOCK_DIR, file);
  if (!fs.existsSync(full)) {
    const err = new Error(`No StockCalculator file for ${state} ${month}`);
    err.status = 404;
    throw err;
  }
  const raw = fs.readFileSync(full, "utf8");
  const items = unwrapItems(JSON.parse(raw));
  const st = fs.statSync(full);
  return { file, full, items, modified: st.mtime.toISOString() };
}

function suggestedInto(it, state) {
  const others = STATES.filter((s) => s !== state);
  const out = {};
  for (const src of others) {
    out[src] = num(it[`c_${state}_From${src}_Suggested`] ?? it[`${state}_From${src}`]);
  }
  return out;
}

function mapRow(it, state) {
  const landed = num(it.LandedCost);
  const orderQty = num(it[`${state}_OrderNumber`]);
  const sysQty = num(it[`${state}_SystemRequestedQty`]);
  const target = num(it[`${state}_ForcastUtilisation`]) || num(it.UtilisationTarget) || num(it.c_UtilisationTarget);
  return {
    code: String(it.InventoryCode || ""),
    name: String(it.InventoryName || ""),
    family: String(it.Family || ""),
    supplier: String(it.PreferredSupplier || ""),
    landed,
    onHand: num(it[`${state}_StockOnHand`]),
    inService: num(it[`${state}_StockInService`]),
    totalStock: num(it[`${state}_StockTotal`]),
    util: num(it[`${state}_Utilisation`]),
    targetUtil: target,
    currentOrders: num(it[`${state}_CurrentOrders`]),
    surplus: num(it[`${state}_Surplus`]),
    forecastMax: num(it[`${state}_ForcastMaxInService`]),
    orderQty,
    orderCost: num(it[`${state}_OrderCost`]) || orderQty * landed,
    sysQty,
    sysCost: sysQty * landed,
    suggested: suggestedInto(it, state),
    proposedSurplus: num(it[`${state}_ProposedSurplus`]),
    proposedUtil: num(it[`${state}_ProposedUtil`]),
    accessType: String(it.AccessType || ""),
  };
}

function loadDecisions() {
  try {
    return JSON.parse(fs.readFileSync(DECISIONS_FILE, "utf8"));
  } catch {
    return {};
  }
}

function saveDecisions(all) {
  fs.mkdirSync(DATA_DIR, { recursive: true });
  fs.writeFileSync(DECISIONS_FILE, JSON.stringify(all, null, 2));
}

function decisionKey(state, month) {
  return `${state}-${month}`;
}

function parseQtyPerM(raw) {
  if (!raw) return {};
  if (typeof raw === "object") return raw;
  try {
    return JSON.parse(String(raw));
  } catch {
    return {};
  }
}

function prevMonth(month) {
  const [y, m] = month.split("-").map(Number);
  const d = new Date(y, m - 2, 1);
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
}

function loadSalesForecast(state, month) {
  const file = `${state}-${month}.json`;
  const full = path.join(FORECAST_DIR, file);
  if (!fs.existsSync(full)) {
    const err = new Error(`No SalesForecast file for ${state} ${month}`);
    err.status = 404;
    throw err;
  }
  const parsed = JSON.parse(fs.readFileSync(full, "utf8"));
  const wrap = Array.isArray(parsed) ? parsed[0] : parsed;
  const records = wrap.records || unwrapItems(wrap);
  return { file, records: Array.isArray(records) ? records : [], metadata: wrap.metadata || {} };
}

function sfNum(row, field) {
  return num(row[field]);
}

function sfForecastLm(row, index) {
  const total = sfNum(row, `Forecast Month${index}_c_TotalLM`);
  if (total) return total;
  const submitted = sfNum(row, `Forecast Month${index}_c_Submitted_LM`);
  if (submitted) return submitted;
  const ave = sfNum(row, `Forecast Month${index}_c_TotalLM_UsingAve`);
  if (ave) return ave;
  return sfNum(row, `Forecast Month${index}_TotalLM_Override`);
}

function buildDurationSignal({ products, productIndex, priorIx, peakIdx, peak, priorMonthId }) {
  let wDays = 0;
  let w = 0;
  let wNowMatched = 0;
  let wPriorDays = 0;
  let wPrior = 0;
  const byProduct = [];
  const movers = [];
  for (const p of products) {
    const info = productIndex[p.id];
    if (!info) continue;
    let pw = 0;
    let pDays = 0;
    let pPriorDays = 0;
    let pPriorW = 0;
    for (const rec of info.rows) {
      const outLm = sfForecastLm(rec, peakIdx);
      if (outLm <= 0) continue;
      const units = outLm * p.rate;
      const days = num(rec.CompanyProduct_cv_AvgDays);
      if (days <= 0) continue;
      const company = String(rec["CompanyProduct_Company Name"] || "Unknown");
      const priorRec = priorIx[`${company}|${p.id}`];
      const priorDays = priorRec ? num(priorRec.CompanyProduct_cv_AvgDays) : 0;
      w += units;
      wDays += days * units;
      pw += units;
      pDays += days * units;
      if (priorDays > 0) {
        wPrior += units;
        wNowMatched += days * units;
        wPriorDays += priorDays * units;
        pPriorW += units;
        pPriorDays += priorDays * units;
        const delta = days - priorDays;
        if (Math.abs(delta) >= 1) {
          movers.push({
            company,
            product: p.code || p.id,
            days,
            priorDays,
            delta,
            units,
            extraUnits: units * (delta / priorDays),
          });
        }
      }
    }
    const daysNow = pw ? pDays / pw : 0;
    const daysWas = pPriorW ? pPriorDays / pPriorW : null;
    byProduct.push({
      code: p.code,
      name: p.name,
      days: daysNow,
      priorDays: daysWas,
      delta: daysWas != null ? daysNow - daysWas : null,
      units: pw,
    });
  }
  const current = w ? wDays / w : 0;
  const compared = wPrior ? wNowMatched / wPrior : null;
  const prior = wPrior ? wPriorDays / wPrior : null;
  const delta = compared != null && prior != null ? compared - prior : null;
  const wd = num(peak.workingDays) || 21;
  const extraUnits = delta != null && wd > 0 ? (peak.outUnits || 0) * (delta / wd) : 0;
  let direction = "stable";
  if (delta != null && Math.abs(delta) >= 0.5) direction = delta > 0 ? "longer" : "shorter";
  const longer = movers.filter((m) => m.delta > 0).sort((a, b) => b.extraUnits - a.extraUnits).slice(0, 8);
  const shorter = movers.filter((m) => m.delta < 0).sort((a, b) => a.extraUnits - b.extraUnits).slice(0, 8);
  return {
    priorMonth: priorMonthId,
    current,
    compared,
    prior,
    delta,
    deltaPct: prior ? (delta / prior) * 100 : null,
    direction,
    extraUnits,
    weightUnits: w,
    comparedUnits: wPrior,
    workingDays: wd,
    byProduct,
    longer,
    shorter,
  };
}

function indexSalesForecast(records) {
  const first = records[0] || {};
  const monthMeta = [1, 2, 3, 4].map((i) => ({
    index: i,
    date: String(first[`CompanyProduct_g_ForecastDate${i}`] || ""),
    workingDays: num(first[`CompanyProduct_g_WorkingDays${i}`]),
  }));
  const productIndex = {};
  for (const rec of records) {
    const id = String(rec.CompanyProduct__ProductID ?? "");
    if (!productIndex[id]) {
      productIndex[id] = {
        id,
        code: String(rec.CompanyProduct_ProductCode || ""),
        name: String(rec.CompanyProduct_ProductName || `Product ${id}`),
        rows: [],
      };
    }
    productIndex[id].rows.push(rec);
  }
  return { monthMeta, productIndex };
}

function productMonthCache(productIndex, monthMeta) {
  const cache = {};
  for (const [id, info] of Object.entries(productIndex)) {
    cache[id] = monthMeta.map((m) => {
      let outLm = 0;
      let netLm = 0;
      let hires = 0;
      for (const r of info.rows) {
        outLm += sfForecastLm(r, m.index);
        netLm += sfNum(r, `StockReturning Net Month ${m.index}_NetRequired`);
        hires += sfNum(r, `Forecast Month${m.index}_NoOfHires`);
      }
      return { outLm, netLm, hires, inLm: outLm - netLm };
    });
  }
  return cache;
}

function computeItemMonths(rawItem, state, monthMeta, productIndex, priorRates, monthCache) {
  const item = mapRow(rawItem, state);
  const rates = parseQtyPerM(rawItem[`${state}_QtyPerM_JSON`]);
  const productIds = Object.keys(rates).filter((id) => num(rates[id]) > 0);
  const products = productIds.map((id) => {
    const info = productIndex[id] || { id, code: "", name: `Product ${id}`, rows: [] };
    const rate = num(rates[id]);
    const prior = num(priorRates[id]);
    const cached = monthCache && monthCache[id];
    const months = monthMeta.map((m, i) => {
      const tot = cached
        ? cached[i]
        : (() => {
            const outLm = info.rows.reduce((s, r) => s + sfForecastLm(r, m.index), 0);
            const netLm = info.rows.reduce((s, r) => s + sfNum(r, `StockReturning Net Month ${m.index}_NetRequired`), 0);
            const hires = info.rows.reduce((s, r) => s + sfNum(r, `Forecast Month${m.index}_NoOfHires`), 0);
            return { outLm, netLm, hires, inLm: outLm - netLm };
          })();
      return {
        index: m.index,
        date: m.date,
        workingDays: m.workingDays,
        hires: tot.hires,
        outLm: tot.outLm,
        inLm: tot.inLm,
        netLm: tot.netLm,
        outUnits: tot.outLm * rate,
        inUnits: tot.inLm * rate,
        netUnits: tot.netLm * rate,
      };
    });
    return {
      id,
      code: info.code,
      name: info.name,
      rate,
      priorRate: prior || null,
      rateDelta: prior ? rate - prior : null,
      companies: info.rows.length,
      months,
    };
  });
  const months = monthMeta.map((m) => {
    const slice = {
      index: m.index,
      date: m.date,
      workingDays: m.workingDays,
      hires: 0,
      outLm: 0,
      inLm: 0,
      netLm: 0,
      outUnits: 0,
      inUnits: 0,
      netUnits: 0,
    };
    for (const p of products) {
      const pm = p.months[m.index - 1];
      slice.hires += pm.hires;
      slice.outLm += pm.outLm;
      slice.inLm += pm.inLm;
      slice.netLm += pm.netLm;
      slice.outUnits += pm.outUnits;
      slice.inUnits += pm.inUnits;
      slice.netUnits += pm.netUnits;
    }
    return slice;
  });
  let running = item.inService;
  const trajectory = months.map((m) => {
    running += m.netUnits;
    return { ...m, projectedInService: running };
  });
  const peak = trajectory.reduce((a, b) => (b.outUnits > a.outUnits ? b : a), trajectory[0] || { index: 1, outUnits: 0 });
  const required = item.targetUtil > 0 ? Math.ceil(item.forecastMax / item.targetUtil) : item.forecastMax;
  const cover = item.totalStock + item.currentOrders;
  return {
    item,
    products,
    months: trajectory,
    peakMonth: peak,
    required,
    cover,
    shortfall: required - cover,
  };
}

function sumSfField(rows, field) {
  if (!field) return 0;
  return (rows || []).reduce((s, r) => s + sfNum(r, field), 0);
}

function buildFamilyTimeline(familyItems, state, records, monthMeta, productIndex) {
  const first = records[0] || {};
  const slots = [
    {
      date: first.CompanyProduct_g_ForecastDate_Minus2,
      role: "history",
      fLm: "Forecast Month Minus 2_c_Submitted_LM",
      aLm: "CompanyProduct_cv_ContractAvgQty_Total_MonthMinus2",
      fH: "Forecast Month Minus 2_NoOfHires",
      aH: "CompanyProduct_cv_ContractCountMonthMinus2",
    },
    {
      date: first.CompanyProduct_g_ForecastDate_Minus1,
      role: "history",
      fLm: "Forecast Month Minus 1_c_Submitted_LM",
      aLm: "CompanyProduct_cv_ContractAvgQty_Total_MonthMinus1",
      fH: "Forecast Month Minus 1_NoOfHires",
      aH: "CompanyProduct_cv_ContractCountMonthMinus1",
    },
    ...monthMeta.map((m) => ({
      date: m.date,
      role: "forecast",
      monthIndex: m.index,
      aLm: null,
      fH: `Forecast Month${m.index}_NoOfHires`,
      aH: null,
    })),
  ];
  const items = (familyItems || []).map((it) => ({
    rates: parseQtyPerM(it[`${state}_QtyPerM_JSON`]),
  }));
  return slots.map((slot) => {
    let forecastOut = 0;
    let actualOut = 0;
    let forecastLm = 0;
    let actualLm = 0;
    let forecastHires = 0;
    let actualHires = 0;
    const seen = new Set();
    for (const it of items) {
      for (const [pid, rate] of Object.entries(it.rates || {})) {
        if (num(rate) <= 0) continue;
        const info = productIndex[pid];
        if (!info) continue;
        const fLm = slot.monthIndex
          ? info.rows.reduce((s, r) => s + sfForecastLm(r, slot.monthIndex), 0)
          : sumSfField(info.rows, slot.fLm);
        const aLm = slot.aLm ? sumSfField(info.rows, slot.aLm) : 0;
        forecastOut += fLm * num(rate);
        actualOut += aLm * num(rate);
        if (!seen.has(pid)) {
          seen.add(pid);
          forecastLm += fLm;
          actualLm += aLm;
          forecastHires += sumSfField(info.rows, slot.fH);
          if (slot.aH) actualHires += sumSfField(info.rows, slot.aH);
        }
      }
    }
    const variance = slot.role === "history" ? actualOut - forecastOut : null;
    return {
      date: slot.date || "",
      key: String(slot.date || "").slice(0, 7),
      role: slot.role,
      forecastOut,
      actualOut: slot.role === "history" ? actualOut : null,
      variance,
      variancePct: slot.role === "history" && forecastOut ? variance / forecastOut : null,
      forecastLm,
      actualLm: slot.role === "history" ? actualLm : null,
      forecastHires,
      actualHires: slot.role === "history" ? actualHires : null,
    };
  });
}

function buildFamilyDemand(items, state, familyName, monthMeta, productIndex, monthCache, records) {
  const familyItems = items.filter((it) => String(it.Family || "") === familyName && String(it.InventoryCode || ""));
  const parts = familyItems.map((it) => computeItemMonths(it, state, monthMeta, productIndex, {}, monthCache));
  const months = monthMeta.map((m, i) => {
    const slice = {
      index: m.index,
      date: m.date,
      workingDays: m.workingDays,
      hires: 0,
      outLm: 0,
      inLm: 0,
      netLm: 0,
      outUnits: 0,
      inUnits: 0,
      netUnits: 0,
      projectedInService: 0,
    };
    for (const p of parts) {
      const pm = p.months[i] || {};
      slice.outUnits += pm.outUnits || 0;
      slice.inUnits += pm.inUnits || 0;
      slice.netUnits += pm.netUnits || 0;
      slice.projectedInService += pm.projectedInService || 0;
    }
    return slice;
  });
  const peak = months.reduce((a, b) => (b.outUnits > a.outUnits ? b : a), months[0] || { index: 1, outUnits: 0 });
  const forecastMax = parts.reduce((s, p) => s + p.item.forecastMax, 0);
  const required = parts.reduce((s, p) => s + p.required, 0);
  const cover = parts.reduce((s, p) => s + p.cover, 0);
  const inService = parts.reduce((s, p) => s + p.item.inService, 0);
  const totalStock = parts.reduce((s, p) => s + p.item.totalStock, 0);
  const currentOrders = parts.reduce((s, p) => s + p.item.currentOrders, 0);
  const orderQty = parts.reduce((s, p) => s + p.item.orderQty, 0);
  const orderCost = parts.reduce((s, p) => s + p.item.orderCost, 0);
  const members = parts
    .map((p) => ({
      code: p.item.code,
      name: p.item.name,
      forecastMax: p.item.forecastMax,
      required: p.required,
      cover: p.cover,
      shortfall: p.shortfall,
      orderQty: p.item.orderQty,
      orderCost: p.item.orderCost,
      inService: p.item.inService,
      totalStock: p.item.totalStock,
      peakOut: p.peakMonth.outUnits || 0,
    }))
    .sort((a, b) => b.orderCost - a.orderCost || b.peakOut - a.peakOut);
  return {
    name: familyName,
    itemCount: parts.length,
    orderLines: members.filter((m) => m.orderQty > 0).length,
    forecastMax,
    required,
    cover,
    shortfall: required - cover,
    inService,
    totalStock,
    currentOrders,
    orderQty,
    orderCost,
    months,
    timeline: buildFamilyTimeline(familyItems, state, records || [], monthMeta, productIndex),
    peakMonth: peak,
    members,
  };
}

function buildConversion(companies, months, peak) {
  const byMonth = (months || []).map((m) => ({
    index: m.index,
    date: m.date,
    key: String(m.date || "").slice(0, 7),
    forecastOut: m.outUnits || 0,
    poUnits: 0,
  }));
  const monthIx = {};
  byMonth.forEach((m, i) => {
    if (m.key) monthIx[m.key] = i;
  });
  let waiting = 0;
  let firmUnits = 0;
  let dayS = 0;
  let dayW = 0;
  const withFirm = (companies || []).map((c) => {
    const po = c.po || {};
    const rate = c.outLm > 0 ? c.units / c.outLm : 0;
    const unitPerHire = (c.avgQty || 0) * rate;
    const fu = (po.waiting || 0) * unitPerHire;
    waiting += po.waiting || 0;
    firmUnits += fu;
    if (po.avgDays != null && fu > 0) {
      dayS += po.avgDays * fu;
      dayW += fu;
    }
    const lag = po.avgDays != null ? po.avgDays : 70;
    for (const iso of po.waitingRaised || []) {
      const d = new Date(iso);
      if (Number.isNaN(d.getTime())) continue;
      d.setDate(d.getDate() + Math.round(lag));
      const key = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
      if (monthIx[key] != null) byMonth[monthIx[key]].poUnits += unitPerHire;
    }
    return { ...c, firmUnits: fu };
  });
  const peakOut = peak && peak.outUnits ? peak.outUnits : 0;
  const peakKey = peak && peak.date ? String(peak.date).slice(0, 7) : "";
  const peakPoUnits = (byMonth.find((m) => m.key === peakKey) || {}).poUnits || 0;
  const firmPct = peakOut > 0 ? peakPoUnits / peakOut : null;
  let confidence = "no-data";
  if (firmPct != null) {
    if (firmPct >= 0.7) confidence = "firm";
    else if (firmPct >= 0.4) confidence = "mixed";
    else confidence = "soft";
  }
  return {
    waiting,
    firmUnits,
    peakPoUnits,
    peakOut,
    firmPct,
    avgDays: dayW ? dayS / dayW : null,
    confidence,
    months: byMonth,
    companies: withFirm,
  };
}

async function buildDemand(state, month, code) {
  const { items } = loadCalculator(state, month);
  const rawItem = items.find((it) => String(it.InventoryCode) === code);
  if (!rawItem) {
    const err = new Error(`Item ${code} not in ${state} ${month} calculator`);
    err.status = 404;
    throw err;
  }

  let priorRates = {};
  try {
    const prior = loadCalculator(state, prevMonth(month));
    const prevItem = prior.items.find((it) => String(it.InventoryCode) === code);
    if (prevItem) priorRates = parseQtyPerM(prevItem[`${state}_QtyPerM_JSON`]);
  } catch {
    priorRates = {};
  }

  const { records, metadata, file: sfFile } = loadSalesForecast(state, month);
  let priorRecords = [];
  const priorMonthId = prevMonth(month);
  try {
    priorRecords = loadSalesForecast(state, priorMonthId).records;
  } catch {
    priorRecords = [];
  }
  const priorIx = {};
  for (const rec of priorRecords) {
    const k = `${rec["CompanyProduct_Company Name"] || ""}|${String(rec.CompanyProduct__ProductID ?? "")}`;
    priorIx[k] = rec;
  }
  const { monthMeta, productIndex } = indexSalesForecast(records);
  const monthCache = productMonthCache(productIndex, monthMeta);
  const computed = computeItemMonths(rawItem, state, monthMeta, productIndex, priorRates, monthCache);
  const item = computed.item;
  const products = computed.products;
  const trajectory = computed.months;
  const peak = computed.peakMonth;
  const familyDemand = buildFamilyDemand(items, state, item.family, monthMeta, productIndex, monthCache, records);

  const peakIdx = peak.index || 1;
  const companyMap = {};
  for (const p of products) {
    const info = productIndex[p.id];
    if (!info) continue;
    for (const rec of info.rows) {
      const key = String(rec["CompanyProduct_Company Name"] || rec.CompanyProduct_CompanyIdFk || "Unknown");
      if (!companyMap[key]) {
        companyMap[key] = {
          company: key,
          manager: String(rec.CompanyProduct_cv_AccountManager || ""),
          avgDays: 0,
          avgQty: 0,
          daysWeight: 0,
          priorDays: 0,
          priorDaysWeight: 0,
          hires: 0,
          outLm: 0,
          netLm: 0,
          units: 0,
          products: {},
        };
      }
      const hires = sfNum(rec, `Forecast Month${peakIdx}_NoOfHires`);
      const outLm = sfForecastLm(rec, peakIdx);
      const netLm = sfNum(rec, `StockReturning Net Month ${peakIdx}_NetRequired`);
      const days = num(rec.CompanyProduct_cv_AvgDays);
      const qty = num(rec.CompanyProduct_cv_AvgQty);
      companyMap[key].hires += hires;
      companyMap[key].outLm += outLm;
      companyMap[key].netLm += netLm;
      companyMap[key].units += outLm * p.rate;
      if (outLm > 0) {
        companyMap[key].avgDays += days * outLm;
        companyMap[key].avgQty += qty * outLm;
        companyMap[key].daysWeight += outLm;
        const pk = `${key}|${p.id}`;
        const priorRec = priorIx[`${key}|${p.id}`];
        const priorDays = priorRec ? num(priorRec.CompanyProduct_cv_AvgDays) : 0;
        if (priorDays > 0) {
          companyMap[key].priorDays += priorDays * outLm;
          companyMap[key].priorDaysWeight += outLm;
        }
      }
      companyMap[key].products[p.code || p.id] = (companyMap[key].products[p.code || p.id] || 0) + outLm * p.rate;
    }
  }
  const companies = Object.values(companyMap)
    .filter((c) => c.outLm || c.hires)
    .sort((a, b) => b.units - a.units)
    .slice(0, 20)
    .map((c) => {
      const w = c.daysWeight || 0;
      const pw = c.priorDaysWeight || 0;
      const avgDays = w ? c.avgDays / w : 0;
      const priorDays = pw ? c.priorDays / pw : null;
      return {
        company: c.company,
        manager: c.manager,
        avgDays,
        priorDays,
        daysDelta: priorDays ? avgDays - priorDays : null,
        avgQty: w ? c.avgQty / w : 0,
        hires: c.hires,
        outLm: c.outLm,
        netLm: c.netLm,
        units: c.units,
        products: Object.entries(c.products).map(([k, v]) => ({ product: k, units: v })),
      };
    });

  const duration = buildDurationSignal({
    products,
    productIndex,
    priorIx,
    peakIdx,
    peak,
    priorMonthId,
  });

  let companiesOut = companies;
  let poSource = "";
  let conversion = null;
  try {
    const productCodes = products.map((p) => p.code).filter(Boolean);
    const posMap = await loadCompanyPos(state, productCodes);
    companiesOut = attachPos(companies, posMap);
    conversion = buildConversion(companiesOut, trajectory, peak);
    companiesOut = conversion.companies;
    poSource = "Live ERP contracts (last 12 months). Days = Date raised → on hire. Firm units = waiting POs × avg LM/hire × qty/m.";
  } catch (err) {
    poSource = `Customer POs unavailable: ${err.message}`;
  }

  return {
    state,
    month,
    sfFile,
    sfExport: metadata.exportDate || "",
    item,
    required: computed.required,
    cover: computed.cover,
    shortfall: computed.shortfall,
    peakMonth: peak,
    products,
    months: trajectory,
    companies: companiesOut,
    duration,
    family: familyDemand,
    poSource,
    conversion,
    note:
      "Units ≈ forecast LM × this item's Qty per metre. FileMaker Forecast Max is peak in-service (already on hire + net starts), not the sum of monthly OUT.",
  };
}

function buildDetails(state, month) {
  const { file, items, modified } = loadCalculator(state, month);
  const first = items[0] || {};
  const rows = items
    .map((it) => mapRow(it, state))
    .filter((r) => r.code)
    .sort((a, b) => a.code.localeCompare(b.code));
  const decisions = loadDecisions()[decisionKey(state, month)] || {};
  for (const r of rows) {
    const d = decisions[r.code];
    r.decision = d == null || d === "" ? null : num(d);
    r.decisionCost = r.decision == null ? 0 : r.decision * r.landed;
  }
  const orderLines = rows.filter((r) => r.orderQty > 0);
  const ask = orderLines.reduce((s, r) => s + r.orderCost, 0);
  const families = {};
  for (const r of rows) {
    const fam = r.family || "(blank)";
    if (!families[fam]) families[fam] = { family: fam, n: 0, orderLines: 0, orderCost: 0, sysCost: 0, decisionCost: 0 };
    families[fam].n += 1;
    if (r.orderQty > 0) {
      families[fam].orderLines += 1;
      families[fam].orderCost += r.orderCost;
      families[fam].sysCost += r.sysCost;
    }
    families[fam].decisionCost += r.decisionCost;
  }
  return {
    state,
    month,
    file,
    modified,
    cacheDate: first.FO_SalesForcastCache_Date || "",
    forecastSubmitted: String(first.FO_ForecastSubmittedStates || "").replace(/\r/g, "+").replace(/\n/g, "+"),
    orderSubmitted: String(first.FO_OrderSubmittedStates || ""),
    limitDate: first.FO_LimitDate || "",
    totals: {
      items: rows.length,
      orderLines: orderLines.length,
      orderQty: orderLines.reduce((s, r) => s + r.orderQty, 0),
      orderCost: ask,
      sysCost: rows.reduce((s, r) => s + r.sysCost, 0),
      decisionCost: rows.reduce((s, r) => s + r.decisionCost, 0),
      shortfallLines: rows.filter((r) => r.surplus < 0).length,
    },
    families: Object.values(families).sort((a, b) => b.orderCost - a.orderCost),
    rows,
  };
}

const STOCKTAKES_FILE = path.join(DATA_DIR, "stocktakes.json");
const FM_FILE = path.join(DATA_DIR, "fm.json");
const BRANCH_EXPORT_DIR = "\\\\fs\\apps\\FilemakerCSVExport\\ERP\\BSQ";

function parseCsv(text) {
  const rows = [];
  let row = [];
  let cell = "";
  let q = false;
  const src = String(text || "").replace(/^\uFEFF/, "");
  for (let i = 0; i < src.length; i++) {
    const ch = src[i];
    if (q) {
      if (ch === '"') {
        if (src[i + 1] === '"') {
          cell += '"';
          i++;
        } else q = false;
      } else cell += ch;
    } else if (ch === '"') q = true;
    else if (ch === "," || ch === "\t") {
      row.push(cell);
      cell = "";
    } else if (ch === "\n") {
      row.push(cell);
      if (row.some((c) => String(c).trim())) rows.push(row);
      row = [];
      cell = "";
    } else if (ch !== "\r") cell += ch;
  }
  if (cell || row.length) {
    row.push(cell);
    if (row.some((c) => String(c).trim())) rows.push(row);
  }
  return rows;
}

function parseDate(v, prefer) {
  if (v == null || v === "" || v === "?") return null;
  if (v instanceof Date && !Number.isNaN(v.getTime())) return v;
  const s = String(v).trim();
  const iso = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (iso) return new Date(Number(iso[1]), Number(iso[2]) - 1, Number(iso[3]));
  const m = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})/);
  if (!m) {
    const t = Date.parse(s);
    return Number.isNaN(t) ? null : new Date(t);
  }
  let a = Number(m[1]);
  let b = Number(m[2]);
  let y = Number(m[3]);
  if (y < 100) y += 2000;
  let day;
  let monthN;
  if (a > 12) {
    day = a;
    monthN = b;
  } else if (b > 12) {
    monthN = a;
    day = b;
  } else if (prefer === "us") {
    monthN = a;
    day = b;
  } else {
    day = a;
    monthN = b;
  }
  return new Date(y, monthN - 1, day);
}

function isoDay(d) {
  if (!d) return null;
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
}

function daysBetween(a, b) {
  if (!a || !b) return null;
  return Math.round((b.getTime() - a.getTime()) / 86400000);
}

function parseBranchStock(raw) {
  if (!raw) return {};
  try {
    const o = typeof raw === "object" ? raw : JSON.parse(String(raw));
    return o && typeof o === "object" && !Array.isArray(o) ? o : {};
  } catch {
    return {};
  }
}

function loadStocktakeImport() {
  try {
    return JSON.parse(fs.readFileSync(STOCKTAKES_FILE, "utf8"));
  } catch {
    return { importedAt: "", source: "", rows: [] };
  }
}

function saveStocktakeImport(payload) {
  fs.mkdirSync(DATA_DIR, { recursive: true });
  fs.writeFileSync(STOCKTAKES_FILE, JSON.stringify(payload, null, 2));
}

function headerKey(name) {
  return String(name || "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "");
}

function rowsFromCsv(text) {
  const grid = parseCsv(text);
  if (!grid.length) return [];
  const headers = grid[0].map(headerKey);
  const idx = {};
  headers.forEach((h, i) => {
    if (h && idx[h] == null) idx[h] = i;
  });
  const col = (row, names) => {
    for (const n of names) {
      const i = idx[headerKey(n)];
      if (i != null) return row[i];
    }
    return "";
  };
  const out = [];
  for (const row of grid.slice(1)) {
    const code = String(col(row, ["InventoryCode", "ItemCode", "code", "Item", "ITEM CODE"]) || "").trim();
    if (!code) continue;
    const last = parseDate(col(row, ["LastStocktake", "cv_LastStocktakeDate", "lastStocktake", "Date", "StocktakeDate"]));
    const nr = String(col(row, ["StocktakeNotRequired", "notRequired", "NotRequired", "Not Req"]) || "").trim();
    out.push({
      code,
      location: String(col(row, ["BranchCode", "Location", "Branch", "LocationCode", "_LocationID"]) || "").trim(),
      locationName: String(col(row, ["BranchName", "LocationName"]) || "").trim(),
      lastStocktake: isoDay(last),
      notRequired: /^(1|true|yes|y)$/i.test(nr),
    });
  }
  return out;
}

function loadBranchMap() {
  try {
    if (!fs.existsSync(BRANCH_EXPORT_DIR)) return [];
    const dirs = fs
      .readdirSync(BRANCH_EXPORT_DIR)
      .filter((d) => fs.existsSync(path.join(BRANCH_EXPORT_DIR, d, "Branch.csv")))
      .sort()
      .reverse();
    if (!dirs.length) return [];
    const grid = parseCsv(fs.readFileSync(path.join(BRANCH_EXPORT_DIR, dirs[0], "Branch.csv"), "utf8"));
    if (grid.length < 2) return [];
    const headers = grid[0].map(headerKey);
    const idx = {};
    headers.forEach((h, i) => {
      if (h && idx[h] == null) idx[h] = i;
    });
    const val = (row, name) => {
      const i = idx[headerKey(name)];
      return i == null ? "" : String(row[i] || "").trim();
    };
    return grid.slice(1).map((row) => ({
      id: val(row, "_BranchID"),
      code: val(row, "BranchCode"),
      name: val(row, "BranchName"),
      state: val(row, "A1_State").toUpperCase(),
    }));
  } catch {
    return [];
  }
}

function branchState(loc, name, branches) {
  const key = String(loc || "").trim().toUpperCase();
  const nm = String(name || "").trim().toUpperCase();
  const hit = branches.find(
    (b) =>
      (key && (b.code.toUpperCase() === key || b.id === key || b.name.toUpperCase() === key)) ||
      (nm && b.name.toUpperCase() === nm)
  );
  return hit ? hit.state : "";
}

function loadFmConfig() {
  try {
    const raw = JSON.parse(fs.readFileSync(FM_FILE, "utf8"));
    return {
      host: String(raw.host || process.env.FM_HOST || "https://fms.bsdomain.local").replace(/\/$/, ""),
      database: String(raw.database || process.env.FM_DATABASE || "Inventory"),
      user: String(raw.user || process.env.FM_USER || ""),
      password: String(raw.password || process.env.FM_PASSWORD || ""),
      layout: String(raw.layout || process.env.FM_LAYOUT || ""),
    };
  } catch {
    return {
      host: String(process.env.FM_HOST || "https://fms.bsdomain.local").replace(/\/$/, ""),
      database: String(process.env.FM_DATABASE || "Inventory"),
      user: String(process.env.FM_USER || ""),
      password: String(process.env.FM_PASSWORD || ""),
      layout: String(process.env.FM_LAYOUT || ""),
    };
  }
}

function httpsJson(url, { method = "GET", headers = {}, body } = {}) {
  return new Promise((resolve, reject) => {
    const req = require("https").request(
      url,
      { method, headers, rejectUnauthorized: false },
      (res) => {
        let data = "";
        res.on("data", (c) => (data += c));
        res.on("end", () => {
          try {
            resolve({ status: res.statusCode, json: data ? JSON.parse(data) : {} });
          } catch (e) {
            reject(new Error(`FileMaker response ${res.statusCode}`));
          }
        });
      }
    );
    req.on("error", reject);
    req.setTimeout(30000, () => req.destroy(new Error("FileMaker timeout")));
    if (body) req.write(body);
    req.end();
  });
}

async function fetchFileMakerStocktakes() {
  const cfg = loadFmConfig();
  if (!cfg.user || !cfg.password) return null;
  const auth = Buffer.from(`${cfg.user}:${cfg.password}`).toString("base64");
  const db = encodeURIComponent(cfg.database);
  const session = await httpsJson(`${cfg.host}/fmi/data/v1/databases/${db}/sessions`, {
    method: "POST",
    headers: { Authorization: `Basic ${auth}`, "Content-Type": "application/json" },
    body: "{}",
  });
  const token = session.json && session.json.response && session.json.response.token;
  if (!token) {
    const msg = (session.json.messages && session.json.messages[0] && session.json.messages[0].message) || "login failed";
    throw new Error(`FileMaker Data API: ${msg}`);
  }
  const hdr = { Authorization: `Bearer ${token}` };
  try {
    let layout = cfg.layout;
    if (!layout) {
      const layouts = await httpsJson(`${cfg.host}/fmi/data/v1/databases/${db}/layouts`, { headers: hdr });
      const names = ((layouts.json.response && layouts.json.response.layouts) || [])
        .flatMap((x) => [x.name, ...((x.folder || x.layouts || []).map((y) => y.name || y))])
        .filter(Boolean);
      layout =
        names.find((n) => /locationitem/i.test(n)) ||
        names.find((n) => /stock\s*take/i.test(n) && !/pending/i.test(n)) ||
        names.find((n) => /monthly stock/i.test(n)) ||
        "";
    }
    if (!layout) throw new Error("No LocationItem / Stock Take layout on Inventory Data API");
    const rows = [];
    let offset = 1;
    const limit = 100;
    for (;;) {
      const page = await httpsJson(
        `${cfg.host}/fmi/data/v1/databases/${db}/layouts/${encodeURIComponent(layout)}/records?_limit=${limit}&_offset=${offset}`,
        { headers: hdr }
      );
      const recs = (page.json.response && page.json.response.data) || [];
      for (const rec of recs) {
        const f = rec.fieldData || {};
        const code = String(f.InventoryCode || f.ItemCode || f["LocationItem::InventoryCode"] || "").trim();
        if (!code) continue;
        const last = parseDate(
          f.LastStocktake || f.cv_LastStocktakeDate || f["LocationItem::LastStocktake"] || f.Date,
          "us"
        );
        const nr = f.StocktakeNotRequired || f["LocationItem::StocktakeNotRequired"];
        rows.push({
          code,
          location: String(f.BranchCode || f.Location || f._LocationID || "").trim(),
          locationName: String(f.BranchName || f.LocationName || "").trim(),
          lastStocktake: isoDay(last),
          notRequired: nr === 1 || nr === "1" || nr === true,
        });
      }
      if (recs.length < limit) break;
      offset += recs.length;
      if (offset > 20000) break;
    }
    return { importedAt: new Date().toISOString(), source: `FileMaker ${cfg.database} / ${layout}`, rows };
  } finally {
    await httpsJson(`${cfg.host}/fmi/data/v1/databases/${db}/sessions/${token}`, {
      method: "DELETE",
      headers: hdr,
    }).catch(() => {});
  }
}

function buildStocktakes(state, month, staleDays) {
  const details = buildDetails(state, month);
  const { items } = loadCalculator(state, month);
  const byCode = {};
  for (const it of items) {
    const code = String(it.InventoryCode || "");
    if (!code) continue;
    byCode[code] = parseBranchStock(it[`${state}_BranchStockJSON`]);
  }
  const cacheDate = parseDate(details.cacheDate) || parseDate(`${month}-01`);
  const branches = loadBranchMap();
  const imported = loadStocktakeImport();
  const takes = (imported.rows || []).filter((r) => {
    if (!r.location && !r.locationName) return true;
    const st = branchState(r.location, r.locationName, branches);
    return !st || st === state;
  });
  const byItem = {};
  for (const t of takes) {
    if (!byItem[t.code]) byItem[t.code] = [];
    byItem[t.code].push(t);
  }
  const rows = details.rows.map((r) => {
    const lines = byItem[r.code] || [];
    const required = lines.filter((x) => !x.notRequired);
    const dates = required.map((x) => parseDate(x.lastStocktake)).filter(Boolean);
    const last = dates.length ? new Date(Math.max(...dates.map((d) => d.getTime()))) : null;
    const oldest = dates.length ? new Date(Math.min(...dates.map((d) => d.getTime()))) : null;
    const never = required.filter((x) => !x.lastStocktake).length + (lines.length ? 0 : 0);
    const daysBefore = daysBetween(last, cacheDate);
    let status = "no-data";
    if (!imported.rows.length) status = "no-data";
    else if (!required.length && lines.some((x) => x.notRequired) && !dates.length) status = "not-required";
    else if (!last) status = "never";
    else if (last > cacheDate) status = "after-forecast";
    else if (daysBefore != null && daysBefore > staleDays) status = "stale";
    else status = "fresh";
    const branchStock = byCode[r.code] || {};
    const branchCount = Object.keys(branchStock).length;
    return {
      code: r.code,
      name: r.name,
      family: r.family,
      onHand: r.onHand,
      inService: r.inService,
      totalStock: r.totalStock,
      orderQty: r.orderQty,
      lastStocktake: isoDay(last),
      oldestStocktake: isoDay(oldest),
      daysBeforeForecast: daysBefore,
      status,
      branches: required.length || branchCount,
      neverCounted: lines.length ? required.filter((x) => !x.lastStocktake).length : null,
      notRequired: lines.filter((x) => x.notRequired).length,
      locations: required
        .filter((x) => x.lastStocktake || x.location)
        .sort((a, b) => String(a.lastStocktake || "").localeCompare(String(b.lastStocktake || "")))
        .slice(0, 8)
        .map((x) => ({
          location: x.location || x.locationName,
          lastStocktake: x.lastStocktake,
          notRequired: x.notRequired,
        })),
    };
  });
  const counted = rows.filter((r) => r.status !== "no-data");
  const stale = rows.filter((r) => r.status === "stale" || r.status === "never");
  const families = {};
  for (const r of rows) {
    const fam = r.family || "(blank)";
    if (!families[fam]) families[fam] = { family: fam, n: 0, stale: 0, never: 0, fresh: 0 };
    families[fam].n += 1;
    if (r.status === "stale") families[fam].stale += 1;
    if (r.status === "never") families[fam].never += 1;
    if (r.status === "fresh") families[fam].fresh += 1;
  }
  return {
    state,
    month,
    cacheDate: isoDay(cacheDate),
    staleDays,
    source: imported.source || "",
    importedAt: imported.importedAt || "",
    hasData: Boolean(imported.rows && imported.rows.length),
    totals: {
      items: rows.length,
      withDate: rows.filter((r) => r.lastStocktake).length,
      stale: stale.length,
      never: rows.filter((r) => r.status === "never").length,
      fresh: rows.filter((r) => r.status === "fresh").length,
      afterForecast: rows.filter((r) => r.status === "after-forecast").length,
      notRequired: rows.filter((r) => r.status === "not-required").length,
      onAskStale: stale.filter((r) => r.orderQty > 0).length,
      noData: rows.filter((r) => r.status === "no-data").length,
    },
    families: Object.values(families).sort((a, b) => b.stale + b.never - (a.stale + a.never) || a.family.localeCompare(b.family)),
    rows,
    counted: counted.length,
  };
}

const app = express();
app.use(express.json({ limit: "8mb" }));
app.use(express.static(path.join(__dirname, "public")));

app.get("/api/health", (_req, res) => {
  res.json({
    ok: true,
    stockDir: STOCK_DIR,
    reachable: fs.existsSync(STOCK_DIR),
    files: listFiles().length,
  });
});

app.get("/api/files", (_req, res) => {
  res.json({ files: listFiles() });
});

app.get("/api/demand", async (req, res) => {
  try {
    const state = String(req.query.state || "QLD").toUpperCase();
    const month = String(req.query.month || "");
    const code = String(req.query.code || "");
    if (!STATES.includes(state) || !/^\d{4}-\d{2}$/.test(month) || !code) {
      res.status(400).json({ error: "state, month, code required" });
      return;
    }
    res.json(await buildDemand(state, month, code));
  } catch (err) {
    res.status(err.status || 500).json({ error: err.message });
  }
});

app.get("/api/stocktakes", (req, res) => {
  try {
    const state = String(req.query.state || "QLD").toUpperCase();
    const month = String(req.query.month || "");
    const staleDays = Math.max(1, Number(req.query.staleDays || 14) || 14);
    if (!STATES.includes(state) || !/^\d{4}-\d{2}$/.test(month)) {
      res.status(400).json({ error: "state and month=YYYY-MM required" });
      return;
    }
    res.json(buildStocktakes(state, month, staleDays));
  } catch (err) {
    res.status(err.status || 500).json({ error: err.message });
  }
});

app.post("/api/stocktakes/import", (req, res) => {
  try {
    const csv = String(req.body.csv || "");
    const rows = rowsFromCsv(csv);
    if (!rows.length) {
      res.status(400).json({ error: "No item rows. Export ItemCode, LastStocktake, Branch/Location, StocktakeNotRequired." });
      return;
    }
    const payload = {
      importedAt: new Date().toISOString(),
      source: String(req.body.source || "Stocktake List CSV"),
      rows,
    };
    saveStocktakeImport(payload);
    res.json({ ok: true, rows: rows.length, importedAt: payload.importedAt });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.post("/api/stocktakes/refresh", async (req, res) => {
  try {
    const pulled = await fetchFileMakerStocktakes();
    if (!pulled) {
      res.status(400).json({
        error:
          "FileMaker last-stocktake dates are not in the forecast JSON. Add data/fm.json (user, password) or import a Stocktake List CSV.",
      });
      return;
    }
    saveStocktakeImport(pulled);
    res.json({ ok: true, rows: pulled.rows.length, source: pulled.source, importedAt: pulled.importedAt });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.get("/api/details", (req, res) => {
  try {
    const state = String(req.query.state || "QLD").toUpperCase();
    const month = String(req.query.month || "");
    if (!STATES.includes(state) || !/^\d{4}-\d{2}$/.test(month)) {
      res.status(400).json({ error: "state and month=YYYY-MM required" });
      return;
    }
    res.json(buildDetails(state, month));
  } catch (err) {
    res.status(err.status || 500).json({ error: err.message });
  }
});

app.put("/api/decisions", (req, res) => {
  const state = String(req.body.state || "").toUpperCase();
  const month = String(req.body.month || "");
  const code = String(req.body.code || "");
  if (!STATES.includes(state) || !/^\d{4}-\d{2}$/.test(month) || !code) {
    res.status(400).json({ error: "state, month, code required" });
    return;
  }
  const all = loadDecisions();
  const key = decisionKey(state, month);
  if (!all[key]) all[key] = {};
  if (req.body.decision == null || req.body.decision === "") delete all[key][code];
  else all[key][code] = num(req.body.decision);
  saveDecisions(all);
  res.json({ ok: true, decision: all[key][code] ?? null });
});

app.get("/api/export.xlsx", async (req, res) => {
  try {
    const state = String(req.query.state || "QLD").toUpperCase();
    const month = String(req.query.month || "");
    const data = buildDetails(state, month);
    const others = STATES.filter((s) => s !== state);
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("DETAILS");
    ws.addRow(Array.from({ length: 22 }, (_, i) => i + 1));
    const group = ws.addRow([]);
    group.getCell(6).value = state;
    group.getCell(20).value = "Proposed";
    ws.mergeCells(2, 6, 2, 19);
    ws.mergeCells(2, 20, 2, 22);
    ws.addRow([
      "ITEM CODE",
      "DESCRIPTION",
      "FAMILY",
      "SUPPLIER",
      "LANDED COST",
      "On Hand",
      "In Service",
      "Total Stock",
      "Utilisation",
      "Current Orders",
      "Surplus / Shortfall",
      "Order Qty",
      `${others[0]} Suggested`,
      `${others[1]} Suggested`,
      `${others[2]} Suggested`,
      `${state} Decision`,
      `${state} Decision Cost`,
      `${state} System Requested Qty`,
      `${state} System Requested Cost`,
      "Surplus / Shortfall",
      "Utilisation",
      `${state} Order Cost`,
    ]);
    for (const r of data.rows) {
      ws.addRow([
        r.code,
        r.name,
        r.family,
        r.supplier,
        r.landed,
        r.onHand,
        r.inService,
        r.totalStock,
        r.util,
        r.currentOrders,
        r.surplus,
        r.orderQty || null,
        r.suggested[others[0]] || null,
        r.suggested[others[1]] || null,
        r.suggested[others[2]] || null,
        r.decision,
        r.decisionCost || null,
        r.sysQty,
        r.sysCost,
        r.proposedSurplus,
        r.proposedUtil,
        r.orderCost || null,
      ]);
    }
    ws.getRow(3).font = { bold: true };
    ws.columns.forEach((c, i) => {
      c.width = i === 1 ? 36 : i === 3 ? 28 : 14;
    });
    const totals = wb.addWorksheet("TOTALS");
    totals.addRow(["Original Ask"]);
    totals.addRow(["Family", `Sum of ${state} Order Cost`, `Sum of ${state} System Requested Cost`]);
    for (const f of data.families) {
      totals.addRow([f.family, f.orderCost, f.sysCost]);
    }
    totals.addRow(["Grand Total", data.totals.orderCost, data.totals.sysCost]);
    totals.addRow([]);
    totals.addRow(["Decision"]);
    totals.addRow(["Family", `Sum of ${state} Decision Cost`]);
    for (const f of data.families) {
      totals.addRow([f.family, f.decisionCost]);
    }
    totals.addRow(["Grand Total", data.totals.decisionCost]);

    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader(
      "Content-Disposition",
      `attachment; filename="${state}-${month}-Stock-Review-DETAILS.xlsx"`
    );
    await wb.xlsx.write(res);
    res.end();
  } catch (err) {
    res.status(err.status || 500).json({ error: err.message });
  }
});

app.listen(PORT, () => {
  const host = os.hostname();
  console.log(`Stock Review is running. Leave this window open.`);
  console.log(`  This screen:  http://localhost:${PORT}/?state=QLD&month=2026-09&code=SS-DS24`);
  console.log(`  This PC:      http://${host}:${PORT}/?state=QLD&month=2026-09&code=SS-DS24`);
  console.log(`StockCalculator  ${STOCK_DIR}  reachable=${fs.existsSync(STOCK_DIR)}`);
});
