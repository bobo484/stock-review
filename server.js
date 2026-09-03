"use strict";

const path = require("path");
const fs = require("fs");
const express = require("express");
const ExcelJS = require("exceljs");

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

function buildDemand(state, month, code) {
  const { items } = loadCalculator(state, month);
  const rawItem = items.find((it) => String(it.InventoryCode) === code);
  if (!rawItem) {
    const err = new Error(`Item ${code} not in ${state} ${month} calculator`);
    err.status = 404;
    throw err;
  }
  const item = mapRow(rawItem, state);
  const rates = parseQtyPerM(rawItem[`${state}_QtyPerM_JSON`]);
  const productIds = Object.keys(rates).filter((id) => num(rates[id]) > 0);

  let priorRates = {};
  try {
    const prior = loadCalculator(state, prevMonth(month));
    const prevItem = prior.items.find((it) => String(it.InventoryCode) === code);
    if (prevItem) priorRates = parseQtyPerM(prevItem[`${state}_QtyPerM_JSON`]);
  } catch {
    priorRates = {};
  }

  const { records, metadata, file: sfFile } = loadSalesForecast(state, month);
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

  const products = productIds.map((id) => {
    const info = productIndex[id] || { id, code: "", name: `Product ${id}`, rows: [] };
    const rate = num(rates[id]);
    const prior = num(priorRates[id]);
    const months = monthMeta.map((m) => {
      const outLm = info.rows.reduce((s, r) => s + sfNum(r, `Forecast Month${m.index}_c_TotalLM`), 0);
      const netLm = info.rows.reduce((s, r) => s + sfNum(r, `StockReturning Net Month ${m.index}_NetRequired`), 0);
      const hires = info.rows.reduce((s, r) => s + sfNum(r, `Forecast Month${m.index}_NoOfHires`), 0);
      const inLm = outLm - netLm;
      return {
        index: m.index,
        date: m.date,
        workingDays: m.workingDays,
        hires,
        outLm,
        inLm,
        netLm,
        outUnits: outLm * rate,
        inUnits: inLm * rate,
        netUnits: netLm * rate,
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
          hires: 0,
          outLm: 0,
          netLm: 0,
          units: 0,
          products: {},
        };
      }
      const hires = sfNum(rec, `Forecast Month${peakIdx}_NoOfHires`);
      const outLm = sfNum(rec, `Forecast Month${peakIdx}_c_TotalLM`);
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
      return {
        company: c.company,
        manager: c.manager,
        avgDays: w ? c.avgDays / w : 0,
        avgQty: w ? c.avgQty / w : 0,
        hires: c.hires,
        outLm: c.outLm,
        netLm: c.netLm,
        units: c.units,
        products: Object.entries(c.products).map(([k, v]) => ({ product: k, units: v })),
      };
    });

  const required = item.targetUtil > 0 ? Math.ceil(item.forecastMax / item.targetUtil) : item.forecastMax;
  const cover = item.totalStock + item.currentOrders;
  const shortfall = required - cover;

  return {
    state,
    month,
    sfFile,
    sfExport: metadata.exportDate || "",
    item,
    required,
    cover,
    shortfall,
    peakMonth: peak,
    products,
    months: trajectory,
    companies,
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

const app = express();
app.use(express.json({ limit: "4mb" }));
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

app.get("/api/demand", (req, res) => {
  try {
    const state = String(req.query.state || "QLD").toUpperCase();
    const month = String(req.query.month || "");
    const code = String(req.query.code || "");
    if (!STATES.includes(state) || !/^\d{4}-\d{2}$/.test(month) || !code) {
      res.status(400).json({ error: "state, month, code required" });
      return;
    }
    res.json(buildDemand(state, month, code));
  } catch (err) {
    res.status(err.status || 500).json({ error: err.message });
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
  console.log(`Stock Review  http://localhost:${PORT}`);
  console.log(`StockCalculator  ${STOCK_DIR}  reachable=${fs.existsSync(STOCK_DIR)}`);
});
