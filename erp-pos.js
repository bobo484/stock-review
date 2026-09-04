"use strict";

const sql = require("mssql");

const DB_CONFIG = {
  server: process.env.BS_SERVER || "192.168.1.240",
  database: process.env.BS_DATABASE || "ERP",
  user: process.env.BS_USER || "erp_reader",
  password: process.env.BS_PASSWORD || "BSQerpread123!",
  options: { encrypt: false, trustServerCertificate: true, enableArithAbort: true },
  pool: { max: 4, min: 0, idleTimeoutMillis: 30000 },
  requestTimeout: 60000,
  connectionTimeout: 15000,
};

const LOOKBACK_MONTHS = Number(process.env.PO_LOOKBACK_MONTHS || 12);
let poolPromise = null;
const cache = new Map();
const CACHE_MS = 2 * 60 * 1000;

function getPool() {
  if (!poolPromise) poolPromise = sql.connect(DB_CONFIG);
  return poolPromise;
}

function normName(s) {
  return String(s || "")
    .toUpperCase()
    .replace(/[^A-Z0-9]+/g, " ")
    .replace(/\b(PTY|LTD|LIMITED|THE|CO|COMPANY)\b/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function isoDay(d) {
  if (!d) return null;
  const dt = d instanceof Date ? d : new Date(d);
  if (Number.isNaN(dt.getTime())) return null;
  return `${dt.getFullYear()}-${String(dt.getMonth() + 1).padStart(2, "0")}-${String(dt.getDate()).padStart(2, "0")}`;
}

function realPo(s) {
  const po = String(s || "").trim();
  if (!po) return "";
  if (/^(awaiting|tba|tbc|pending|n\/a|na|none|nil|-|\.+)$/i.test(po)) return "";
  return po;
}

function emptyPos() {
  return {
    contracts: 0,
    withPo: 0,
    waiting: 0,
    avgDays: null,
    latestPo: "",
    latestRaised: null,
    poNumbers: [],
    waitingRaised: [],
    raiseRate: 0,
    raise3: 0,
    raise12: 0,
    raisedThisMonth: 0,
  };
}

async function loadCompanyPos(state, productCodes) {
  const codes = (productCodes || []).map((c) => String(c || "").toUpperCase()).filter(Boolean);
  const key = `${state}|${codes.slice().sort().join(",")}`;
  const hit = cache.get(key);
  if (hit && Date.now() - hit.at < CACHE_MS) return hit.map;
  const pool = await getPool();
  const req = pool.request();
  req.input("state", sql.NVarChar(8), state);
  req.input("months", sql.Int, LOOKBACK_MONTHS);
  const inList = codes.length ? codes.map((c, i) => {
    req.input(`p${i}`, sql.NVarChar(20), c);
    return `@p${i}`;
  }).join(",") : "''";
  const result = await req.query(`
    SELECT
      co.CompanyName AS CompanyName,
      c.PurchaseOrder AS PurchaseOrder,
      c.DateRaised AS DateRaised,
      c.HireDate AS HireDate,
      c.InstallDateCompleted AS InstallDateCompleted
    FROM dbo.Contracts c
    JOIN dbo.Companies co ON c.CompanyId = co.Id
    JOIN dbo.Products p ON c.ProductId = p.Id
    JOIN dbo.Branches b ON c.BranchId = b.Id
    LEFT JOIN dbo.Addresses a ON c.AddressId = a.Id
    WHERE c.Status <> -1
      AND NOT (c.Status = 0 AND ISNULL(c.ClosedPo, 0) = 1)
      AND p.ProductCode IN (${inList})
      AND (
        CASE
          WHEN b.BranchCode IN (N'GCS', N'NR', N'NR (BSQ)') THEN N'QLD'
          ELSE a.State
        END
      ) = @state
      AND COALESCE(c.DateRaised, c.HireDate) >= DATEADD(month, -@months, GETDATE())
  `);
  const map = {};
  for (const row of result.recordset) {
    const name = String(row.CompanyName || "").trim();
    if (!name) continue;
    const k = normName(name);
    if (!map[k]) {
      map[k] = {
        company: name,
        contracts: 0,
        withPo: 0,
        waiting: 0,
        daySum: 0,
        dayN: 0,
        pos: [],
        waitingRaised: [],
        raise3: 0,
        raise12: 0,
        raisedThisMonth: 0,
      };
    }
    const rec = map[k];
    rec.contracts += 1;
    const po = realPo(row.PurchaseOrder);
    const raised = row.DateRaised ? new Date(row.DateRaised) : null;
    const onHire = row.HireDate
      ? new Date(row.HireDate)
      : row.InstallDateCompleted
        ? new Date(row.InstallDateCompleted)
        : null;
    if (po) {
      rec.withPo += 1;
      rec.pos.push({ po, raised });
      if (!onHire) {
        rec.waiting += 1;
        if (raised && !Number.isNaN(raised.getTime())) rec.waitingRaised.push(raised.toISOString());
      }
      if (raised && !Number.isNaN(raised.getTime())) {
        rec.raise12 += 1;
        if (raised.getTime() >= Date.now() - 90 * 86400000) rec.raise3 += 1;
        const now = new Date();
        if (raised.getFullYear() === now.getFullYear() && raised.getMonth() === now.getMonth()) {
          rec.raisedThisMonth += 1;
        }
      }
    }
    if (raised && onHire && !Number.isNaN(raised.getTime()) && !Number.isNaN(onHire.getTime())) {
      rec.daySum += Math.round((onHire.getTime() - raised.getTime()) / 86400000);
      rec.dayN += 1;
    }
  }
  const out = {};
  for (const [k, rec] of Object.entries(map)) {
    rec.pos.sort((a, b) => (b.raised || 0) - (a.raised || 0));
    const uniq = [];
    const seen = new Set();
    for (const p of rec.pos) {
      if (!p.po || seen.has(p.po)) continue;
      seen.add(p.po);
      uniq.push(p.po);
      if (uniq.length >= 4) break;
    }
    const payload = {
      contracts: rec.contracts,
      withPo: rec.withPo,
      waiting: rec.waiting,
      avgDays: rec.dayN ? rec.daySum / rec.dayN : null,
      latestPo: uniq[0] || "",
      latestRaised: rec.pos[0] && rec.pos[0].raised ? isoDay(rec.pos[0].raised) : null,
      poNumbers: uniq,
      waitingRaised: rec.waitingRaised.slice(0, 400),
      raise3: rec.raise3,
      raise12: rec.raise12,
      raisedThisMonth: rec.raisedThisMonth,
      raiseRate: rec.raise3 >= 2 ? rec.raise3 / 3 : rec.raise12 / 12,
    };
    out[k] = payload;
    out[normName(rec.company)] = payload;
  }
  cache.set(key, { at: Date.now(), map: out });
  return out;
}

function attachPos(companies, posMap) {
  return (companies || []).map((c) => {
    const rec = posMap[normName(c.company)] || emptyPos();
    return { ...c, po: rec };
  });
}

module.exports = { loadCompanyPos, attachPos, normName };
