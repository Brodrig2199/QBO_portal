const express = require("express");
const session = require("express-session");
const fetch = require("node-fetch");
const { Pool } = require("pg");
require("dotenv").config();

const app = express();
app.set("view engine", "ejs");
app.use(express.urlencoded({ extended: true }));
app.use(express.json());
app.set("trust proxy", 1);

app.use(
  session({
    secret: process.env.SESSION_SECRET || "secret",
    resave: false,
    saveUninitialized: false,
    cookie: {
      httpOnly: true,
      sameSite: "lax",
      secure: process.env.NODE_ENV === "production",
    },
  })
);

function auth(req, res, next) {
  if (req.session?.user) return next();
  return res.redirect("/login");
}

// ====== Postgres ======
const pool = new Pool({
  connectionString: process.env.DATABASE_URL,
  ssl:
    process.env.NODE_ENV === "production"
      ? { rejectUnauthorized: false }
      : false,
});

async function listCompanies() {
  const { rows } = await pool.query(
    "select id, name, realm_id from companies where is_active=true order by created_at desc"
  );
  return rows;
}

async function upsertCompany({ id, name, realmId }) {
  await pool.query(
    `insert into companies (id, name, realm_id)
     values ($1,$2,$3)
     on conflict (id) do update set
       name=excluded.name,
       realm_id=excluded.realm_id,
       is_active=true`,
    [id, name, realmId]
  );
}

// ====== Static report types ======
const REPORTS = [
  { id: "ProfitAndLoss", name: "Profit & Loss" },
  { id: "BalanceSheet", name: "Balance Sheet" },
  { id: "TrialBalance", name: "Trial Balance" },
  { id: "GeneralLedger", name: "General Ledger" },
];

// ====== Helpers to call n8n ======
async function n8nGet(url) {
  if (!process.env.ALIADA_WEBHOOK_KEY) throw new Error("Falta ALIADA_WEBHOOK_KEY");
  const r = await fetch(url, { headers: { "x-aliada-key": process.env.ALIADA_WEBHOOK_KEY } });
  if (!r.ok) throw new Error(await r.text());
  return r.json();
}

async function n8nPost(url, body) {
  if (!process.env.ALIADA_WEBHOOK_KEY) throw new Error("Falta ALIADA_WEBHOOK_KEY");
  const r = await fetch(url, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "x-aliada-key": process.env.ALIADA_WEBHOOK_KEY,
    },
    body: JSON.stringify(body),
  });
  return r;
}

// ====== Routes ======
app.get("/login", (req, res) => res.render("login", { error: null }));

app.post("/login", (req, res) => {
  const { username, password } = req.body;
  if (username === process.env.APP_USER && password === process.env.APP_PASS) {
    req.session.user = { username };
    return res.redirect("/");
  }
  res.status(401).render("login", { error: "Credenciales incorrectas" });
});

app.post("/logout", (req, res) => req.session.destroy(() => res.redirect("/login")));

// Home: form
app.get("/", auth, async (req, res) => {
  const companies = await listCompanies();
  res.render("form", {
    user: req.session.user,
    companies,
    reports: REPORTS,
    error: null,
  });
});

// Admin: create/update company
app.post("/admin/companies", auth, async (req, res) => {
  try {
    const id = String(req.body.id || "").trim();
    const name = String(req.body.name || "").trim();
    const realmId = String(req.body.realmId || "").trim();

    if (!id || !name || !realmId) throw new Error("Faltan campos (id, name, realmId).");
    await upsertCompany({ id, name, realmId });
    res.redirect("/");
  } catch (e) {
    const companies = await listCompanies();
    res.status(400).render("form", {
      user: req.session.user,
      companies,
      reports: REPORTS,
      error: e.message,
    });
  }
});

// Proxy: get accounts from n8n by realmId
app.get("/api/meta/accounts", auth, async (req, res) => {
  try {
    const realmId = String(req.query.realmId || "").trim();
    if (!realmId) throw new Error("Falta realmId.");

    if (!process.env.N8N_META_ACCOUNTS_URL) throw new Error("Falta N8N_META_ACCOUNTS_URL.");

    const url = `${process.env.N8N_META_ACCOUNTS_URL}?realmId=${encodeURIComponent(realmId)}`;
    res.json(await n8nGet(url));
  } catch (e) {
    res.status(400).json({ error: e.message });
  }
});

// Run report: call n8n and return XLSX to browser
app.post("/api/run-report", auth, async (req, res) => {
  try {
    const { realmId, reportType, startDate, endDate, excludeAccountIds } = req.body;

    if (!realmId) throw new Error("Falta realmId.");
    if (!reportType) throw new Error("Falta reportType.");
    if (!startDate || !endDate) throw new Error("Faltan fechas.");
    if (String(startDate) > String(endDate)) throw new Error("Fecha inicio > fecha final.");

    if (!process.env.N8N_RUN_REPORT_URL) throw new Error("Falta N8N_RUN_REPORT_URL.");

    const payload = {
      realmId: String(realmId),
      reportType: String(reportType),
      startDate: String(startDate),
      endDate: String(endDate),
      excludeAccountIds: Array.isArray(excludeAccountIds) ? excludeAccountIds.map(String) : [],
      format: "xlsx",
    };

    const resp = await n8nPost(process.env.N8N_RUN_REPORT_URL, payload);

    if (!resp.ok) {
      const txt = await resp.text();
      throw new Error(`n8n error (${resp.status}): ${txt}`);
    }

    const buf = await resp.buffer();
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", `attachment; filename="QBO_${reportType}_${startDate}_${endDate}.xlsx"`);
    res.send(buf);
  } catch (e) {
    res.status(400).json({ error: e.message });
  }
});

const port = process.env.PORT || 3000;
app.listen(port, () => console.log("Listening on", port));

