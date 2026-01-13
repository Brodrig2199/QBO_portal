const express = require("express");
const session = require("express-session");
const fetch = require("node-fetch");
require("dotenv").config();

const app = express();
app.set("view engine", "ejs");
app.use(express.urlencoded({ extended: true }));
app.use(express.json());
app.set("trust proxy", 1);

app.use(session({
  secret: process.env.SESSION_SECRET || "secret",
  resave: false,
  saveUninitialized: false,
  cookie: {
    httpOnly: true,
    sameSite: "lax",
    secure: process.env.NODE_ENV === "production"
  }
}));

function auth(req, res, next) {
  if (req.session.user) return next();
  return res.redirect("/login");
}

// TODO: puedes mover esto a DB luego
const COMPANIES = [
  { clientId: "cli_001", name: "Empresa A", realmId: "12314567890" },
  { clientId: "cli_002", name: "Empresa B", realmId: "09876543210" }
];

const REPORTS = [
  { id: "ProfitAndLoss", name: "Profit & Loss" },
  { id: "BalanceSheet", name: "Balance Sheet" },
  { id: "TrialBalance", name: "Trial Balance" },
  { id: "GeneralLedger", name: "General Ledger" }
];

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

app.get("/", auth, (req, res) => {
  res.render("form", { user: req.session.user, companies: COMPANIES, reports: REPORTS });
});

// El browser pega a tu backend; tu backend llama a n8n
app.post("/api/run-report", auth, async (req, res) => {
  try {
    const { companyClientId, reportType, startDate, endDate, excludeAccounts } = req.body;

    const company = COMPANIES.find(c => c.clientId === companyClientId);
    if (!company) throw new Error("Empresa inválida");

    if (!reportType || !startDate || !endDate) throw new Error("Faltan campos");
    if (startDate > endDate) throw new Error("Fecha inicio no puede ser mayor a fecha final");

    const payload = {
      company: { clientId: company.clientId, realmId: company.realmId },
      reportType,
      startDate,
      endDate,
      excludeAccounts: Array.isArray(excludeAccounts) ? excludeAccounts : []
    };

    if (!process.env.N8N_WEBHOOK_URL) throw new Error("Falta N8N_WEBHOOK_URL");
    if (!process.env.ALIADA_WEBHOOK_KEY) throw new Error("Falta ALIADA_WEBHOOK_KEY");

    const resp = await fetch(process.env.N8N_WEBHOOK_URL, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-aliada-key": process.env.ALIADA_WEBHOOK_KEY
      },
      body: JSON.stringify(payload)
    });

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

app.listen(process.env.PORT || 3000, () => console.log("Running"));

