// ====== CONFIG ======
// 1) Pon aquí tu webhook de n8n (producción)
const N8N_WEBHOOK_URL = "https://TU-N8N/webhook/qbo-report/run";

// 2) Un secreto que también validas en n8n (header x-aliada-key)
const ALIADA_WEBHOOK_KEY = "CAMBIA_ESTE_SECRETO";

// 3) “Login” básico (solo UI). NO es seguridad real.
const DEMO_USER = "admin";
const DEMO_PASS = "admin123";

// 4) Lista de empresas (tú pones realmId por empresa)
const COMPANIES = [
  { clientId: "cli_001", name: "Empresa A", realmId: "12314567890" },
  { clientId: "cli_002", name: "Empresa B", realmId: "09876543210" },
];

// ====== LOGIN UI ======
const loginPage = document.getElementById("loginPage");
const formPage = document.getElementById("formPage");
const loginErr = document.getElementById("loginErr");

document.getElementById("btnLogin").addEventListener("click", () => {
  const u = document.getElementById("user").value.trim();
  const p = document.getElementById("pass").value;

  if (u === DEMO_USER && p === DEMO_PASS) {
    localStorage.setItem("aliada_logged", "1");
    loginErr.textContent = "";
    showForm();
  } else {
    loginErr.textContent = "Usuario o contraseña incorrectos.";
  }
});

document.getElementById("btnLogout").addEventListener("click", () => {
  localStorage.removeItem("aliada_logged");
  showLogin();
});

function showLogin() {
  loginPage.classList.remove("hidden");
  formPage.classList.add("hidden");
}
function showForm() {
  loginPage.classList.add("hidden");
  formPage.classList.remove("hidden");
}

// ====== Populate companies ======
const companySelect = document.getElementById("company");
function loadCompanies() {
  companySelect.innerHTML = "";
  for (const c of COMPANIES) {
    const opt = document.createElement("option");
    opt.value = c.clientId;
    opt.textContent = c.name;
    companySelect.appendChild(opt);
  }
}
loadCompanies();

// ====== FORM SUBMIT ======
const form = document.getElementById("reportForm");
const err = document.getElementById("err");
const statusEl = document.getElementById("status");
const btnSubmit = document.getElementById("btnSubmit");

form.addEventListener("submit", async (e) => {
  e.preventDefault();
  err.textContent = "";
  statusEl.textContent = "";

  const companyClientId = companySelect.value;
  const company = COMPANIES.find(c => c.clientId === companyClientId);

  const reportType = document.getElementById("reportType").value;
  const startDate = document.getElementById("startDate").value;
  const endDate = document.getElementById("endDate").value;
  const excludeAccountsRaw = document.getElementById("excludeAccounts").value;

  if (!company) return (err.textContent = "Empresa inválida.");
  if (!startDate || !endDate) return (err.textContent = "Selecciona fecha inicio y final.");
  if (startDate > endDate) return (err.textContent = "La fecha inicio no puede ser mayor a la final.");

  const excludeAccounts = (excludeAccountsRaw || "")
    .split(",")
    .map(s => s.trim())
    .filter(Boolean);

  const payload = {
    company: { clientId: company.clientId, realmId: company.realmId },
    reportType,
    startDate,
    endDate,
    excludeAccounts,
    format: "xlsx"
  };

  try {
    btnSubmit.disabled = true;
    statusEl.textContent = "Generando…";

    const resp = await fetch(N8N_WEBHOOK_URL, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-aliada-key": ALIADA_WEBHOOK_KEY
      },
      body: JSON.stringify(payload)
    });

    if (!resp.ok) {
      const txt = await resp.text();
      throw new Error(`n8n error (${resp.status}): ${txt}`);
    }

    // Descargar XLSX (binary)
    const blob = await resp.blob();
    const filename = `QBO_${reportType}_${startDate}_${endDate}.xlsx`;

    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(a.href);

    statusEl.textContent = "Listo ✅";
  } catch (e2) {
    err.textContent = e2.message || "Error generando reporte.";
  } finally {
    btnSubmit.disabled = false;
  }
});

// ====== First load ======
if (localStorage.getItem("aliada_logged") === "1") showForm();
else showLogin();
