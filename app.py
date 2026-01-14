import os
from datetime import datetime
from functools import wraps
from token_store import init_db
import secrets
import base64
import requests
from datetime import datetime, timedelta, timezone
from token_store import init_db, save_tokens
from flask import Flask, render_template, request, redirect, url_for, session, flash
from qbo_client import get_valid_access_token, get_customers, get_accounts


app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_SECRET_KEY", "dev-secret-change-me")

# Credenciales del login (ponlas en Render como Environment Variables)
LOGIN_USER = os.environ.get("LOGIN_USER", "admin")
LOGIN_PASS = os.environ.get("LOGIN_PASS", "admin123")

# Clientes (por ahora estático; luego lo puedes traer desde QuickBooks Customers)
CLIENTS = [
    {"id": "all", "name": "Todos los clientes"},
    {"id": "c001", "name": "Cliente A"},
    {"id": "c002", "name": "Cliente B"},
    {"id": "c003", "name": "Cliente C"},
]

REPORT_TYPES = [
    {"id": "profit_and_loss", "name": "Profit & Loss (Pérdidas y Ganancias)"},
    {"id": "balance_sheet", "name": "Balance Sheet (Balance General)"},
    {"id": "trial_balance", "name": "Trial Balance (Balanza de Comprobación)"},
    {"id": "general_ledger", "name": "General Ledger (Libro Mayor)"},
]


def login_required(f):
    @wraps(f)
    def wrapper(*args, **kwargs):
        if not session.get("logged_in"):
            return redirect(url_for("login"))
        return f(*args, **kwargs)
    return wrapper


def parse_date(date_str: str) -> str:
    """
    Convierte yyyy-mm-dd a fecha válida y devuelve el mismo string (para usarlo tal cual).
    Lanza ValueError si no es válido.
    """
    datetime.strptime(date_str, "%Y-%m-%d")
    return date_str


def parse_excluded_accounts(raw: str):
    """
    Recibe: "5100, 5200,Utilities, Bank Fees"
    Devuelve lista limpia: ["5100", "5200", "Utilities", "Bank Fees"]
    """
    if not raw:
        return []
    items = [x.strip() for x in raw.split(",")]
    return [x for x in items if x]


def fetch_qbo_report(report_type: str, start_date: str, end_date: str, client_id: str, excluded_accounts: list[str]):
    """
    Aquí conectarás QuickBooks.
    Por ahora devuelve datos demo para validar la web y el submit.
    """
    # TODO: Integrar QBO API / Reports / Query
    demo_rows = [
        {"account": "Sales", "amount": 2500.00},
        {"account": "Cost of Goods Sold", "amount": -900.00},
        {"account": "Utilities", "amount": -120.50},
        {"account": "Bank Fees", "amount": -15.00},
    ]

    if excluded_accounts:
        demo_rows = [
            r for r in demo_rows
            if r["account"] not in excluded_accounts and str(r["account"]) not in excluded_accounts
        ]

    return {
        "meta": {
            "report_type": report_type,
            "start_date": start_date,
            "end_date": end_date,
            "client_id": client_id,
            "excluded_accounts": excluded_accounts,
        },
        "rows": demo_rows
    }


@app.get("/")
def home():
    if session.get("logged_in"):
        return redirect(url_for("reports"))
    return redirect(url_for("login"))


@app.get("/login")
def login():
    return render_template("login.html")


@app.post("/login")
def login_post():
    username = request.form.get("username", "")
    password = request.form.get("password", "")

    if username == LOGIN_USER and password == LOGIN_PASS:
        session["logged_in"] = True
        session["username"] = username
        return redirect(url_for("reports"))

    flash("Usuario o contraseña incorrectos.")
    return redirect(url_for("login"))


@app.get("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))
@app.get("/connect")
def connect():
    """
    Redirige al usuario a Intuit para autorizar la app.
    """
    client_id = os.environ.get("QBO_CLIENT_ID", "")
    redirect_uri = os.environ.get("QBO_REDIRECT_URI", "")
    if not client_id or not redirect_uri:
        return "Faltan QBO_CLIENT_ID o QBO_REDIRECT_URI en env vars", 500

    # Sandbox y Production usan el mismo authorize host
    authorize_url = "https://appcenter.intuit.com/connect/oauth2"

    # Scope: contabilidad (para Customers, Accounts, Reports)
    scope = "com.intuit.quickbooks.accounting"

    # State: anti-CSRF
    state = secrets.token_urlsafe(24)
    session["oauth_state"] = state

    # (Opcional) guarda “a dónde volver” después
    session["after_auth"] = request.args.get("next") or url_for("reports")

    params = {
        "client_id": client_id,
        "response_type": "code",
        "scope": scope,
        "redirect_uri": redirect_uri,
        "state": state,
    }

    # Construir URL con querystring
    from urllib.parse import urlencode
    return redirect(f"{authorize_url}?{urlencode(params)}")

@app.get("/callback")
def callback():
    """
    Recibe code + realmId desde Intuit, intercambia por tokens, y guarda en DB.
    """
    code = request.args.get("code")
    realm_id = request.args.get("realmId")
    state = request.args.get("state")
    err = request.args.get("error")

    if err:
        desc = request.args.get("error_description", err)
        return f"Autorización falló: {desc}", 400

    # Validar state anti-CSRF
    saved_state = session.get("oauth_state")
    if not saved_state or state != saved_state:
        return "State inválido. Reintenta /connect.", 400

    if not code or not realm_id:
        return "Faltan parámetros code o realmId.", 400

    client_id = os.environ.get("QBO_CLIENT_ID", "")
    client_secret = os.environ.get("QBO_CLIENT_SECRET", "")
    redirect_uri = os.environ.get("QBO_REDIRECT_URI", "")
    if not client_id or not client_secret or not redirect_uri:
        return "Faltan env vars QBO_CLIENT_ID/SECRET/REDIRECT_URI", 500

    token_url = "https://oauth.platform.intuit.com/oauth2/v1/tokens/bearer"

    # Basic Auth = base64(client_id:client_secret)
    basic = base64.b64encode(f"{client_id}:{client_secret}".encode()).decode()

    headers = {
        "Authorization": f"Basic {basic}",
        "Accept": "application/json",
        "Content-Type": "application/x-www-form-urlencoded",
    }

    data = {
        "grant_type": "authorization_code",
        "code": code,
        "redirect_uri": redirect_uri,
    }

    r = requests.post(token_url, headers=headers, data=data, timeout=30)
    if r.status_code >= 400:
        return f"Token exchange failed ({r.status_code}): {r.text}", 400

    payload = r.json()
    access_token = payload.get("access_token")
    refresh_token = payload.get("refresh_token")
    expires_in = int(payload.get("expires_in", 3600))

    if not access_token or not refresh_token:
        return f"Respuesta incompleta: {payload}", 400

    expires_at = datetime.now(timezone.utc) + timedelta(seconds=expires_in)

    # Guardar tokens en DB (fuente de verdad)
    save_tokens(realm_id=realm_id, access_token=access_token, refresh_token=refresh_token, access_expires_at=expires_at)

    # Limpia state
    session.pop("oauth_state", None)

    flash("QuickBooks conectado ✅")
    return redirect(session.pop("after_auth", url_for("reports")))



@app.get("/reports")
@login_required
def reports():
    try:
        access_token, realm_id = get_valid_access_token()
        clients = [{"id": "all", "name": "Todos los clientes"}] + get_customers(access_token, realm_id)
        accounts = get_accounts(access_token, realm_id)

        return render_template("reports.html", clients=clients, accounts=accounts, report_types=REPORT_TYPES)
    except Exception as e:
        flash(f"QuickBooks no conectado o error: {e}. Ve a /connect.")
        return render_template("reports.html", clients=[{"id":"all","name":"Todos los clientes"}], accounts=[], report_types=REPORT_TYPES)



@app.post("/run-report")
@login_required
def run_report():
    try:
        report_type = request.form.get("report_type", "")
        start_date = parse_date(request.form.get("start_date", ""))
        end_date = parse_date(request.form.get("end_date", ""))
        client_id = request.form.get("client_id", "all")
        excluded_accounts = request.form.getlist("excluded_accounts")


        if report_type not in [r["id"] for r in REPORT_TYPES]:
            flash("Tipo de reporte inválido.")
            return redirect(url_for("reports"))

        data = fetch_qbo_report(
            report_type=report_type,
            start_date=start_date,
            end_date=end_date,
            client_id=client_id,
            excluded_accounts=excluded_accounts
        )

        # Mostrar resultados en pantalla (luego puedes permitir descargar CSV/PDF)
        return render_template("results.html", data=data, clients=CLIENTS, report_types=REPORT_TYPES)

    except ValueError:
        flash("Fechas inválidas. Usa el selector de fecha.")
        return redirect(url_for("reports"))
    except Exception as e:
        flash(f"Error ejecutando el reporte: {e}")
        return redirect(url_for("reports"))


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", "5000")))
    init_db()
