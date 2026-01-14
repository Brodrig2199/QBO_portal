import os
import secrets
import base64
from datetime import datetime, timedelta, timezone
from functools import wraps

import requests
from flask import Flask, render_template, request, redirect, url_for, session, flash

from token_store import init_db, save_tokens
from qbo_client import (
    get_valid_access_token,
    get_customers,
    get_accounts,
    get_profit_and_loss,
    parse_pl_to_rows
)

app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_SECRET_KEY", "dev-secret-change-me")

# ✅ Esto SÍ corre en Render porque ocurre al importar el módulo
try:
    init_db()
except Exception as e:
    print("DB init skipped:", e)

LOGIN_USER = os.environ.get("LOGIN_USER", "admin")
LOGIN_PASS = os.environ.get("LOGIN_PASS", "admin123")

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
    datetime.strptime(date_str, "%Y-%m-%d")
    return date_str

def fetch_qbo_report(report_type: str, start_date: str, end_date: str, client_id: str, excluded_accounts: list[str]):
    # Solo implementamos P&L primero
    if report_type != "profit_and_loss":
        raise RuntimeError("Por ahora solo está habilitado Profit & Loss.")

    access_token, realm_id = get_valid_access_token()

    report_json = get_profit_and_loss(
        access_token=access_token,
        realm_id=realm_id,
        start_date=start_date,
        end_date=end_date,
        accounting_method="Accrual",
        summarize_column_by="Total",
        customer_id=client_id if client_id != "all" else None
    )

    rows = parse_pl_to_rows(report_json)

    # Excluir cuentas seleccionadas (por nombre)
    if excluded_accounts:
        excluded_set = set(excluded_accounts)
        rows = [r for r in rows if r["account"] not in excluded_set]

    return {
        "meta": {
            "report_type": report_type,
            "start_date": start_date,
            "end_date": end_date,
            "client_id": client_id,
            "excluded_accounts": excluded_accounts,
        },
        "rows": rows
    }


@app.get("/")
def home():
    return redirect(url_for("reports")) if session.get("logged_in") else redirect(url_for("login"))

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
    client_id = os.environ.get("QBO_CLIENT_ID", "")
    redirect_uri = os.environ.get("QBO_REDIRECT_URI", "")
    if not client_id or not redirect_uri:
        return "Faltan QBO_CLIENT_ID o QBO_REDIRECT_URI en env vars", 500

    authorize_url = "https://appcenter.intuit.com/connect/oauth2"
    scope = "com.intuit.quickbooks.accounting"

    state = secrets.token_urlsafe(24)
    session["oauth_state"] = state
    session["after_auth"] = request.args.get("next") or url_for("reports")

    from urllib.parse import urlencode
    params = {
        "client_id": client_id,
        "response_type": "code",
        "scope": scope,
        "redirect_uri": redirect_uri,
        "state": state,
    }
    return redirect(f"{authorize_url}?{urlencode(params)}")

@app.get("/callback")
def callback():
    code = request.args.get("code")
    realm_id = request.args.get("realmId")
    state = request.args.get("state")
    err = request.args.get("error")

    if err:
        return f"Autorización falló: {request.args.get('error_description', err)}", 400

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
    basic = base64.b64encode(f"{client_id}:{client_secret}".encode()).decode()

    headers = {
        "Authorization": f"Basic {basic}",
        "Accept": "application/json",
        "Content-Type": "application/x-www-form-urlencoded",
    }
    data = {"grant_type": "authorization_code", "code": code, "redirect_uri": redirect_uri}

    r = requests.post(token_url, headers=headers, data=data, timeout=30)
    if r.status_code >= 400:
        return f"Token exchange failed ({r.status_code}): {r.text}", 400

    payload = r.json()
    access_token = payload.get("access_token")
    refresh_token = payload.get("refresh_token")
    expires_in = int(payload.get("expires_in", 3600))

    expires_at = datetime.now(timezone.utc) + timedelta(seconds=expires_in)
    save_tokens(realm_id=realm_id, access_token=access_token, refresh_token=refresh_token, access_expires_at=expires_at)

    session.pop("oauth_state", None)
    flash("QuickBooks conectado ✅")
    return redirect(session.pop("after_auth", url_for("reports")))

@app.get("/reports")
@login_required
def reports():
    try:
        access_token, realm_id = get_valid_access_token()
        clients = [{"id":"all","name":"Todos los clientes"}] + get_customers(access_token, realm_id)
        accounts = get_accounts(access_token, realm_id)

        print("REPORTS OK -> clients:", len(clients), "accounts:", len(accounts))
        return render_template("reports.html", clients=clients, accounts=accounts, report_types=REPORT_TYPES)

    except Exception as e:
        print("REPORTS ERROR ->", repr(e))
        flash(f"QuickBooks no conectado o error: {e}. Ve a /connect.")
        return render_template("reports.html", clients=[{"id":"all","name":"Todos los clientes"}], accounts=[], report_types=REPORT_TYPES)


@app.post("/run-report")
@login_required
def run_report():
    report_type = request.form.get("report_type", "")
    start_date = parse_date(request.form.get("start_date", ""))
    end_date = parse_date(request.form.get("end_date", ""))
    client_id = request.form.get("client_id", "all")
    excluded_accounts = request.form.getlist("excluded_accounts")

    data = fetch_qbo_report(report_type, start_date, end_date, client_id, excluded_accounts)
    return render_template("results.html", data=data)
