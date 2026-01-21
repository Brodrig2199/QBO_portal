import os
import secrets
import base64
import io
from datetime import datetime, timedelta, timezone
from functools import wraps

import requests
from flask import Flask, render_template, request, redirect, url_for, session, flash, send_file

from token_store import init_db, save_tokens
from qbo_client import (
    get_valid_access_token,
    get_customers,
    get_accounts,
    get_profit_and_loss_detail,
    get_vat_tax_detail,
    parse_report_to_table
)

app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_SECRET_KEY", "dev-secret-change-me")

try:
    init_db()
except Exception as e:
    print("DB init skipped:", e)

LOGIN_USER = os.environ.get("LOGIN_USER", "admin")
LOGIN_PASS = os.environ.get("LOGIN_PASS", "admin123")

REPORT_TYPES = [
    {"id": "profit_and_loss_detail", "name": "Detalle de Pérdidas y Ganancias", "qbo": "ProfitAndLossDetail"},
    {"id": "vat_tax_detail", "name": "VAT - Detalle de Impuestos", "qbo": "TaxDetail"},
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
    access_token, realm_id = get_valid_access_token()

    if report_type == "profit_and_loss_detail":
        report_json = get_profit_and_loss_detail(
            access_token=access_token,
            realm_id=realm_id,
            start_date=start_date,
            end_date=end_date,
            accounting_method="Accrual",
            customer_id=None if client_id == "all" else client_id
        )
        table = parse_report_to_table(report_json)

        return {"meta": {"report_type": report_type, "qbo_report_name": "ProfitAndLossDetail",
                         "start_date": start_date, "end_date": end_date, "client_id": client_id,
                         "accounting_method": "Accrual", "summarize_column_by": "Total",
                         "excluded_accounts": excluded_accounts},
                "table": table, "raw": report_json}

    if report_type == "vat_tax_detail":
        report_json = get_vat_tax_detail(access_token, realm_id, start_date, end_date)
        table = parse_report_to_table(report_json)

        return {"meta": {"report_type": report_type, "qbo_report_name": "TaxDetail",
                         "start_date": start_date, "end_date": end_date, "client_id": client_id,
                         "excluded_accounts": excluded_accounts},
                "table": table, "raw": report_json}

    raise RuntimeError(f"Tipo de reporte inválido: {report_type}")


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
        clients = [{"id": "all", "name": "Todos los clientes"}] + get_customers(access_token, realm_id)
        accounts = get_accounts(access_token, realm_id)
        return render_template("reports.html", clients=clients, accounts=accounts, report_types=REPORT_TYPES)
    except Exception as e:
        print("REPORTS ERROR ->", repr(e))
        flash(f"QuickBooks no conectado o error: {e}. Ve a /connect.")
        return render_template("reports.html", clients=[{"id": "all", "name": "Todos los clientes"}], accounts=[], report_types=REPORT_TYPES)


@app.post("/run-report")
@login_required
def run_report():
    try:
        report_type = request.form.get("report_type", "")
        start_date = parse_date(request.form.get("start_date", ""))
        end_date = parse_date(request.form.get("end_date", ""))
        client_id = request.form.get("client_id", "all")
        excluded_accounts = request.form.getlist("excluded_accounts")

        print("RUN REPORT -> report_type:", report_type, "start:", start_date, "end:", end_date, "client:", client_id)

        data = fetch_qbo_report(report_type, start_date, end_date, client_id, excluded_accounts)

        # Guardar meta para download
        session["last_report_meta"] = data["meta"]

        return render_template("results.html", data=data)

    except Exception as e:
        print("RUN REPORT ERROR ->", repr(e))
        flash(f"Error generando reporte: {e}")
        return redirect(url_for("reports"))


@app.get("/download/qbo/report.xlsx")
@login_required
def download_qbo_report_xlsx():
    meta = session.get("last_report_meta")
    if not meta:
        flash("No hay parámetros del reporte. Genera uno primero.")
        return redirect(url_for("reports"))

    access_token, realm_id = get_valid_access_token()

    # Volvemos a traer el reporte (para descargar siempre lo más actualizado)
    rt = meta.get("report_type")
    if rt == "profit_and_loss_detail":
        report_json = get_profit_and_loss_detail(
            access_token=access_token,
            realm_id=realm_id,
            start_date=meta["start_date"],
            end_date=meta["end_date"],
            accounting_method=meta.get("accounting_method", "Accrual"),
            summarize_column_by=meta.get("summarize_column_by", "Total"),
            customer_id=None if meta.get("client_id") in (None, "", "all") else meta["client_id"],
        )
        sheet_title = "P&L Detail"
        filename = f"QBO_ProfitAndLossDetail_{meta['start_date']}_{meta['end_date']}.xlsx"

    elif rt == "vat_tax_detail":
        report_json = get_vat_tax_detail(
            access_token=access_token,
            realm_id=realm_id,
            start_date=meta["start_date"],
            end_date=meta["end_date"],
        )
        sheet_title = "VAT Tax Detail"
        filename = f"QBO_TaxDetail_{meta['start_date']}_{meta['end_date']}.xlsx"

    else:
        flash("Reporte inválido en sesión.")
        return redirect(url_for("reports"))

    # Convertimos a tabla genérica
    table = parse_report_to_table(report_json)

    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    ws = wb.active
    ws.title = sheet_title

    header = report_json.get("Header", {}) or {}
    report_name = header.get("ReportName") or meta.get("qbo_report_name") or "QuickBooks Report"

    # Título
    ws["A1"] = report_name
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = f"Periodo: {header.get('StartPeriod', meta['start_date'])} a {header.get('EndPeriod', meta['end_date'])}"
    ws["A2"].font = Font(size=11)
    ws.append([])

    # Encabezados
    col_titles = table["columns"] if table["columns"] else ["Column"]
    col_types = table.get("col_types", []) or [""] * len(col_titles)

    start_row = ws.max_row + 1
    ws.append(col_titles)
    for j in range(1, len(col_titles) + 1):
        c = ws.cell(row=start_row, column=j)
        c.font = Font(bold=True)
        c.alignment = Alignment(horizontal="left")

    # Money cols
    money_cols = [i for i, t in enumerate(col_types, start=1) if (t or "").lower() == "money"]

    def safe_float(x):
        if x is None:
            return 0.0
        if isinstance(x, (int, float)):
            return float(x)
        s = str(x).replace(",", "").strip()
        if s in ("", "-"):
            return 0.0
        try:
            return float(s)
        except:
            return 0.0

    def write_row(cells, level=0, bold=False):
        # asegurar longitud exacta
        vals = list(cells) + [""] * (len(col_titles) - len(cells))
        vals = vals[:len(col_titles)]

        r = ws.max_row + 1
        ws.append(vals)

        # indent primera columna
        ws.cell(row=r, column=1).alignment = Alignment(indent=level, horizontal="left")

        if bold:
            for j in range(1, len(col_titles) + 1):
                ws.cell(row=r, column=j).font = Font(bold=True)

        # money formatting
        for mc in money_cols:
            cell = ws.cell(row=r, column=mc)
            if isinstance(cell.value, (int, float)):
                cell.number_format = '#,##0.00'

    # escribir data del reporte “tal cual” (Header/Data/Summary con indent)
    for row in table["rows"]:
        vals = list(row["cells"])

        # convertir money cols a float
        for idx in money_cols:
            if idx - 1 < len(vals):
                vals[idx - 1] = safe_float(vals[idx - 1])

        write_row(
            vals,
            level=int(row.get("level", 0)),
            bold=bool(row.get("is_header")) or bool(row.get("is_summary"))
        )

    # Auto width
    for col in range(1, ws.max_column + 1):
        max_len = 0
        for cell in ws[get_column_letter(col)]:
            val = "" if cell.value is None else str(cell.value)
            if len(val) > max_len:
                max_len = len(val)
        ws.column_dimensions[get_column_letter(col)].width = min(max_len + 2, 60)

    stream = io.BytesIO()
    wb.save(stream)
    stream.seek(0)

    return send_file(
        stream,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
