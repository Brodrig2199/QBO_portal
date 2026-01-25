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
    parse_report_to_table,
    get_vendors,
    get_vendor,
    extract_vendor_otros
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
                         "accounting_method": "Accrual",
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

    # 🔹 Re-descargar el reporte original de QuickBooks (preview completo)
    if meta["report_type"] == "profit_and_loss_detail":
        report_json = get_profit_and_loss_detail(
            access_token=access_token,
            realm_id=realm_id,
            start_date=meta["start_date"],
            end_date=meta["end_date"],
            accounting_method="Accrual",
            customer_id=None if meta.get("client_id") in (None, "", "all") else meta["client_id"],
        )
        sheet_title = "Profit & Loss Detail"
        filename = f"QBO_ProfitAndLossDetail_{meta['start_date']}_{meta['end_date']}.xlsx"

    elif meta["report_type"] == "vat_tax_detail":
        report_json = get_vat_tax_detail(
            access_token=access_token,
            realm_id=realm_id,
            start_date=meta["start_date"],
            end_date=meta["end_date"],
        )
        sheet_title = "VAT Tax Detail"
        filename = f"QBO_TaxDetail_{meta['start_date']}_{meta['end_date']}.xlsx"

    else:
        flash("Tipo de reporte no soportado.")
        return redirect(url_for("reports"))

    table = parse_report_to_table(report_json)

    # --- Crear Excel genérico (tal cual QuickBooks) ---
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment
    from openpyxl.utils import get_column_letter
    import io

    wb = Workbook()
    ws = wb.active
    ws.title = sheet_title

    # headers
    ws.append(table["columns"])
    for c in range(1, len(table["columns"]) + 1):
        ws.cell(row=1, column=c).font = Font(bold=True)

    # rows
    for r in table["rows"]:
        ws.append(r["cells"])

    for i in range(1, ws.max_column + 1):
        ws.column_dimensions[get_column_letter(i)].width = 22

    stream = io.BytesIO()
    wb.save(stream)
    stream.seek(0)

    return send_file(
        stream,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
@app.get("/download/informe43.xlsx")
@login_required
def download_informe43_xlsx():
    meta = session.get("last_report_meta")
    if not meta:
        flash("No hay parámetros del reporte. Genera uno primero.")
        return redirect(url_for("reports"))

    # INFORME 43 basado en P&L DETAIL
    if meta.get("report_type") != "profit_and_loss_detail":
        flash("El INFORME 43 se genera desde Detalle de Pérdidas y Ganancias.")
        return redirect(url_for("reports"))

    access_token, realm_id = get_valid_access_token()

    report_json = get_profit_and_loss_detail(
        access_token=access_token,
        realm_id=realm_id,
        start_date=meta["start_date"],
        end_date=meta["end_date"],
        accounting_method="Accrual",
        customer_id=None if meta.get("client_id") in (None, "", "all") else meta["client_id"],
    )

    table = parse_report_to_table(report_json)

    # -------------------------
    # Helpers
    # -------------------------
    import re, io
    from datetime import datetime
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
    from openpyxl.utils import get_column_letter

    cols = [(c or "").strip().lower() for c in table.get("columns", [])]

    def find_col_contains(*keys):
        for k in keys:
            kk = (k or "").strip().lower()
            for i, c in enumerate(cols):
                if kk in c:
                    return i
        return None

    def cell(row, idx):
        if idx is None:
            return ""
        cells = row.get("cells") or []
        if idx < 0 or idx >= len(cells):
            return ""
        return (cells[idx] or "").strip()

    def to_float(x):
        try:
            s = str(x).replace(",", "").strip()
            return float(s) if s else 0.0
        except:
            return 0.0

    def to_yyyymmdd(s):
        s = (s or "").strip()
        if not s:
            return ""
        for f in ("%Y-%m-%d", "%m/%d/%Y", "%d/%m/%Y"):
            try:
                return datetime.strptime(s, f).strftime("%Y%m%d")
            except:
                pass
        # si ya viene como 20251205
        if len(s) == 8 and s.isdigit():
            return s
        return ""

    def parse_vendor(name):
        """
        Soporta:
        1) NOMBRE/TIPO/RUC/DV  Ej: BANCO GENERAL/2/280-134-61098/2
        2) NOMBRE/TIPO         Ej: AMAZON/3
        """
        raw = (name or "").strip()
        if not raw:
            return ("", "", "", "")

        tipo_map = {"1": "N", "2": "J", "3": "E"}

        m = re.match(r'^\s*(.+?)\s*/\s*([123])\s*/\s*([^/]+)\s*/\s*([^/]+)\s*$', raw)
        if m:
            return (
                tipo_map.get(m.group(2).strip(), ""),
                m.group(3).strip(),
                m.group(4).strip(),
                m.group(1).strip()
            )

        m2 = re.match(r'^\s*(.+?)\s*/\s*([123])\s*$', raw)
        if m2:
            return (
                tipo_map.get(m2.group(2).strip(), ""),
                "",
                "",
                m2.group(1).strip()
            )

        return ("", "", "", raw.replace("/", " ").strip())

    def parse_otros(otros_raw: str):
        """
        OTROS esperado: "CONCEPTO/COMPRAS"
        Ej: "1/505" => concepto=1, compras=505
        """
        s = (otros_raw or "").strip()
        if not s:
            return ("", "")
        parts = [p.strip() for p in s.split("/") if p.strip()]
        concepto = parts[0] if len(parts) >= 1 else ""
        compras = parts[1] if len(parts) >= 2 else ""
        return (concepto, compras)

    def is_panama_cedula(ruc: str) -> bool:
        s = (ruc or "").strip().upper()
        return bool(re.match(r'^\d{1,2}-\d{1,6}-\d{1,6}$', s))

    def infer_tipo_persona(tipo_from_name: str, ruc: str) -> str:
        t = (tipo_from_name or "").strip().upper()
        if t in ("N", "J", "E"):
            return t

        r = (ruc or "").strip().upper()
        if is_panama_cedula(r):
            return "N"
        if r.startswith("E") or "PASAPORTE" in r or "PASS" in r:
            return "E"

        digits = re.sub(r"\D", "", r)
        if len(digits) >= 10:
            return "J"
        return ""

    def normalize_factura(factura_raw: str, seq_num: int) -> str:
        f = (factura_raw or "").strip()
        return f if f else f"F-{seq_num}"

        # -------------------------
    # ✅ MAPA: nombre_proveedor -> "Otro"
    # -------------------------
    vendors = get_vendors(access_token, realm_id, active_only=True)
    vendor_id_by_name = { (v["name"] or "").strip().lower(): v["id"] for v in vendors if v.get("name") and v.get("id") }

    vendor_otros_cache = {}  # name_lower -> "2/1"

    def get_otros_for_vendor_name(vendor_name_clean: str) -> str:
        """
        vendor_name_clean = nombre ya limpio (sin /tipo/ruc/dv)
        """
        key = (vendor_name_clean or "").strip().lower()
        if not key:
            return ""
        if key in vendor_otros_cache:
            return vendor_otros_cache[key]

        vid = vendor_id_by_name.get(key)
        if not vid:
            vendor_otros_cache[key] = ""
            return ""

        try:
            payload = get_vendor(access_token, realm_id, vid)
            otros = extract_vendor_otros(payload)  # "2/1"
            vendor_otros_cache[key] = otros
            return otros
        except Exception:
            vendor_otros_cache[key] = ""
            return ""


    # -------------------------
    # MAP COLUMNAS (P&L DETAIL)
    # -------------------------
    idx_fecha   = find_col_contains("fecha", "date")
    idx_no      = find_col_contains("n.", "no", "nº", "numero")
    idx_nombre  = find_col_contains("nombre", "name")
    idx_importe = find_col_contains("importe", "amount")
    idx_otros   = find_col_contains("otros", "other", "notas", "descripción", "descripcion", "notas/descripcion", "notes")

    # Opcional: Cuenta contable si quieres usarla luego
    idx_cuenta_div = find_col_contains("cuenta de división", "cuenta de division", "dividir", "split")


    # -------------------------
    # CONSTRUIR FILAS INFORME 43 (DESDE P&L)
    # -------------------------
    rows_out = []
    seq = 1

    for r in (table.get("rows") or []):
        if r.get("is_header") or r.get("is_summary"):
            continue

        nombre_raw = cell(r, idx_nombre)
        if not nombre_raw:
            continue

        tipo_from_name, ruc_from_name, dv, nombre = parse_vendor(nombre_raw)

        # si el nombre no trae ruc, quedará vacío (ok). Igual inferimos tipo si se puede.
        ruc = ruc_from_name or ""
        tipo = infer_tipo_persona(tipo_from_name, ruc)

        # Concepto/Compras desde OTROS
        otros_raw = get_otros_for_vendor_name(nombre)
        concepto, compras = parse_otros(otros_raw)

        factura = normalize_factura(cell(r, idx_no), seq)
        seq += 1

        monto = to_float(cell(r, idx_importe))

        rows_out.append([
            tipo,                              # TIPO
            ruc,                               # RUC
            dv,                                # DV
            nombre,                            # NOMBRE
            factura,                           # FACTURA (F-n)
            to_yyyymmdd(cell(r, idx_fecha)),   # FECHA yyyymmdd
            concepto,                          # CONCEPTO (de OTROS)
            compras,                           # COMPRAS (de OTROS)
            monto,                             # MONTO EN BALBOAS (importe)
            0.00                               # ITBMS PAGADO (cero)
        ])




    # -------------------------
    # Crear Excel
    # -------------------------
    wb = Workbook()
    ws = wb.active
    ws.title = "INFORME 43 (VAT)"

    headers = [
        "TIPO DE PERSONA",
        "RUC",
        "DV",
        "NOMBRE O RAZON SOCIAL",
        "FACTURA",
        "FECHA",
        "CONCEPTO",
        "COMPRAS DE BIENES Y SERVICIOS",
        "MONTO EN BALBOAS",
        "ITBMS PAGADO EN BALBOAS",
    ]

    bold = Font(bold=True)
    fill = PatternFill("solid", fgColor="EFEFEF")
    thin = Side(style="thin")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(headers))
    ws["A1"] = "INFORME 43 - FORMATO A DILIGENCIAR (VAT)"
    ws["A1"].font = Font(bold=True, size=13)

    ws.append([])
    ws.append([])

    for i, h in enumerate(headers, start=1):
        c = ws.cell(row=5, column=i, value=h)
        c.font = bold
        c.fill = fill
        c.border = border
        c.alignment = Alignment(horizontal="center", wrap_text=True)

    for rr, rowvals in enumerate(rows_out, start=6):
        for cc, val in enumerate(rowvals, start=1):
            cellx = ws.cell(row=rr, column=cc, value=val)
            cellx.border = border

            # Texto para RUC y DV
            if cc in (2, 3):
                cellx.number_format = '@'

            # Formato moneda para MONTO e ITBMS
            if cc in (9, 10):
                cellx.number_format = '#,##0.00'

    widths = [16, 18, 6, 35, 14, 12, 18, 30, 18, 22]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.freeze_panes = "A6"

    stream = io.BytesIO()
    wb.save(stream)
    stream.seek(0)

    return send_file(
        stream,
        as_attachment=True,
        download_name=f"INFORME43_VAT_{meta['start_date']}_{meta['end_date']}.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
