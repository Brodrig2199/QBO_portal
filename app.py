import os
from datetime import datetime
from functools import wraps

from flask import Flask, render_template, request, redirect, url_for, session, flash

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


@app.get("/reports")
@login_required
def reports():
    # Defaults: hoy y hace 30 días (opcional)
    return render_template(
        "reports.html",
        clients=CLIENTS,
        report_types=REPORT_TYPES
    )


@app.post("/run-report")
@login_required
def run_report():
    try:
        report_type = request.form.get("report_type", "")
        start_date = parse_date(request.form.get("start_date", ""))
        end_date = parse_date(request.form.get("end_date", ""))
        client_id = request.form.get("client_id", "all")
        excluded_accounts = parse_excluded_accounts(request.form.get("excluded_accounts", ""))

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
