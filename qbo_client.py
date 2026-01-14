import os
import base64
from datetime import datetime, timedelta, timezone
from urllib.parse import urlencode

import requests

from token_store import get_tokens, save_tokens, is_access_token_valid

QBO_ENV = os.environ.get("QBO_ENV", "sandbox").lower()
QBO_CLIENT_ID = os.environ.get("QBO_CLIENT_ID", "")
QBO_CLIENT_SECRET = os.environ.get("QBO_CLIENT_SECRET", "")
QBO_REDIRECT_URI = os.environ.get("QBO_REDIRECT_URI", "")
QBO_MINORVERSION = os.environ.get("QBO_MINORVERSION", "75")

TOKEN_URL = "https://oauth.platform.intuit.com/oauth2/v1/tokens/bearer"

def _api_base() -> str:
    return "https://quickbooks.api.intuit.com" if QBO_ENV == "production" else "https://sandbox-quickbooks.api.intuit.com"

def _basic_auth_header() -> str:
    raw = f"{QBO_CLIENT_ID}:{QBO_CLIENT_SECRET}".encode("utf-8")
    return base64.b64encode(raw).decode("utf-8")

def get_valid_access_token() -> tuple[str, str]:
    """
    Devuelve (access_token, realm_id) usando DB como fuente de verdad.
    - Si access_token vigente -> lo usa
    - Si expiró -> refresca con refresh_token, guarda el refresh_token nuevo (rotación)
    """
    row = get_tokens() or {}
    realm_id = row.get("realm_id")

    # 1) Si el access token todavía sirve
    if row.get("access_token") and row.get("access_expires_at") and is_access_token_valid(row["access_expires_at"]):
        if not realm_id:
            raise RuntimeError("No hay realm_id guardado. Conecta QuickBooks en /connect.")
        return row["access_token"], realm_id

    # 2) Refrescar tokens
    refresh_token = row.get("refresh_token")
    if not refresh_token:
        raise RuntimeError("No hay refresh_token guardado. Conecta QuickBooks en /connect.")

    if not QBO_CLIENT_ID or not QBO_CLIENT_SECRET:
        raise RuntimeError("Faltan QBO_CLIENT_ID / QBO_CLIENT_SECRET en env vars.")

    headers = {
        "Authorization": f"Basic {_basic_auth_header()}",
        "Accept": "application/json",
        "Content-Type": "application/x-www-form-urlencoded",
    }
    data = {
        "grant_type": "refresh_token",
        "refresh_token": refresh_token,
    }

    r = requests.post(TOKEN_URL, headers=headers, data=data, timeout=30)
    if r.status_code >= 400:
        raise RuntimeError(f"Token refresh failed ({r.status_code}): {r.text}")

    payload = r.json()
    access_token = payload.get("access_token")
    new_refresh = payload.get("refresh_token", refresh_token)  # Intuit rota, guarda el nuevo
    expires_in = int(payload.get("expires_in", 3600))

    if not access_token:
        raise RuntimeError(f"Respuesta sin access_token: {payload}")

    expires_at = datetime.now(timezone.utc) + timedelta(seconds=expires_in)

    # Guardar tokens actualizados (MISMO realm_id)
    if not realm_id:
        # si por alguna razón no existe, no podemos construir URLs de company
        raise RuntimeError("No hay realm_id guardado. Re-conecta en /connect.")
    save_tokens(realm_id=realm_id, access_token=access_token, refresh_token=new_refresh, access_expires_at=expires_at)

    return access_token, realm_id

def _request(method: str, url: str, access_token: str, **kwargs):
    headers = kwargs.pop("headers", {})
    headers.update({
        "Authorization": f"Bearer {access_token}",
        "Accept": "application/json",
    })
    return requests.request(method, url, headers=headers, timeout=30, **kwargs)

def qbo_query(select_statement: str, access_token: str, realm_id: str) -> dict:
    url = f"{_api_base()}/v3/company/{realm_id}/query"
    params = {"query": select_statement, "minorversion": QBO_MINORVERSION}
    r = _request("GET", url, access_token, params=params)
    if r.status_code >= 400:
        raise RuntimeError(f"QBO query failed ({r.status_code}): {r.text}")
    return r.json()

def get_customers(access_token: str, realm_id: str, active_only: bool = True, max_results: int = 1000):
    where = " WHERE Active = true" if active_only else ""
    q = f"SELECT Id, DisplayName, Active FROM Customer{where} MAXRESULTS {max_results}"
    data = qbo_query(q, access_token, realm_id)
    customers = data.get("QueryResponse", {}).get("Customer", []) or []
    return [{"id": c["Id"], "name": c.get("DisplayName", f"Customer {c['Id']}")} for c in customers]

def get_accounts(access_token: str, realm_id: str, active_only: bool = True, max_results: int = 1000):
    where = " WHERE Active = true" if active_only else ""
    q = f"SELECT Id, Name, AccountType, AccountSubType, Active FROM Account{where} MAXRESULTS {max_results}"
    data = qbo_query(q, access_token, realm_id)
    accounts = data.get("QueryResponse", {}).get("Account", []) or []
    return [{
        "id": a["Id"],
        "name": a.get("Name", f"Account {a['Id']}"),
        "type": a.get("AccountType"),
        "subtype": a.get("AccountSubType"),
    } for a in accounts]
def get_profit_and_loss(access_token: str, realm_id: str, start_date: str, end_date: str,
                        accounting_method: str = "Accrual",
                        summarize_column_by: str = "Total",
                        customer_id: str | None = None) -> dict:
    """
    Llama Reports API: ProfitAndLoss
    Docs: Run Reports (P&L) y query params start_date/end_date/accounting_method. :contentReference[oaicite:1]{index=1}
    """
    url = f"{_api_base()}/v3/company/{realm_id}/reports/ProfitAndLoss"
    params = {
        "start_date": start_date,
        "end_date": end_date,
        "accounting_method": accounting_method,   # Accrual o Cash
        "summarize_column_by": summarize_column_by,  # Total, Month, etc.
        "minorversion": QBO_MINORVERSION,
    }

    # OJO: algunos reportes aceptan filtro por customer; si te diera error, lo quitamos
    if customer_id and customer_id != "all":
        params["customer"] = customer_id

    r = _request("GET", url, access_token, params=params)
    if r.status_code >= 400:
        raise RuntimeError(f"P&L report failed ({r.status_code}): {r.text}")
    return r.json()
def parse_pl_to_rows(report_json: dict) -> list[dict]:
    """
    Convierte el report JSON (Rows/Row/ColData) a una lista:
      [{"account": "Sales", "amount": 123.45}, ...]
    """
    rows_out = []

    def walk(node):
        if not node:
            return

        # "Rows" puede tener "Row" (lista)
        if isinstance(node, dict):
            if "Row" in node and isinstance(node["Row"], list):
                for r in node["Row"]:
                    walk(r)
                return

            # Cada Row puede venir como "Header", "Rows" (sub-secciones), o "ColData"
            if "Header" in node:
                # Header solo describe sección, seguimos
                pass

            # Si tiene ColData, normalmente es una fila con columnas
            if "ColData" in node and isinstance(node["ColData"], list) and len(node["ColData"]) >= 2:
                # En P&L: ColData[0] suele ser nombre, y la última el monto
                name = (node["ColData"][0].get("value") or "").strip()
                amt_raw = (node["ColData"][-1].get("value") or "").replace(",", "").strip()

                # Evitar filas vacías o de totales raros
                if name:
                    try:
                        amount = float(amt_raw) if amt_raw else 0.0
                        rows_out.append({"account": name, "amount": amount})
                    except ValueError:
                        # Si no es número, igual guardamos la fila pero amount=0
                        rows_out.append({"account": name, "amount": 0.0})

            # Si tiene sub Rows, seguimos
            if "Rows" in node:
                walk(node["Rows"])

    walk(report_json.get("Rows", {}))
    return rows_out
