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

    if not realm_id:
        raise RuntimeError("No hay realm_id guardado. Re-conecta en /connect.")

    save_tokens(
        realm_id=realm_id,
        access_token=access_token,
        refresh_token=new_refresh,
        access_expires_at=expires_at
    )
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


# -------------------------
# ✅ REPORTS API (GENÉRICO)
# -------------------------
def get_report(access_token: str, realm_id: str, report_name: str, **params) -> dict:
    """
    Llama Reports API genérico:
      /v3/company/{realm_id}/reports/{report_name}
    """
    url = f"{_api_base()}/v3/company/{realm_id}/reports/{report_name}"

    # minorversion siempre
    params = {k: v for k, v in params.items() if v is not None}
    params["minorversion"] = QBO_MINORVERSION

    r = _request("GET", url, access_token, params=params)
    if r.status_code >= 400:
        raise RuntimeError(f"QBO report '{report_name}' failed ({r.status_code}): {r.text}")
    return r.json()


# ✅ Detalle de Pérdidas y Ganancias
def get_profit_and_loss_detail(
    access_token: str,
    realm_id: str,
    start_date: str,
    end_date: str,
    accounting_method: str = "Accrual",
    summarize_column_by: str = "Total",
    customer_id: str | None = None,
) -> dict:
    params = {
        "start_date": start_date,
        "end_date": end_date,
        "accounting_method": accounting_method,
        "summarize_column_by": summarize_column_by,
    }
    if customer_id and customer_id != "all":
        params["customer"] = customer_id

    return get_report(access_token, realm_id, "ProfitAndLossDetail", **params)


# ✅ VAT - Detalle de impuesto
def get_vat_tax_detail(
    access_token: str,
    realm_id: str,
    start_date: str,
    end_date: str,
) -> dict:
    return get_report(access_token, realm_id, "TaxDetail", start_date=start_date, end_date=end_date)


# -------------------------
# ✅ Parser genérico “tal cual” Columns + Rows
# -------------------------
def parse_report_to_table(report_json: dict) -> dict:
    """
    Devuelve:
      {
        "columns": [..titulos..],
        "col_types": [..tipos..],
        "rows": [
          {"level":0, "row_type":"Header|Data|Summary", "cells":[...], "is_header":bool, "is_summary":bool},
          ...
        ]
      }
    """
    cols = report_json.get("Columns", {}).get("Column", []) or []
    col_titles = []
    col_types = []

    for c in cols:
        title = (c.get("ColTitle") or c.get("Title") or c.get("Name") or "").strip()
        col_titles.append(title if title else "Column")
        col_types.append((c.get("ColType") or "").strip())

    out_rows = []

    def row_to_cells(row_obj: dict) -> list[str]:
        coldata = row_obj.get("ColData", []) or []
        cells = []
        for i in range(len(col_titles)):
            v = ""
            if i < len(coldata):
                v = coldata[i].get("value") or ""
            cells.append(v)
        return cells

    def emit(level: int, row_type: str, cells: list[str], is_header: bool, is_summary: bool):
        out_rows.append({
            "level": level,
            "row_type": row_type,
            "cells": cells,
            "is_header": is_header,
            "is_summary": is_summary,
        })

    def walk(node, level: int):
        if not node:
            return

        if isinstance(node, dict) and "Row" in node and isinstance(node["Row"], list):
            for r in node["Row"]:
                walk(r, level)
            return

        if isinstance(node, dict):
            rt = (node.get("RowType") or "").strip()  # ✅ RowType real

            if "Header" in node and isinstance(node["Header"], dict):
                emit(level, "Header", row_to_cells(node["Header"]), True, False)

            if "ColData" in node and isinstance(node["ColData"], list) and node["ColData"]:
                if rt.lower() == "summary":
                    emit(level, "Summary", row_to_cells(node), False, True)
                else:
                    emit(level, rt if rt else "Data", row_to_cells(node), False, False)

            if "Rows" in node:
                next_level = level + 1 if rt.lower() == "section" else level
                walk(node["Rows"], next_level)

            if "Summary" in node and isinstance(node["Summary"], dict):
                emit(level, "Summary", row_to_cells(node["Summary"]), False, True)

    walk(report_json.get("Rows", {}), 0)

    return {"columns": col_titles, "col_types": col_types, "rows": out_rows}
