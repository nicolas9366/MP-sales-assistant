#!/usr/bin/env python3
"""
Mainpaper AI Sales System — Data Loader
========================================
Carga datos desde tres fuentes configuradas en config.json:
  1. Google Drive (carpeta con reportes Power BI)
  2. Gmail (adjuntos Excel de facturas/pedidos)
  3. Archivos locales (rutas absolutas configuradas)

Uso:
  python data_loader.py              # Carga todo (Drive + Gmail + Local)
  python data_loader.py --drive      # Solo Google Drive
  python data_loader.py --gmail      # Solo Gmail
  python data_loader.py --local      # Solo archivos locales
  python data_loader.py --status     # Muestra estado de la configuración
"""

import os
import sys
import json
import logging
import argparse
import base64
import io
from datetime import datetime
from pathlib import Path

import openpyxl

# ─────────────────────────────────────────────────────────────────────────────
# Logging
# ─────────────────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler("ingest_agent.log", encoding="utf-8"),
        logging.StreamHandler(sys.stdout)
    ]
)
log = logging.getLogger("data_loader")

# ─────────────────────────────────────────────────────────────────────────────
# Config Loader
# ─────────────────────────────────────────────────────────────────────────────
def load_config() -> dict:
    config_path = Path(__file__).parent / "config.json"
    if not config_path.exists():
        log.error("config.json no encontrado. Por favor crea el archivo de configuración.")
        sys.exit(1)
    with open(config_path, "r", encoding="utf-8") as f:
        return json.load(f)

# ─────────────────────────────────────────────────────────────────────────────
# Google API Auth
# ─────────────────────────────────────────────────────────────────────────────
def get_google_credentials(cfg: dict):
    """Obtiene credenciales OAuth2 para Google APIs."""
    try:
        from google.auth.transport.requests import Request
        from google.oauth2.credentials import Credentials
        from google_auth_oauthlib.flow import InstalledAppFlow
    except ImportError:
        log.error("Paquetes de Google no instalados. Ejecuta: pip install -r requirements.txt")
        return None

    SCOPES = [
        'https://www.googleapis.com/auth/gmail.modify',
        'https://www.googleapis.com/auth/drive.readonly'
    ]

    creds_file = cfg.get("google_drive", {}).get("credentials_file", "credentials.json")
    token_file = cfg.get("google_drive", {}).get("token_file", "token.json")

    creds = None
    if os.path.exists(token_file):
        creds = Credentials.from_authorized_user_file(token_file, SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
                log.info("Token renovado correctamente.")
            except Exception as e:
                log.warning(f"No se pudo renovar el token: {e}")
                creds = None

        if not creds:
            if not os.path.exists(creds_file):
                log.error(
                    f"Archivo '{creds_file}' no encontrado.\n"
                    "  → Descárgalo desde Google Cloud Console:\n"
                    "    console.cloud.google.com → APIs & Services → Credentials\n"
                    "    → Create Credentials → OAuth client ID → Desktop app\n"
                    "    → Descarga JSON y guárdalo como credentials.json"
                )
                return None
            flow = InstalledAppFlow.from_client_secrets_file(creds_file, SCOPES)
            creds = flow.run_local_server(port=0)
            log.info("Autorización completada y token guardado.")

        with open(token_file, 'w') as token:
            token.write(creds.to_json())

    return creds


# ─────────────────────────────────────────────────────────────────────────────
# Excel Parsers
# ─────────────────────────────────────────────────────────────────────────────
def parse_excel_generic(file_path: str, sheet_type: str, col_mapping: dict) -> list:
    """
    Parsea cualquier Excel/CSV según el tipo de hoja (powerbi_estadisticas / mail_pedidos).
    Intenta usar el mapeo de columnas de config.json, con fallback a auto-detección.
    """
    records = []
    ext = Path(file_path).suffix.lower()

    try:
        if ext == ".csv":
            return _parse_csv(file_path, sheet_type, col_mapping)

        wb = openpyxl.load_workbook(file_path, data_only=True)
        sheet = wb.active
        headers_raw = [str(cell.value).strip() if cell.value else "" for cell in sheet[1]]
        headers = [h.lower() for h in headers_raw]

        if sheet_type == "powerbi_estadisticas":
            records = _parse_powerbi_rows(sheet, headers, headers_raw, col_mapping)
        elif sheet_type == "mail_pedidos":
            records = _parse_orders_rows(sheet, headers, headers_raw, col_mapping)
        else:
            log.warning(f"Tipo de hoja desconocido: {sheet_type}. Usando auto-detección.")
            records = _parse_orders_rows(sheet, headers, headers_raw, col_mapping)

        wb.close()
        log.info(f"Parseadas {len(records)} filas desde: {Path(file_path).name}")

    except Exception as e:
        log.error(f"Error parseando {file_path}: {e}")

    return records


def _find_col(headers: list, aliases: list) -> int:
    """Encuentra el índice de columna buscando aliases (case-insensitive)."""
    for i, h in enumerate(headers):
        for alias in aliases:
            if alias.lower() in h:
                return i
    return -1


def _parse_powerbi_rows(sheet, headers: list, headers_raw: list, col_mapping: dict) -> list:
    m = col_mapping.get("powerbi", {})
    idx = {
        'fecha':    _find_col(headers, [m.get("fecha","fecha").lower(), "date", "periodo", "mes", "reporte"]),
        'categoria': _find_col(headers, [m.get("categoria","categoria").lower(), "category", "categoría"]),
        'subcat':   _find_col(headers, [m.get("subcategoria","subcategoria").lower(), "sub", "subcategoría"]),
        'prod':     _find_col(headers, [m.get("producto","producto").lower(), "product", "nombre", "artículo"]),
        'ventas':   _find_col(headers, [m.get("ventas","ventas").lower(), "sales", "importe", "total €"]),
        'prev':     _find_col(headers, [m.get("ventas_ano_anterior","anterior").lower(), "prev", "año anterior"]),
        'qty':      _find_col(headers, [m.get("n_total_articulos","unidades").lower(), "qty", "cantidad", "units"]),
    }

    records = []
    for row in sheet.iter_rows(min_row=2, values_only=True):
        prod = row[idx['prod']] if idx['prod'] != -1 else None
        if not prod:
            continue
        records.append({
            "fecha_reporte": _safe_date(row[idx['fecha']] if idx['fecha'] != -1 else None),
            "categoria":      str(row[idx['categoria']] or "Sin categoría").strip() if idx['categoria'] != -1 else "Sin categoría",
            "subcategoria":   str(row[idx['subcat']] or "Sin subcategoría").strip() if idx['subcat'] != -1 else "",
            "producto":       str(prod).strip(),
            "ventas":         _safe_float(row[idx['ventas']] if idx['ventas'] != -1 else 0),
            "ventas_ano_anterior": _safe_float(row[idx['prev']] if idx['prev'] != -1 else 0),
            "n_total_articulos":   _safe_int(row[idx['qty']] if idx['qty'] != -1 else 0),
            "_source": "excel"
        })
    return records


def _parse_orders_rows(sheet, headers: list, headers_raw: list, col_mapping: dict) -> list:
    m = col_mapping.get("gmail_orders", {})
    idx = {
        'order_id': _find_col(headers, [m.get("order_id","pedido").lower(), "n° pedido", "id", "folio"]),
        'fecha':    _find_col(headers, [m.get("fecha_pedido","fecha").lower(), "date", "fecha"]),
        'cliente':  _find_col(headers, [m.get("cliente_email","email").lower(), "client", "email", "comprador"]),
        'sku':      _find_col(headers, [m.get("producto_sku","sku").lower(), "sku", "barcode", "código", "referencia"]),
        'qty':      _find_col(headers, [m.get("cantidad","cantidad").lower(), "qty", "unidades", "cant"]),
        'price':    _find_col(headers, [m.get("precio_unitario","precio").lower(), "unit price", "p. unit", "precio"]),
        'total':    _find_col(headers, [m.get("total_linea","total").lower(), "total", "importe", "subtotal"]),
    }

    records = []
    for row_idx, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=2):
        sku = row[idx['sku']] if idx['sku'] != -1 else None
        if not sku:
            continue
        qty   = _safe_int(row[idx['qty']] if idx['qty'] != -1 else 1) or 1
        price = _safe_float(row[idx['price']] if idx['price'] != -1 else 0)
        total = _safe_float(row[idx['total']] if idx['total'] != -1 else 0) or round(price * qty, 2)
        records.append({
            "order_id":        str(row[idx['order_id']] or f"LOCAL_{row_idx}").strip() if idx['order_id'] != -1 else f"LOCAL_{row_idx}",
            "fecha_pedido":    _safe_date(row[idx['fecha']] if idx['fecha'] != -1 else None),
            "cliente_email":   str(row[idx['cliente']] or "desconocido@local").strip() if idx['cliente'] != -1 else "local",
            "producto_sku":    str(sku).strip(),
            "cantidad":        qty,
            "precio_unitario": price,
            "total_linea":     total,
            "_source": "excel"
        })
    return records


def _parse_csv(file_path: str, sheet_type: str, col_mapping: dict) -> list:
    """Parsea CSV usando el mismo sistema de mapeo."""
    import csv
    records = []
    try:
        with open(file_path, "r", encoding="utf-8-sig") as f:
            reader = csv.DictReader(f)
            headers = [h.lower().strip() for h in (reader.fieldnames or [])]
            for row in reader:
                row_lower = {k.lower().strip(): v for k, v in row.items()}
                if sheet_type == "powerbi_estadisticas":
                    prod = row_lower.get("producto") or row_lower.get("product") or row_lower.get("nombre")
                    if prod:
                        records.append({
                            "fecha_reporte": row_lower.get("fecha_reporte", datetime.now().strftime("%Y-%m-%d")),
                            "categoria":     row_lower.get("categoria", "Sin categoría"),
                            "subcategoria":  row_lower.get("subcategoria", ""),
                            "producto":      prod.strip(),
                            "ventas":        _safe_float(row_lower.get("ventas", 0)),
                            "ventas_ano_anterior": _safe_float(row_lower.get("ventas_ano_anterior", 0)),
                            "n_total_articulos":   _safe_int(row_lower.get("n_total_articulos", 0)),
                            "_source": "csv"
                        })
                else:
                    sku = row_lower.get("sku") or row_lower.get("producto_sku") or row_lower.get("código")
                    if sku:
                        records.append({
                            "order_id":       row_lower.get("order_id", f"CSV_{len(records)}"),
                            "fecha_pedido":   row_lower.get("fecha_pedido", datetime.now().strftime("%Y-%m-%d")),
                            "cliente_email":  row_lower.get("cliente_email", "csv@import"),
                            "producto_sku":   sku.strip(),
                            "cantidad":       _safe_int(row_lower.get("cantidad", 1)),
                            "precio_unitario": _safe_float(row_lower.get("precio_unitario", 0)),
                            "total_linea":    _safe_float(row_lower.get("total_linea", 0)),
                            "_source": "csv"
                        })
    except Exception as e:
        log.error(f"Error parseando CSV {file_path}: {e}")
    return records


def _safe_float(v) -> float:
    try:
        return float(str(v).replace(",", ".").replace("€", "").strip()) if v else 0.0
    except Exception:
        return 0.0

def _safe_int(v) -> int:
    try:
        return int(float(str(v).replace(",", "").strip())) if v else 0
    except Exception:
        return 0

def _safe_date(v) -> str:
    if not v:
        return datetime.now().strftime("%Y-%m-%d")
    if isinstance(v, datetime):
        return v.strftime("%Y-%m-%d")
    s = str(v).split(" ")[0].split("T")[0]
    return s if s else datetime.now().strftime("%Y-%m-%d")


# ─────────────────────────────────────────────────────────────────────────────
# JSON Database Helpers
# ─────────────────────────────────────────────────────────────────────────────
def load_json_db(path: str) -> list:
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return []

def save_json_db(path: str, records: list):
    os.makedirs(os.path.dirname(path) if os.path.dirname(path) else ".", exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(records, f, ensure_ascii=False, indent=2)
    log.info(f"💾 Guardados {len(records)} registros → {path}")

def load_processed_ids(log_path: str) -> set:
    if os.path.exists(log_path):
        try:
            with open(log_path, "r", encoding="utf-8") as f:
                return set(json.load(f))
        except Exception:
            pass
    return set()

def mark_processed(log_path: str, file_id: str):
    ids = load_processed_ids(log_path)
    ids.add(file_id)
    os.makedirs(os.path.dirname(log_path) if os.path.dirname(log_path) else ".", exist_ok=True)
    with open(log_path, "w", encoding="utf-8") as f:
        json.dump(list(ids), f, ensure_ascii=False, indent=2)


# ─────────────────────────────────────────────────────────────────────────────
# SOURCE 1: Google Drive
# ─────────────────────────────────────────────────────────────────────────────
def sync_google_drive(cfg: dict):
    drive_cfg = cfg.get("google_drive", {})
    if not drive_cfg.get("enabled", False):
        log.info("Google Drive desactivado en config.json. Saltando.")
        return

    folder_id = drive_cfg.get("folder_id", "")
    if not folder_id:
        log.warning("google_drive.folder_id vacío en config.json. Saltando Drive.")
        return

    log.info(f"📂 Sincronizando Google Drive — Carpeta: {folder_id}")

    creds = get_google_credentials(cfg)
    if not creds:
        return

    try:
        from googleapiclient.discovery import build
        from googleapiclient.errors import HttpError
        from googleapiclient.http import MediaIoBaseDownload
    except ImportError:
        log.error("google-api-python-client no instalado. Ejecuta: pip install -r requirements.txt")
        return

    out_cfg = cfg.get("output", {})
    pbi_path = out_cfg.get("powerbi_file", "data/powerbi_estadisticas.json")
    processed_log = out_cfg.get("processed_files_log", "data/processed_files.json")
    col_mapping = cfg.get("column_mapping", {})

    try:
        service = build('drive', 'v3', credentials=creds)
        query = (
            f"'{folder_id}' in parents and trashed = false and ("
            "mimeType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' or "
            "mimeType = 'application/vnd.ms-excel' or "
            "mimeType = 'text/csv'"
            ")"
        )
        results = service.files().list(q=query, fields="files(id, name, modifiedTime)").execute()
        files = results.get("files", [])

        if not files:
            log.info("No se encontraron archivos Excel/CSV en la carpeta de Drive.")
            return

        processed_ids = load_processed_ids(processed_log)
        new_records = []
        new_count = 0

        for f in files:
            file_id = f["id"]
            file_name = f["name"]

            if file_id in processed_ids:
                log.info(f"  Saltando (ya procesado): {file_name}")
                continue

            log.info(f"  ⬇️  Descargando: {file_name}")
            temp_path = f"data/temp_drive_{file_id}.xlsx"

            try:
                request = service.files().get_media(fileId=file_id)
                fh = io.BytesIO()
                downloader = MediaIoBaseDownload(fh, request)
                done = False
                while not done:
                    _, done = downloader.next_chunk()

                os.makedirs("data", exist_ok=True)
                with open(temp_path, "wb") as out_file:
                    out_file.write(fh.getvalue())

                records = parse_excel_generic(temp_path, "powerbi_estadisticas", col_mapping)
                if records:
                    new_records.extend(records)
                    new_count += 1
                    log.info(f"  ✅ {file_name}: {len(records)} filas importadas")

                mark_processed(processed_log, file_id)

            except Exception as e:
                log.error(f"  ❌ Error procesando {file_name}: {e}")
            finally:
                if os.path.exists(temp_path):
                    os.remove(temp_path)

        if new_records:
            existing = load_json_db(pbi_path)
            existing.extend(new_records)
            save_json_db(pbi_path, existing)
            log.info(f"📊 Drive: {new_count} archivo(s) nuevos, {len(new_records)} registros añadidos → {pbi_path}")
        else:
            log.info("Drive: No hay archivos nuevos para procesar.")

    except Exception as e:
        log.error(f"Error en sincronización de Drive: {e}")


# ─────────────────────────────────────────────────────────────────────────────
# SOURCE 2: Gmail
# ─────────────────────────────────────────────────────────────────────────────
def sync_gmail(cfg: dict):
    gmail_cfg = cfg.get("gmail", {})
    if not gmail_cfg.get("enabled", False):
        log.info("Gmail desactivado en config.json. Saltando.")
        return

    log.info(f"📧 Sincronizando Gmail — Cuenta: {gmail_cfg.get('account', '')}")

    creds = get_google_credentials(cfg)
    if not creds:
        return

    try:
        from googleapiclient.discovery import build
        from googleapiclient.errors import HttpError
    except ImportError:
        log.error("google-api-python-client no instalado.")
        return

    out_cfg = cfg.get("output", {})
    orders_path = out_cfg.get("gmail_orders_file", "data/mail_pedidos.json")
    col_mapping = cfg.get("column_mapping", {})

    search_query = gmail_cfg.get(
        "search_query",
        "has:attachment filename:xlsx subject:(pedido OR facturación)"
    )

    try:
        service = build('gmail', 'v1', credentials=creds)
        results = service.users().messages().list(userId='me', q=search_query).execute()
        messages = results.get('messages', [])

        if not messages:
            log.info("Gmail: No se encontraron correos que coincidan con la búsqueda.")
            return

        log.info(f"Gmail: {len(messages)} correo(s) encontrado(s).")
        new_records = []

        for msg_meta in messages:
            msg_id = msg_meta['id']
            message = service.users().messages().get(userId='me', id=msg_id).execute()

            headers = message['payload']['headers']
            subject = next((h['value'] for h in headers if h['name'].lower() == 'subject'), 'Sin asunto')
            sender  = next((h['value'] for h in headers if h['name'].lower() == 'from'), '')
            date_str = next((h['value'] for h in headers if h['name'].lower() == 'date'), '')

            log.info(f"  📨 Correo: '{subject}' — {sender}")

            parts = message['payload'].get('parts', [])
            # Aplanar partes anidadas
            all_parts = []
            def flatten_parts(p_list):
                for p in p_list:
                    all_parts.append(p)
                    if p.get("parts"):
                        flatten_parts(p["parts"])
            flatten_parts(parts)

            for part in all_parts:
                fname = part.get('filename', '')
                if fname and any(fname.lower().endswith(ext) for ext in ['.xlsx', '.xls']):
                    attachment_id = part['body'].get('attachmentId')
                    if not attachment_id:
                        continue

                    try:
                        attachment = service.users().messages().attachments().get(
                            userId='me', messageId=msg_id, id=attachment_id
                        ).execute()
                        file_data = base64.urlsafe_b64decode(attachment['data'].encode('UTF-8'))

                        os.makedirs("data", exist_ok=True)
                        temp_path = f"data/temp_gmail_{msg_id}.xlsx"
                        with open(temp_path, 'wb') as f:
                            f.write(file_data)

                        records = parse_excel_generic(temp_path, "mail_pedidos", col_mapping)

                        # Enriquecer con metadatos del email
                        for r in records:
                            r["_email_subject"] = subject
                            r["_email_from"]    = sender
                            r["_email_date"]    = date_str
                            r["_msg_id"]        = msg_id
                            if r.get("cliente_email") in ("local", "desconocido@local", ""):
                                r["cliente_email"] = sender

                        new_records.extend(records)
                        log.info(f"  ✅ Adjunto '{fname}': {len(records)} líneas de pedido")

                        if os.path.exists(temp_path):
                            os.remove(temp_path)

                    except Exception as e:
                        log.error(f"  ❌ Error procesando adjunto '{fname}': {e}")

            # Marcar como leído si está configurado
            if gmail_cfg.get("mark_as_read_after_process", True):
                try:
                    service.users().messages().batchModify(
                        userId='me',
                        body={'ids': [msg_id], 'removeLabelIds': ['UNREAD']}
                    ).execute()
                except Exception:
                    pass

        if new_records:
            existing = load_json_db(orders_path)
            # Deduplicar por order_id + msg_id
            existing_keys = {r.get("order_id", "") + r.get("_msg_id", "") for r in existing}
            unique_new = [r for r in new_records if (r.get("order_id","") + r.get("_msg_id","")) not in existing_keys]
            existing.extend(unique_new)
            save_json_db(orders_path, existing)
            log.info(f"📬 Gmail: {len(unique_new)} líneas nuevas de pedido guardadas → {orders_path}")
        else:
            log.info("Gmail: No se encontraron adjuntos Excel con pedidos.")

    except Exception as e:
        log.error(f"Error en sincronización de Gmail: {e}")


# ─────────────────────────────────────────────────────────────────────────────
# SOURCE 3: Archivos Locales
# ─────────────────────────────────────────────────────────────────────────────
def sync_local_files(cfg: dict):
    local_cfg = cfg.get("local_files", {})
    if not local_cfg.get("enabled", False):
        log.info("Archivos locales desactivados en config.json. Saltando.")
        return

    paths = local_cfg.get("paths", [])
    col_mapping = cfg.get("column_mapping", {})
    out_cfg = cfg.get("output", {})

    for entry in paths:
        file_path = entry.get("path", "").strip()
        label     = entry.get("label", "Archivo local")
        ftype     = entry.get("type", "powerbi_estadisticas")

        if not file_path:
            log.info(f"  '{label}': ruta vacía, saltando. Configura 'path' en config.json → local_files.paths")
            continue

        if not os.path.exists(file_path):
            log.warning(f"  '{label}': archivo no encontrado → {file_path}")
            continue

        log.info(f"📁 Cargando local — {label}: {file_path}")
        records = parse_excel_generic(file_path, ftype, col_mapping)

        if not records:
            log.warning(f"  No se extrajeron registros de: {file_path}")
            continue

        if ftype == "powerbi_estadisticas":
            out_path = out_cfg.get("powerbi_file", "data/powerbi_estadisticas.json")
        else:
            out_path = out_cfg.get("gmail_orders_file", "data/mail_pedidos.json")

        existing = load_json_db(out_path)
        existing.extend(records)
        save_json_db(out_path, existing)
        log.info(f"✅ '{label}': {len(records)} registros guardados → {out_path}")


# ─────────────────────────────────────────────────────────────────────────────
# Status Report
# ─────────────────────────────────────────────────────────────────────────────
def print_status(cfg: dict):
    print("\n" + "─" * 60)
    print("  MAINPAPER DATA LOADER — Estado de Configuración")
    print("─" * 60)

    # Google Drive
    d = cfg.get("google_drive", {})
    folder_id = d.get("folder_id", "")
    drive_ok = bool(folder_id and folder_id != "YOUR_GOOGLE_DRIVE_FOLDER_ID_HERE")
    print(f"\n📂 Google Drive:  {'✅ Configurado' if drive_ok else '⚠️  folder_id vacío'}")
    if drive_ok:
        print(f"   Carpeta: {d.get('folder_url', folder_id)}")
    print(f"   Habilitado: {d.get('enabled', False)}")

    # Gmail
    g = cfg.get("gmail", {})
    print(f"\n📧 Gmail:         {'✅ Configurado' if g.get('account') else '⚠️  Sin cuenta'}")
    print(f"   Cuenta: {g.get('account', 'no configurada')}")
    print(f"   Habilitado: {g.get('enabled', False)}")

    # Local files
    lf = cfg.get("local_files", {})
    paths = [p for p in lf.get("paths", []) if p.get("path", "").strip()]
    print(f"\n📁 Archivos locales: {len(paths)} ruta(s) configurada(s)")
    for p in paths:
        exists = os.path.exists(p["path"])
        print(f"   {'✅' if exists else '❌'} [{p.get('type','')}] {p.get('label','')}: {p['path']}")

    # Output files
    out = cfg.get("output", {})
    print("\n📊 Datos actuales:")
    for label, key in [("Power BI",  "powerbi_file"), ("Pedidos Gmail", "gmail_orders_file")]:
        path = out.get(key, "")
        if path and os.path.exists(path):
            data = load_json_db(path)
            print(f"   ✅ {label}: {len(data)} registros → {path}")
        else:
            print(f"   ⚪ {label}: sin datos → {path}")

    # Credentials
    creds_file = d.get("credentials_file", "credentials.json")
    token_file = d.get("token_file", "token.json")
    print(f"\n🔑 Credenciales:")
    print(f"   credentials.json: {'✅ existe' if os.path.exists(creds_file) else '❌ FALTA'}")
    print(f"   token.json:       {'✅ existe' if os.path.exists(token_file) else '⚠️  se generará en primer uso'}")
    print("\n" + "─" * 60 + "\n")


# ─────────────────────────────────────────────────────────────────────────────
# Main
# ─────────────────────────────────────────────────────────────────────────────
def main():
    parser = argparse.ArgumentParser(
        description="Mainpaper Data Loader — Sincroniza Excel desde Google Drive, Gmail y archivos locales."
    )
    parser.add_argument("--drive",  action="store_true", help="Solo sincronizar Google Drive")
    parser.add_argument("--gmail",  action="store_true", help="Solo sincronizar Gmail")
    parser.add_argument("--local",  action="store_true", help="Solo archivos locales")
    parser.add_argument("--status", action="store_true", help="Ver estado de la configuración")
    args = parser.parse_args()

    cfg = load_config()

    if args.status:
        print_status(cfg)
        return

    os.makedirs(cfg.get("output", {}).get("data_directory", "data"), exist_ok=True)

    if args.drive:
        sync_google_drive(cfg)
    elif args.gmail:
        sync_gmail(cfg)
    elif args.local:
        sync_local_files(cfg)
    else:
        # Carga todo
        log.info("═══════ Mainpaper Data Loader — Inicio ═══════")
        sync_local_files(cfg)
        sync_google_drive(cfg)
        sync_gmail(cfg)
        log.info("═══════ Sincronización completa ═══════")


if __name__ == "__main__":
    main()
