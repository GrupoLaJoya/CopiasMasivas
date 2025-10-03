# upload_graph_from_excel.py
import os
import math
import json
import msal
import pandas as pd
import requests
from urllib.parse import quote

GRAPH = "https://graph.microsoft.com/v1.0"

# =========================
# 1) CONFIG
# =========================

with open("config_tenant.json", "r", encoding="utf-8") as f:
    config = json.load(f)

TENANT_ID     = config["sharepoint"]["tenant_id"]
CLIENT_ID     = config["sharepoint"]["client_id"]
CLIENT_SECRET = config["sharepoint"]["client_secret"]

SITE_DOMAIN   = "lajoyaminingsac.sharepoint.com"  # solo host
SITE_NAME     = "BacklogTI"                       # lo que va tras /sites/
DRIVE_NAME    = "Facturas"                        # biblioteca (drive) dentro del sitio

# Excel
EXCEL_FILE    = "masivo.xlsx"
SHEET_NAME    = 0                   # o el nombre de hoja
COL_PREFIX    = "CARPETAS"           # ej. 0701-0057
COL_FILE      = "MASIVO"         # ruta local del archivo a subir por fila
COL_BASE      = "BASE_REL_PATH"     # opcional: si el Excel trae la base por fila (ej. LJC/2025/JUL)

# Si una fila NO tiene base, se usa esta por defecto:
DEFAULT_BASE_REL_PATH = "LJC/2025/JUL"

# ¿Crear segmentos ausentes en la base?
CREATE_MISSING = False


# =========================
# 2) AUTH (app-only)
# =========================
def get_access_token():
    app = msal.ConfidentialClientApplication(
        CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{TENANT_ID}",
        client_credential=CLIENT_SECRET,
    )
    tok = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    if "access_token" not in tok:
        raise RuntimeError(f"No se obtuvo token: {tok}")
    return tok["access_token"]

TOKEN = get_access_token()
HEADERS = {"Authorization": f"Bearer {TOKEN}"}


# =========================
# 3) Helpers HTTP
# =========================
def gget(url, **kw):
    r = requests.get(url, headers=HEADERS, timeout=30, **kw)
    if r.status_code >= 400:
        print("GET ERR:", r.status_code, url, r.text)
        r.raise_for_status()
    return r.json()

def gpost(url, body: dict):
    r = requests.post(url, headers={**HEADERS, "Content-Type": "application/json"}, json=body, timeout=60)
    if r.status_code >= 400:
        print("POST ERR:", r.status_code, url, r.text)
        r.raise_for_status()
    return r.json()

def gput(url, data=None, headers_extra=None):
    hdrs = dict(HEADERS)
    if headers_extra:
        hdrs.update(headers_extra)
    r = requests.put(url, headers=hdrs, data=data, timeout=120)
    if r.status_code not in (200, 201, 202):
        print("PUT ERR:", r.status_code, url, r.text)
        r.raise_for_status()
    return r.json() if r.text else {}


# =========================
# 4) site_id y drive_id
# =========================
def get_site_id(site_domain: str, site_name: str) -> str:
    url = f"{GRAPH}/sites/{site_domain}:/sites/{site_name}"
    data = gget(url)
    return data["id"]

def get_drive_id_by_name(site_id: str, drive_name: str) -> str:
    url = f"{GRAPH}/sites/{site_id}/drives?$select=id,name"
    data = gget(url)
    for d in data.get("value", []):
        if d["name"].lower() == drive_name.lower():
            return d["id"]
    raise RuntimeError(f"No encontré la biblioteca/drive '{drive_name}' en el sitio.")

SITE_ID = get_site_id(SITE_DOMAIN, SITE_NAME)
DRIVE_ID = get_drive_id_by_name(SITE_ID, DRIVE_NAME)


# =========================
# 5) Navegación de carpetas
# =========================
def get_drive_root() -> dict:
    url = f"{GRAPH}/sites/{SITE_ID}/drives/{DRIVE_ID}/root?$select=id,name,webUrl"
    return gget(url)

def get_drive_root_(drive_id: str):
    url = f"{GRAPH}/sites/{SITE_ID}/drives/{drive_id}/root?$select=id,name,webUrl"
    return gget(url)

def get_item_by_path(rel_path):
    # rel_path: e.g. "LJC/2025" (sin barra inicial)
    rel = quote(rel_path.strip("/"), safe="/")
    url = f"{GRAPH}/sites/{SITE_ID}/drives/{DRIVE_ID}/root:/{rel}"
    return gget(url)  # driveItem del folder

def list_subfolders_by_id(site_id, drive_id, parent_id):
    url = f"{GRAPH}/sites/{site_id}/drives/{drive_id}/items/{parent_id}/children?$select=id,name,folder,webUrl"
    folders = []
    while url:
        data = gget(url)
        # 👉 solo los que tienen facet 'folder'
        folders.extend([it for it in data.get("value", []) if it.get("folder")])
        url = data.get("@odata.nextLink")
    return folders

def list_subfolders_by_path(site_id, drive_id, rel_path):
    parent = get_item_by_path(site_id, drive_id, rel_path)
    return list_subfolders_by_id(site_id, drive_id, parent["id"])

def list_children_root_(drive_name: str):
    site_list = get_site_id(SITE_DOMAIN, SITE_NAME)
    drive_list = get_drive_id_by_name(site_list, drive_name)
    current = get_drive_root()
    url = f"{GRAPH}/sites/{site_list}/drives/{drive_list}/items/{current['id']}/children?$select=id,name,folder,webUrl"
    items = []
    # print(site_id, drive_id, parent_id, parent_id)
    while url:
        data = gget(url)
        items.extend([it for it in data.get("value", []) if it.get("folder")])
        url = data.get("@odata.nextLink")
    return items

def list_children_root():
    site = get_site_id(SITE_DOMAIN, SITE_NAME)
    drive = get_drive_id_by_name(SITE_ID, DRIVE_NAME)
    current = get_drive_root()
    url = f"{GRAPH}/sites/{site}/drives/{drive}/items/{current['id']}/children?$select=id,name,folder,webUrl"
    items = []
    # print(site_id, drive_id, parent_id, parent_id)
    while url:
        data = gget(url)
        items.extend([it for it in data.get("value", []) if it.get("folder")])
        url = data.get("@odata.nextLink")
    return items

def list_children(site_id: str, drive_id: str, parent_id: str):
    url = f"{GRAPH}/sites/{site_id}/drives/{drive_id}/items/{parent_id}/children?$select=id,name,folder,webUrl"
    items = []
    #print(site_id, drive_id, parent_id, parent_id)
    while url:
        data = gget(url)
        items.extend(data.get("value", []))
        url = data.get("@odata.nextLink")
    return items

def resolve_child_folder(site_id: str, drive_id: str, parent_id: str, segment: str):
    """Busca subcarpeta por igualdad (case-insensitive), prefijo, luego contiene."""
    kids = list_children(site_id, drive_id, parent_id)
    seg = segment.strip().lower()
    folders = [k for k in kids if k.get("folder")]

    for k in folders:
        if k["name"].lower() == seg:
            return k

    pref = [k for k in folders if k["name"].lower().startswith(seg)]
    if len(pref) == 1:
        return pref[0]
    if len(pref) > 1:
        pref.sort(key=lambda x: len(x["name"]), reverse=True)
        return pref[0]

    contains = [k for k in folders if seg in k["name"].lower()]
    return contains[0] if contains else None

def create_child_folder(site_id: str, drive_id: str, parent_id: str, name: str):
    url = f"{GRAPH}/sites/{site_id}/drives/{drive_id}/items/{parent_id}/children"
    body = {"name": name, "folder": {}, "@microsoft.graph.conflictBehavior": "fail"}
    return gpost(url, body)

def walk_path(site_id: str, drive_id: str, rel_path: str, create_if_missing: bool = False) -> dict:
    """Resuelve 'LJC/2025/JUL' segmento por segmento desde el root del drive."""
    current = get_drive_root(site_id, drive_id)
    path = (rel_path or "").strip("/")

    if not path:
        return current

    for seg in path.split("/"):
        match = resolve_child_folder(site_id, drive_id, current["id"], seg)
        if match:
            current = match
            continue
        if not create_if_missing:
            raise FileNotFoundError(f"No encontré la carpeta '{seg}' dentro de '{current['name']}'")
        current = create_child_folder(site_id, drive_id, current["id"], seg)
    return current

def resolve_leaf_by_prefix(site_id: str, drive_id: str, parent_id: str, wanted_prefix: str):
    match = resolve_child_folder(site_id, drive_id, parent_id, wanted_prefix)
    if not match:
        raise FileNotFoundError(f"No encontré subcarpeta que empiece/contenga '{wanted_prefix}'.")
    return match


# =========================
# 6) Upload (simple/sesión)
# =========================
def upload_small(site_id: str, drive_id: str, folder_id: str, file_path: str):
    name = os.path.basename(file_path)
    url = f"{GRAPH}/sites/{site_id}/drives/{drive_id}/items/{folder_id}:/{quote(name)}:/content"
    with open(file_path, "rb") as f:
        up = gput(url, data=f)
    return up

def create_upload_session(site_id: str, drive_id: str, folder_id: str, file_name: str):
    url = f"{GRAPH}/sites/{site_id}/drives/{drive_id}/items/{folder_id}:/{quote(file_name)}:/createUploadSession"
    body = {"item": {"@microsoft.graph.conflictBehavior": "replace"}}
    return gpost(url, body)

def upload_large(site_id: str, drive_id: str, folder_id: str, file_path: str, chunk_size=5*1024*1024):
    name = os.path.basename(file_path)
    session = create_upload_session(site_id, drive_id, folder_id, name)
    upload_url = session["uploadUrl"]

    size = os.path.getsize(file_path)
    with open(file_path, "rb") as f:
        start = 0
        while start < size:
            end = min(start + chunk_size, size) - 1
            length = end - start + 1
            chunk = f.read(length)

            headers = {
                "Content-Length": str(length),
                "Content-Range": f"bytes {start}-{end}/{size}",
            }
            r = requests.put(upload_url, headers=headers, data=chunk, timeout=180)
            if r.status_code not in (200, 201, 202):
                print("Chunk ERR:", r.status_code, r.text)
                r.raise_for_status()

            if r.status_code in (200, 201):
                return r.json()
            start = end + 1

    raise RuntimeError("Upload session no confirmó finalización.")

def upload_auto(site_id: str, drive_id: str, folder_id: str, file_path: str, threshold=4*1024*1024):
    size = os.path.getsize(file_path)
    if size < threshold:
        return upload_small(site_id, drive_id, folder_id, file_path)
    return upload_large(site_id, drive_id, folder_id, file_path)


# =========================
# 7) Procesar Excel
# =========================
def process_excel(
    excel_file: str,
    sheet_name=0,
    col_prefix=COL_PREFIX,
    col_file=COL_FILE,
    col_base=COL_BASE,
    default_base=DEFAULT_BASE_REL_PATH,
    create_missing=CREATE_MISSING,
    detracciones=False
):
    # Leer todo como texto para no romper códigos/prefijos (evita ".0")
    df = pd.read_excel(excel_file, sheet_name=sheet_name, dtype=str, keep_default_na=False)

    # Normaliza encabezados (trim)
    df.columns = [str(c).strip() for c in df.columns]

    # Validar columnas obligatorias
    if col_prefix not in df.columns:
        raise ValueError(f"No encuentro la columna '{col_prefix}' (prefijo) en {excel_file}")
    if col_file not in df.columns:
        raise ValueError(f"No encuentro la columna '{col_file}' (ruta de archivo) en {excel_file}")

    ok = 0
    fail = 0

    for i, row in df.iterrows():
        leaf = (row[col_prefix] or "").strip()
        if not detracciones:
            file_path = (row[col_file] or "").strip()
        else:
            file_path = (f"detracciones/{row[col_file]}.pdf" or "").strip()
        base_rel = (row[col_base].strip() if col_base in df.columns and str(row[col_base]).strip() else default_base)

        if not leaf or not file_path:
            print(f"[{i}] Saltado (faltan datos): leaf='{leaf}' file='{file_path}'")
            fail += 1
            continue

        if not os.path.isfile(file_path):
            print(f"[{i}] Archivo no encontrado: {file_path}")
            fail += 1
            continue

        try:
            base_item = walk_path(SITE_ID, DRIVE_ID, base_rel, create_if_missing=create_missing)
            leaf_item = resolve_leaf_by_prefix(SITE_ID, DRIVE_ID, base_item["id"], leaf)
            up = upload_auto(SITE_ID, DRIVE_ID, leaf_item["id"], file_path)
            print(f"[{i}] ✅ {os.path.basename(file_path)} → {up.get('webUrl')}")
            ok += 1
        except Exception as e:
            print(f"[{i}] ❌ {leaf}: {e}")
            fail += 1

    print(f"\nResumen: OK={ok}  FALLIDOS={fail}")


# =========================
# 8) MAIN
# =========================
if __name__ == "__main__":
    process_excel(
        EXCEL_FILE,
        sheet_name=SHEET_NAME,
        col_prefix=COL_PREFIX,
        col_file=COL_FILE,
        col_base=COL_BASE,                  # si tu Excel NO tiene la base por fila, deja esta col y usará default_base
        default_base=DEFAULT_BASE_REL_PATH,
        create_missing=CREATE_MISSING
    )
