import os
import json
import requests
import msal
import pandas as pd
from urllib.parse import quote

# ------------------------------
# 🔹 1) Leer configuración
# ------------------------------
with open("config_tenant.json", "r", encoding="utf-8") as f:
    config = json.load(f)

tenant_id       = config["sharepoint"]["tenant_id"]
client_id       = config["sharepoint"]["client_id"]
client_secret   = config["sharepoint"]["client_secret"]
site_name       = config["sharepoint"]["site_name"]        # p.ej. "BacklogTI"
site_domain     = config["sharepoint"]["site_domain"]      # p.ej. "lajoyaminingsac.sharepoint.com"
document_library= config["sharepoint"]["document_library"] # p.ej. "Documentos compartidos"

# ------------------------------
# 🔹 2) MSAL: token app-only
# ------------------------------
authority = f"https://login.microsoftonline.com/{tenant_id}"
scope = ["https://graph.microsoft.com/.default"]

app = msal.ConfidentialClientApplication(
    client_id, authority=authority, client_credential=client_secret
)
tok = app.acquire_token_for_client(scopes=scope)
if "access_token" not in tok:
    raise RuntimeError(f"No se pudo obtener token: {tok}")

headers = {"Authorization": f"Bearer {tok['access_token']}"}
GRAPH = "https://graph.microsoft.com/v1.0"

# ------------------------------
# 🔹 3) Utilidades Graph
# ------------------------------
def gget(url, **kw):
    r = requests.get(url, headers=headers, timeout=30, **kw)
    if r.status_code >= 400:
        print("GET ERR:", r.status_code, url, r.text)
        r.raise_for_status()
    return r.json()

def gpost(url, json_body):
    r = requests.post(url, headers={**headers, "Content-Type": "application/json"}, json=json_body, timeout=30)
    if r.status_code >= 400:
        print("POST ERR:", r.status_code, url, r.text)
        r.raise_for_status()
    return r.json()

def gupload(url, data):
    r = requests.put(url, headers=headers, data=data, timeout=120)
    if r.status_code not in (200, 201):
        print("PUT ERR:", r.status_code, url, r.text)
        r.raise_for_status()
    return r.json()

# ------------------------------
# 🔹 4) siteId y driveId
# ------------------------------
def get_site_id():
    url = f"{GRAPH}/sites/{site_domain}:/sites/{site_name}"
    resp = gget(url)
    return resp["id"]

site_id = get_site_id()

def get_drive_id_by_name(drive_name: str):
    url = f"{GRAPH}/sites/{site_id}/drives?$select=id,name"
    data = gget(url)
    for d in data.get("value", []):
        if d["name"].lower() == drive_name.lower():
            return d["id"]
    raise RuntimeError(f"No encontré la biblioteca '{drive_name}' en el sitio.")

drive_id = get_drive_id_by_name(document_library)

# ------------------------------
# 🔹 5) Resolver carpeta por prefijo
# ------------------------------
def get_item_by_path(rel_path: str):
    """
    Obtiene un driveItem por ruta relativa dentro de la biblioteca.
    rel_path: 'Facturas/LJC/2025/JUL'
    """
    rel_path = rel_path.strip("/")
    url = f"{GRAPH}/sites/{site_id}/drives/{drive_id}/root:/{quote(rel_path, safe='/')}"
    return gget(url)  # {id, name, ...}

def list_children(folder_id: str):
    """Lista subcarpetas inmediatas de folder_id (paginado)."""
    url = f"{GRAPH}/sites/{site_id}/drives/{drive_id}/items/{folder_id}/children?$select=id,name,folder"
    items = []
    while url:
        data = gget(url)
        items.extend([it for it in data.get("value", []) if it.get("folder")])
        url = data.get("@odata.nextLink")
    return items

def find_child_by_prefix(parent_id: str, wanted_prefix: str):
    """
    Devuelve la subcarpeta cuyo nombre:
      1) coincide exacto (case-insensitive), o
      2) empieza con wanted_prefix, o
      3) lo contiene (fallback).
    """
    wanted = str(wanted_prefix).strip()
    kids = list_children(parent_id)

    # exacto
    for it in kids:
        if it["name"].lower() == wanted.lower():
            return it

    # empieza con
    pref = wanted.lower()
    matches = [it for it in kids if it["name"].lower().startswith(pref)]
    if len(matches) == 1:
        return matches[0]
    if len(matches) > 1:
        # el más largo (más específico)
        matches.sort(key=lambda x: len(x["name"]), reverse=True)
        return matches[0]

    # contiene
    contains = [it for it in kids if pref in it["name"].lower()]
    return contains[0] if contains else None

def create_folder(parent_id: str, name: str):
    """Crea una subcarpeta (si quieres permitir crear cuando no existe)."""
    url = f"{GRAPH}/sites/{site_id}/drives/{drive_id}/items/{parent_id}/children"
    body = {"name": name, "folder": {}, "@microsoft.graph.conflictBehavior": "fail"}
    return gpost(url, body)

def ensure_target_folder(base_folder: str, excel_leaf: str, create_if_missing=False):
    """
    base_folder = 'Facturas/LJC/2025/JUL'
    excel_leaf  = '0701-0057' (del Excel)
    Retorna {id, name} de la carpeta real (p.ej. '0701-0057_BCP')
    """
    parent = get_item_by_path(base_folder)  # id del padre real
    match = find_child_by_prefix(parent["id"], excel_leaf)
    if match:
        return match
    if create_if_missing:
        created = create_folder(parent["id"], str(excel_leaf).strip())
        return created
    raise FileNotFoundError(
        f"No encontré subcarpeta que empiece por '{excel_leaf}' dentro de '{base_folder}'."
    )

# ------------------------------
# 🔹 6) Subir por ID de carpeta
# ------------------------------
def upload_file_to_folder_id(folder_id: str, local_file: str):
    file_name = os.path.basename(local_file)
    url = f"{GRAPH}/sites/{site_id}/drives/{drive_id}/items/{folder_id}:/{quote(file_name)}:/content"
    with open(local_file, "rb") as f:
        info = gupload(url, f)
    print(f"✅ Subido: {info.get('name')}  →  {info.get('webUrl')}")
    return info

# ------------------------------
# 🔹 7) Carga masiva desde Excel (columna = prefijo)
# ------------------------------
def upload_from_excel(excel_file: str, column_name: str, local_file: str, base_folder: str, create_if_missing=False):
    # Leer TODO como texto para evitar "0701-0057" -> "701-0057.0"
    df = pd.read_excel(excel_file, dtype=str, keep_default_na=False)

    if column_name not in df.columns:
        raise ValueError(f"La columna '{column_name}' no existe en {excel_file}")

    for raw in df[column_name]:
        leaf = (raw or "").strip()
        if not leaf:
            continue
        try:
            target = ensure_target_folder(base_folder, leaf, create_if_missing=create_if_missing)
            upload_file_to_folder_id(target["id"], local_file)
        except Exception as ex:
            print(f"⚠️ {leaf}: {ex}")

# ------------------------------
# 🔹 8) Ejecución
# ------------------------------
if __name__ == "__main__":
    # Parámetros de ejemplo
    excel_path  = "masivo.xlsx"                     # Excel con lista de carpetas
    column_name = "CARPETAS"                        # Columna que trae el prefijo (ej. 0701-0057)
    local_file  = "masivo.pdf"                      # Archivo a subir
    base_folder = "LJC/2025/JUL"           # Carpeta base real en SharePoint

    # Si quieres permitir crear la carpeta cuando no exista, pon create_if_missing=True
    upload_from_excel(excel_path, column_name, local_file, base_folder, create_if_missing=False)
