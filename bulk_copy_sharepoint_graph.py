import os
import re
import sys
import json
import argparse
from pathlib import Path
from typing import Dict, List, Optional

import requests
import pandas as pd

# ============ CONFIG ============
# Credenciales (App Registration en Entra ID)
with open("config_tenant.json", "r", encoding="utf-8") as f:
    config = json.load(f)

TENANT_ID     = config["sharepoint"]["tenant_id"]
CLIENT_ID     = config["sharepoint"]["client_id"]
CLIENT_SECRET = config["sharepoint"]["client_secret"]

# Sitio de SharePoint (ajústalo a tu tenant)
SITE_HOSTNAME = os.environ.get("GRAPH_SITE_HOSTNAME", "lajoyaminingsac.sharepoint.com")
# Ruta del sitio (Site relative path). Ej: /sites/Finanzas  o  /teams/Finanzas
SITE_REL_PATH = config["sharepoint"]["site_name"]
DRIVE_NAME = config["sharepoint"]["document_library"]

# Ruta BASE fija que pediste:
BASE_PATH = config["sharepoint"]["base_path"]

# Lista de meses a verificar (en el orden que quieras):
MESES = config["sharepoint"]["lista_meses"]
# Si quieres solo nombres simples, usa: ["ENERO","FEBRERO",...,"DICIEMBRE"]
# =================================


# ---------- Autenticación / llamadas Graph ----------
def graph_token() -> str:
    url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    data = {
        "client_id": CLIENT_ID,
        "scope": "https://graph.microsoft.com/.default",
        "client_secret": CLIENT_SECRET,
        "grant_type": "client_credentials",
    }
    r = requests.post(url, data=data)
    r.raise_for_status()
    return r.json()["access_token"]

def gget(token: str, url: str, params=None):
    r = requests.get(url, headers={"Authorization": f"Bearer {token}"}, params=params)
    r.raise_for_status()
    return r.json()

def gput_upload(token: str, upload_url: str, content: bytes):
    r = requests.put(
        upload_url,
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/octet-stream",
        },
        data=content,
    )
    r.raise_for_status()
    return r.json()

# ---------- Resolución de sitio/drive y navegación ----------
def resolve_site_and_drive(token: str) -> Dict[str, str]:
    # 1) site
    site = gget(
        token,
        f"https://graph.microsoft.com/v1.0/sites/{SITE_HOSTNAME}:{SITE_REL_PATH}"
    )
    site_id = site["id"]

    # 2) drives (bibliotecas) del sitio
    drives = gget(token, f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives")["value"]

    # 3) buscar por nombre exacto (case-insensitive)
    drive = next(
        (d for d in drives if d.get("name","").strip().lower() == DRIVE_NAME.strip().lower()),
        None
    )
    if not drive:
        names = ", ".join(d.get("name","") for d in drives)
        raise RuntimeError(f"No encontré la biblioteca '{DRIVE_NAME}'. Disponibles: {names}")

    return {"site_id": site_id, "drive_id": drive["id"]}

def list_children(token: str, site_id: str, drive_id: str, parent_item_id: str) -> List[Dict]:
    """
    Lista hijos inmediatos de una carpeta por item-id.
    """
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/items/{parent_item_id}/children"
    items = []
    while True:
        j = gget(token, url)
        items.extend(j.get("value", []))
        if "@odata.nextLink" in j:
            url = j["@odata.nextLink"]
        else:
            break
    return items

def get_item_by_path(token: str, site_id: str, drive_id: str, rel_path: str) -> Dict:
    """
    Obtiene un item (carpeta/archivo) por path relativo al drive.
    Si no existe, lanza error.
    """
    rel = "/" + rel_path.strip("/")
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/root:{rel}"
    return gget(token, url)

def ensure_path_exists(token: str, site_id: str, drive_id: str, rel_path: str) -> Dict:
    """
    Navega segmento a segmento y devuelve el item final (no crea nuevas carpetas;
    si quisieras crear, aquí puedes añadir POST a children para crear faltantes).
    """
    rel_path = rel_path.strip("/")
    if not rel_path:
        return gget(token, f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/root")

    segments = rel_path.split("/")
    current = gget(token, f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/root")
    for seg in segments:
        # Buscar hijo con ese nombre exacto
        childs = list_children(token, site_id, drive_id, current["id"])
        nxt = next((c for c in childs if c["name"].strip().lower() == seg.strip().lower() and "folder" in c), None)
        if not nxt:
            raise FileNotFoundError(f"No existe la carpeta: {seg} en {current['name']}")
        current = nxt
    return current  # carpeta final

# ---------- Utilidades de matching ----------
def starts_with_folder(child_name: str, prefix: str) -> bool:
    return child_name.strip().upper().startswith(prefix.strip().upper())

def find_child_folders_by_prefix(token: str, site_id: str, drive_id: str, parent_id: str, prefix: str) -> List[Dict]:
    items = list_children(token, site_id, drive_id, parent_id)
    return [it for it in items if it.get("folder") and starts_with_folder(it.get("name",""), prefix)]

def find_local_file_by_token(src_dir: Path, token: str, ext: Optional[str] = None) -> Optional[Path]:
    token = str(token).strip()
    if not token:
        return None
    patt = re.compile(re.escape(token))
    for p in src_dir.rglob("*"):
        if p.is_file():
            if ext and p.suffix.lower() != ext.lower():
                continue
            if patt.search(p.name):
                return p
    return None

def upload_file_to_folder(token: str, site_id: str, drive_id: str, folder_id: str, local_path: Path) -> Dict:
    upload_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/items/{folder_id}:/{local_path.name}:/content"
    return gput_upload(token, upload_url, local_path.read_bytes())


# ---------- Lógica principal ----------
def process_masiva(token: str, site_id: str, drive_id: str, base_path: str, excel_path: str, same_file: str, sheet: Optional[str], dry: bool):
    base_folder = ensure_path_exists(token, site_id, drive_id, base_path)
    same_file_path = Path(same_file)
    if not same_file_path.exists():
        raise FileNotFoundError(f"No existe el archivo a copiar: {same_file_path}")

    df = pd.read_excel(excel_path, sheet_name=sheet)
    if "CARPETAS" not in df.columns:
        raise ValueError("El Excel debe tener columna 'CARPETAS'")

    total = 0

    # Iterar meses existentes bajo la base
    meses_encontrados = []
    childs_base = list_children(token, site_id, drive_id, base_folder["id"])
    child_names = {c["name"].strip().upper(): c for c in childs_base if c.get("folder")}
    for mes in MESES:
        key = mes.strip().upper()
        if key in child_names:
            meses_encontrados.append(child_names[key])

    for mes_folder in meses_encontrados:
        print(f"↳ Mes: {mes_folder['name']}")
        for _, row in df.iterrows():
            prefix = str(row["CARPETAS"]).strip()
            if not prefix:
                continue
            matches = find_child_folders_by_prefix(token, site_id, drive_id, mes_folder["id"], prefix)
            if not matches:
                # No todas las filas tendrán carpeta en todos los meses; solo avisamos
                print(f"  ⚠️  No hay carpeta que empiece con '{prefix}' en {mes_folder['name']}")
                continue
            for fol in matches:
                if dry:
                    print(f"  [DRY] Copiaría '{same_file_path.name}' → {base_path}/{mes_folder['name']}/{fol['name']}")
                else:
                    up = upload_file_to_folder(token, site_id, drive_id, fol["id"], same_file_path)
                    print(f"  ✅ Copiado '{same_file_path.name}' → {up.get('webUrl')}")
                    total += 1
    print(f"Listo (MASIVA). Archivos subidos: {total}")

def process_detracciones(token: str, site_id: str, drive_id: str, base_path: str, excel_path: str, src_dir: str, sheet: Optional[str], ext: str, dry: bool):
    base_folder = ensure_path_exists(token, site_id, drive_id, base_path)
    df = pd.read_excel(excel_path, sheet_name=sheet)
    required = {"CARPETAS", "COMPROBANTE"}
    if not required.issubset(df.columns):
        raise ValueError("El Excel debe tener columnas 'CARPETAS' y 'COMPROBANTE'")

    src_root = Path(src_dir)
    if not src_root.exists():
        raise FileNotFoundError(f"No existe el directorio de origen: {src_root}")

    total = 0
    childs_base = list_children(token, site_id, drive_id, base_folder["id"])
    child_names = {c["name"].strip().upper(): c for c in childs_base if c.get("folder")}
    meses_encontrados = []
    for mes in MESES:
        key = mes.strip().upper()
        if key in child_names:
            meses_encontrados.append(child_names[key])

    for mes_folder in meses_encontrados:
        print(f"↳ Mes: {mes_folder['name']}")
        for _, row in df.iterrows():
            prefix = str(row["CARPETAS"]).strip()
            nro = str(row["COMPROBANTE"]).strip()
            if not prefix or not nro:
                continue

            matches = find_child_folders_by_prefix(token, site_id, drive_id, mes_folder["id"], prefix)
            if not matches:
                print(f"  ⚠️  Carpeta prefijo '{prefix}' no encontrada en {mes_folder['name']}")
                continue

            f = find_local_file_by_token(src_root, nro, ext=ext)
            if not f:
                print(f"  ⚠️  No se encontró archivo con comprobante '{nro}' en {src_root}")
                continue

            for fol in matches:
                if dry:
                    print(f"  [DRY] Copiaría '{f.name}' → {base_path}/{mes_folder['name']}/{fol['name']}")
                else:
                    up = upload_file_to_folder(token, site_id, drive_id, fol["id"], f)
                    print(f"  ✅ Copiado '{f.name}' → {up.get('webUrl')}")
                    total += 1
    print(f"Listo (DETRACCIONES). Archivos subidos: {total}")


# ---------- CLI ----------
def main():
    parser = argparse.ArgumentParser(description="Copia masiva de archivos a SharePoint (Graph)")
    parser.add_argument("--mode", choices=["masiva","detracciones"], required=True, help="Tipo de proceso")
    parser.add_argument("--excel", required=True, help="Ruta al Excel (XLSX)")
    parser.add_argument("--sheet", default=None, help="Nombre de hoja (opcional)")
    # MASIVA
    parser.add_argument("--same-file", help="Archivo único a copiar (modo masiva)")
    # DETRACCIONES
    parser.add_argument("--src-dir", help="Directorio donde buscar los PDFs (modo detracciones)")
    parser.add_argument("--ext", default=".pdf", help="Extensión a buscar en detracciones (default .pdf)")
    # General
    parser.add_argument("--dry", action="store_true", help="Simular sin subir")
    args = parser.parse_args()

    # Validaciones mínimas
    if args.mode == "masiva" and not args.same_file:
        parser.error("--same-file es requerido en modo 'masiva'")
    if args.mode == "detracciones" and not args.src_dir:
        parser.error("--src-dir es requerido en modo 'detracciones'")

    if not (TENANT_ID and CLIENT_ID and CLIENT_SECRET):
        print("❌ Falta configurar GRAPH_TENANT_ID / GRAPH_CLIENT_ID / GRAPH_CLIENT_SECRET", file=sys.stderr)
        sys.exit(1)

    token = graph_token()
    ids = resolve_site_and_drive(token)
    site_id, drive_id = ids["site_id"], ids["drive_id"]

    if args.mode == "masiva":
        process_masiva(token, site_id, drive_id, BASE_PATH, args.excel, args.same_file, args.sheet, args.dry)
    else:
        process_detracciones(token, site_id, drive_id, BASE_PATH, args.excel, args.src_dir, args.sheet, args.ext, args.dry)


if __name__ == "__main__":
    main()
