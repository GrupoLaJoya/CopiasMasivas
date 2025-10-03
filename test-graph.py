import requests
import msal
import sys
import os
import json

with open("config_tenant.json", "r") as f:
    config = json.load(f)

TENANT_ID = config["sharepoint"]["tenant_id"]
CLIENT_ID = config["sharepoint"]["client_id"]
CLIENT_SECRET = config["sharepoint"]["client_secret"]
site_name = config["sharepoint"]["site_name"]
site_domain = config["sharepoint"]["site_domain"]
document_library = config["sharepoint"]["document_library"]

SITE_DOMAIN = "lajoyaminingsac.sharepoint.com"   # SOLO host, sin https://
SITE_PATH   = "BacklogTI"                        # lo que va después de /sites/
# Si tu sitio NO está bajo /sites/, ajusta: ej. SITE_PATH = "MiSitio" y usa '/sites/MiSitio'
# (algunas colecciones antiguas usan '/teams/...' — en Graph debes usar /sites/{path} igualmente)

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPE = ["https://graph.microsoft.com/.default"]

app = msal.ConfidentialClientApplication(CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET)
token = app.acquire_token_for_client(scopes=SCOPE)
if "access_token" not in token:
    print("No token:", token)
    sys.exit(1)

headers = {"Authorization": f"Bearer {token['access_token']}"}

def get_json(url):
    r = requests.get(url, headers=headers, timeout=20)
    ok = 200 <= r.status_code < 300
    if not ok:
        print("URL:", url)
        print("STATUS:", r.status_code)
        print("RAW:", r.text)       # <- aquí verás el mensaje real de Graph (403/404/etc.)
        r.raise_for_status()
    return r.json()

# 1) Probar conectividad básica
root = get_json("https://graph.microsoft.com/v1.0/sites/root?$select=id,webUrl,displayName")
print("Root OK:", root.get("webUrl"))

# 2) Obtener site por ruta (lo más estable)
site_url = f"https://graph.microsoft.com/v1.0/sites/{SITE_DOMAIN}:/sites/{SITE_PATH}?$select=id,webUrl,displayName"
site = get_json(site_url)
print("Site OK:", site)

site_id = site["id"]

# 3) (Opcional) Listar bibliotecas para comprobar permisos
drives = get_json(f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives?$select=id,name,webUrl")
print("Drives:", [d["name"] for d in drives.get("value", [])])