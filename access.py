import os, json, base64, requests, msal, sys

TENANT_ID = "5ada65a8-ed0d-481a-9d04-d64e939acb42"
CLIENT_ID = "7de960fa-8155-4bac-b00d-c0951c189d8f"
CLIENT_SECRET = "e1_8Q~HFkNYXC2HfeYlf.tE~4S4W_bYLtaG3PaQC"

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES = ["https://graph.microsoft.com/.default"]  # IMPORTANTE: .default para app-only

app = msal.ConfidentialClientApplication(
    CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET
)
tok = app.acquire_token_for_client(scopes=SCOPES)

if "access_token" not in tok:
    print("No se obtuvo token:", tok)  # aquí verás si hay MFA/CA, invalid_client, etc.
    sys.exit(1)

# DEBUG: inspecciona el token (aud/roles)
payload = tok["access_token"].split(".")[1] + "=="
claims = json.loads(base64.urlsafe_b64decode(payload.encode()))
print("aud:", claims.get("aud"))
print("roles:", claims.get("roles"))   # debe contener Sites.Read.All o Sites.ReadWrite.All
print("scp :", claims.get("scp"))      # en app-only normalmente es None

headers = {"Authorization": f"Bearer {tok['access_token']}"}
url = "https://graph.microsoft.com/v1.0/sites/root?$select=id,webUrl,displayName"
r = requests.get(url, headers=headers)

print("status:", r.status_code)
print("www-authenticate:", r.headers.get("WWW-Authenticate"))
print("body:", r.text)
r.raise_for_status()