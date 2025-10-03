import os
import pandas as pd
import json
import sys
from office365.runtime.auth.client_credential import ClientCredential
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File

# Cargar configuración
with open("config_tenant.json", "r") as f:
    config = json.load(f)

tenant_id = config["sharepoint"]["tenant_id"]
client_id = config["sharepoint"]["client_id"]
client_secret = config["sharepoint"]["client_secret"]
authority = f"https://login.microsoftonline.com/{tenant_id}"
scope = ["https://graph.microsoft.com/.default"]

tenant = config["sharepoint"]["tenant"]
site_url = f"https://{tenant}.sharepoint.com/sites/BacklogTI"
#client_id = config["sharepoint"]["client_id_vendor"]
#client_secret = config["sharepoint"]["client_secret_vendor"]
documento_local = "masivo.pdf"
entrada = int(input("Ingresa 1 para masivo, 2 para uno a uno: "))
if entrada == 1:
    excel_path = "masivo.xlsx"
    nombre_columna_prefijo = "CARPETAS"  # columna con valores tipo 0701-0057

    # === Autenticación con SharePoint ===
    # ctx = ClientContext(site_url).with_credentials(ClientCredential(client_id, client_secret))
    #ctx = ClientContext(site_url).with_user_credentials(client_id, client_secret)
    ctx = ClientContext(site_url).with_credentials(ClientCredential(client_id, client_secret))
    # === Leer Excel ===
    df = pd.read_excel(excel_path)
    library_name = config["sharepoint"]["document_library_prod"] # diccionario de carpetas
    library_folder = config["sharepoint"]["document_library"]
    for i in library_name:
        folder_root = ctx.web.get_folder_by_server_relative_url(f"{library_folder}")
        folders = folder_root.folders.get().execute_query()
        for dir in folders:
            folder_subroot = ctx.web.get_folder_by_server_relative_url(f"{i}/{dir.name}")
            dirs = folder_subroot.folders.get().execute_query()
        # === Buscar cada carpeta por prefijo y copiar el archivo ===
            for prefix in df[nombre_columna_prefijo]:
                matching_folder = next((f for f in dirs if f.name.startswith(str(prefix))), None)

                if matching_folder:
                    target_url = f"{i}/{dir.name}/{matching_folder.name}/{os.path.basename(documento_local)}"
                    with open(documento_local, "rb") as file:
                        ctx.web.get_folder_by_server_relative_url(f"{i}/{dir.name}/{matching_folder.name}")\
                            .upload_file(os.path.basename(documento_local), file.read())\
                            .execute_query()
                    print(f"📁 Copiado en: {matching_folder.name}")
                else:
                    print(f"⚠️ Carpeta no encontrada para prefijo: {prefix}")
else:
    excel_path = "detracciones.xlsx"
    nombre_columna_prefijo = "CARPETAS"  # columna con valores tipo 0701-0057
    nombre_columna_comprobante = "COMPROBANTE"
    ctx = ClientContext(site_url).with_credentials(ClientCredential(client_id, client_secret))

    # === Leer Excel ===
    df = pd.read_excel(excel_path)
    print(df)
    library_name = config["sharepoint"]["document_library"]
    folder_root = ctx.web.get_folder_by_server_relative_url(f"{library_name}")
    token = ctx.authentication_context.acquire_token_for_app(client_id, client_secret)
    print("Access Token:", token.url)
    web = ctx.web.get().execute_query()
    print("Conectado a:", web.properties["Title"])
    lists = ctx.web.lists.get().execute_query()
    for l in lists:
        print(l.properties["Title"])
    folders = folder_root.folders.get().execute_query()
    for dir in folders:
        folder_subroot = ctx.web.get_folder_by_server_relative_url(f"{library_name}/{dir.name}")
        dirs = folder_subroot.folders.get().execute_query()
        for index, row in df.iterrows():
            nombre_carpeta = row[nombre_columna_prefijo]
            nombre_comprobante = f"detracciones/{row[nombre_columna_comprobante]}.pdf"
            documento_local = f"detracciones/"
            matching_folder = next((f for f in dirs if f.name.startswith(str(nombre_carpeta))), None)
            if matching_folder:
                target_url = f"{library_name}/{matching_folder.name}/{os.path.basename(nombre_comprobante)}"
                with open(nombre_comprobante, "rb") as file:
                    ctx.web.get_folder_by_server_relative_url(f"{library_name}/{dir.name}/{matching_folder.name}") \
                        .upload_file(os.path.basename(nombre_comprobante), file.read()) \
                        .execute_query()
                print(f"📁 Copiado en: {matching_folder.name}")
            else:
                print(f"⚠️ Carpeta no encontrada para prefijo: {nombre_carpeta}")