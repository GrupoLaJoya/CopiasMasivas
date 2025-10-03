# extraer_detracciones.py
import re
import json
import unicodedata
from pathlib import Path
from typing import Iterable, List, Tuple, Dict, Optional

import pandas as pd
import pdfplumber
from PIL import Image
from unidecode import unidecode

# -----------------------
# Utilidades de texto
# -----------------------
def norm(s: str) -> str:
    """Normaliza texto para hacer matching robusto (minúsculas, sin tildes)."""
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    return s.lower()

def words_on_line(words: List[dict], y: float, tol: float = 2.0) -> List[dict]:
    """Devuelve palabras cuya línea (top/bottom) está cerca de y."""
    return [w for w in words if abs(w["top"] - y) <= tol or abs(w["bottom"] - y) <= tol]

# -----------------------
# Localización de constancias
# -----------------------
_RE_LABEL = re.compile(r"\bnumero\s*de\s*constancia\b", re.I)

def _only_digits(s: str) -> str:
    return "".join(ch for ch in s if ch.isdigit())

def find_constancia_blocks(
    pdf_path: Path,
    target_numbers: Iterable[str],
    *,
    top_margin: float = 40.0,     # sube un poco para incluir la etiqueta a la izquierda
    gap_margin: float = 10.0      # deja un pequeño espacio antes del siguiente bloque
) -> List[Dict]:
    """
    Devuelve dicts con:
      - 'nro': str         (solo dígitos)
      - 'page_index': int  (0-based)
      - 'bbox': (x0, top, x1, bottom) en coords pdfplumber

    Regla de recorte:
      Ancho completo de página; alto desde el 'nro' actual (un poco más arriba)
      hasta el 'nro' siguiente en la misma página (un poco más arriba de él).
      Si no hay siguiente, recorta hasta el final de la página.
    """
    targets = {_only_digits(str(t)) for t in target_numbers if str(t).strip()}
    results: List[Dict] = []
    if not targets:
        return results

    with pdfplumber.open(str(pdf_path)) as pdf:
        for pidx, page in enumerate(pdf.pages):
            raw_text = page.extract_text() or ""
            ntext = norm(raw_text)

            # Si el label no aparece ni tampoco alguno de los targets en el texto normalizado,
            # igual intentaremos con search() porque algunos PDFs no "pegan" bien el texto.
            # Pero este filtro evita recorrer páginas irrelevantes.
            if not (_RE_LABEL.search(ntext) or any(t in ntext for t in targets)):
                continue

            # 1) Buscar TODAS las apariciones de los targets en la página (con sus rects)
            hits: List[Tuple[float, float, str]] = []  # (top, bottom, nro)
            for nro in targets:
                rects = page.search(nro) or []
                for r in rects:
                    hits.append((r["top"], r["bottom"], nro))

            if not hits:
                continue

            # 2) Ordenar los hits por posición vertical (de arriba hacia abajo)
            hits.sort(key=lambda x: x[0])  # top ascendente

            # 3) Para cada hit, recortar desde su top hacia el siguiente hit.top
            for i, (top, bottom, nro) in enumerate(hits):
                # Subir un poco para incluir la etiqueta (título de la sección)
                top_block = max(top - top_margin, 0.0)

                if i + 1 < len(hits):
                    # siguiente constancia en la misma página
                    next_top = hits[i + 1][0]
                    bottom_block = max(min(next_top - gap_margin, page.height), top_block)
                else:
                    # última constancia detectada en la página -> hasta el final de la página
                    bottom_block = page.height

                # Ancho completo
                x0, x1 = 0.0, page.width
                bbox = (x0, top_block, x1, bottom_block)

                results.append({
                    "nro": nro,
                    "page_index": pidx,
                    "bbox": bbox,
                })

    return results

# -----------------------
# Recorte y exportación
# -----------------------
def crop_to_pdf_and_png(pdf_path: Path, page_index: int, bbox, salida_dir: Path, stem: str):
    """
    Crea dos archivos:
      - <stem>.pdf  (recorte vectorial usando cropbox de PyPDF2)
      - <stem>.png  (render raster del recorte con pdfplumber)
    bbox = (x0, top, x1, bottom) según pdfplumber
    """
    out_pdf = salida_dir / f"{stem}.pdf"
    out_png = salida_dir / f"{stem}.png"

    import pdfplumber
    from PyPDF2 import PdfReader, PdfWriter

    x0, top, x1, bottom = bbox

    # --- PNG con pdfplumber (raster)
    with pdfplumber.open(str(pdf_path)) as pdf:
        page = pdf.pages[page_index]
        # png del recorte
        crop = page.crop((x0, top, x1, bottom))
        # ajusta resolución si quieres más nitidez
        crop.to_image(resolution=200).save(str(out_png), format="PNG")
        page_height = page.height  # lo usamos para transformar coordenadas Y

    # --- PDF vectorial con PyPDF2 usando cropbox
    reader = PdfReader(str(pdf_path))
    pg = reader.pages[page_index]

    # IMPORTANTE: Sistema de coordenadas
    # - pdfplumber bbox: (x0, top, x1, bottom) con "top" medido desde la parte superior
    # - PDF/ PyPDF2: origen abajo-izquierda (y crece hacia arriba)
    # Transformación:
    y0 = page_height - bottom   # lower y
    y1 = page_height - top      # upper y

    # Define cropbox al recorte
    pg.cropbox.lower_left  = (x0, y0)
    pg.cropbox.upper_right = (x1, y1)
    # Opcional: también puedes tocar mediabox/bleed/trim si quieres forzar tamaño de página al recorte
    pg.mediabox.lower_left  = (x0, y0)
    pg.mediabox.upper_right = (x1, y1)

    writer = PdfWriter()
    writer.add_page(pg)
    with out_pdf.open("wb") as f:
        writer.write(f)

    return out_pdf, out_png

# -----------------------
# Lectura de Excel
# -----------------------
def read_constancias_from_excel(xlsx_path: Path, sheet: Optional[str], column_name: str) -> List[str]:
    df = pd.read_excel(xlsx_path, sheet_name=sheet)
    if column_name not in df.columns:
        raise ValueError(f"En el Excel no existe la columna '{column_name}'. Columnas: {list(df.columns)}")
    series = df[column_name].dropna().astype(str).map(lambda s: re.sub(r"\D", "", s))
    return [s for s in series if s.strip()]

# -----------------------
# Orquestador
# -----------------------
def procesar_detracciones(
    xlsx_path: Path,
    sheet: Optional[str],
    column_name: str,
    pdfs_dir: Path,
    salida_dir: Path
):
    """
    - Lee números de constancia de xlsx_path[column_name].
    - Busca en todos los PDFs bajo pdfs_dir los bloques que coincidan.
    - Exporta recortes a salida_dir/<NRO>_{basename_pdf}_{page}.pdf/png
    """
    salida_dir.mkdir(parents=True, exist_ok=True)

    objetivos = read_constancias_from_excel(xlsx_path, sheet, column_name)
    objetivos_set = set(objetivos)

    # Recorremos todos los PDFs (en árbol)
    pdf_paths = [p for p in pdfs_dir.rglob("*.pdf")]
    if not pdf_paths:
        print(f"⚠️  No se encontraron PDFs en {pdfs_dir}")
        return

    print(f"Constancias objetivo: {len(objetivos_set)}")
    print(f"PDFs a revisar: {len(pdf_paths)}")

    encontrados = 0
    for pdf in pdf_paths:
        bloques = find_constancia_blocks(pdf, objetivos_set)
        if not bloques:
            continue
        for b in bloques:
            nro = b["nro"]
            #stem = f"{nro}_{pdf.stem}_p{b['page_index']+1}"
            stem = f"{nro}"
            out_pdf, out_png = crop_to_pdf_and_png(
                pdf, b["page_index"], b["bbox"], salida_dir, stem
            )
            print(f"✅ {nro}: {pdf.name} (p{b['page_index']+1}) → {out_pdf.name}, {out_png.name}")
            encontrados += 1

    if encontrados == 0:
        print("⚠️  No se encontró ninguna constancia que coincida.")
    else:
        print(f"Listo. Recortes generados: {encontrados}")

# -----------------------
# (Opcional) Extraer tabla a CSV con Camelot
# -----------------------
def tabla_a_csv_con_camelot(recorte_pdf: Path, out_csv: Path):
    """
    Si deseas sacar un CSV de la tablita del recorte.
    Requiere: pip install camelot-py[cv]  (y dependencias)
    """
    try:
        import camelot
    except ImportError:
        print("⚠️  camelot no está instalado. Omite extracción CSV.")
        return
    tables = camelot.read_pdf(str(recorte_pdf), pages="1", flavor="lattice")  # prueba también 'stream'
    if tables and len(tables) > 0:
        tables[0].to_csv(str(out_csv))
        print(f"CSV generado: {out_csv.name}")
    else:
        print("⚠️  No se detectó tabla con Camelot en este recorte.")

# -----------------------
# Ejemplo de uso directo
# -----------------------
if __name__ == "__main__":
    # Ajusta estas rutas:
    EXCEL = Path("detracciones.xlsx")     # Excel con columna de números de constancia
    SHEET = "Hoja1"                          # o el nombre de la hoja, p.ej. "Hoja1"
    COL   = "COMPROBANTE"                      # <- nombre exacto de tu columna en el Excel

    PDFS_DIR = Path("detracciones")               # carpeta que contiene los PDFs
    SALIDA   = Path("salida_detracciones")

    # Ejecución:
    procesar_detracciones(EXCEL, SHEET, COL, PDFS_DIR, SALIDA)

    # Si quieres luego convertir algún recorte a CSV:
    # tabla_a_csv_con_camelot(SALIDA / "275378473_miPDF_p1.pdf", SALIDA / "275378473.csv")
