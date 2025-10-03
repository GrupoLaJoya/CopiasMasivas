from copy_children import *

if __name__ == "__main__":
    entrada = int(input("Ingresa 1 para masivo, 2 para uno a uno: "))
    if entrada == 1:
        process_excel(
            EXCEL_FILE,
            sheet_name=SHEET_NAME,
            col_prefix=COL_PREFIX,
            col_file=COL_FILE,
            col_base=COL_BASE,  # si tu Excel NO tiene la base por fila, deja esta col y usará default_base
            default_base=DEFAULT_BASE_REL_PATH,
            create_missing=CREATE_MISSING,
        )
    else:
        process_excel(
            "detracciones.xlsx",
            sheet_name=SHEET_NAME,
            col_prefix=COL_PREFIX,
            col_file="COMPROBANTE",
            col_base=COL_BASE,  # si tu Excel NO tiene la base por fila, deja esta col y usará default_base
            default_base=DEFAULT_BASE_REL_PATH,
            create_missing=CREATE_MISSING,
            detracciones=True
        )


