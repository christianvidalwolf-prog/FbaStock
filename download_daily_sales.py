"""
download_daily_sales.py
-----------------------
Descarga el informe de ventas de los últimos 7 días de Sellerboard
y hace un append al historial de ventas diario (sales_history.csv).

- Guarda un snapshot diario en: /Volumes/USB SSD/Ficheros sellerboard/
- El historial acumulado también vive en esa carpeta.
- Si el registro del día ya existe, no se duplica (idempotente).
- A partir de 60 días de historial, sync_fba_report.py usará este
  fichero en lugar de 'ventas 60 dias.xlsx'.

Ejecutar diariamente a las 9:00 con cron:
  0 9 * * * /usr/bin/python3 /Users/christianvidalwolf/Stock/download_daily_sales.py >> "/Volumes/USB SSD/Ficheros sellerboard/daily_sales.log" 2>&1
"""

import requests
import pandas as pd
import os
from io import StringIO
from datetime import datetime

# ── Configuración ─────────────────────────────────────────────────────────────
REPORT_URL = (
    "https://app.sellerboard.com/es/automation/reports"
    "?id=a258a124dd524541be35028b6a172013&format=csv"
    "&t=bbc9d347dff7407dbd01c90884f31121"
)

# Disco externo destino
USB_DIR = "/Volumes/USB SSD/Ficheros sellerboard"
HISTORY_FILE = os.path.join(USB_DIR, "sales_history.csv")

TODAY = datetime.now().strftime("%Y-%m-%d")
LOG_MARK = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

# Columnas que nos interesan (el resto se descarta para ahorrar espacio)
KEEP_COLS = [
    "Date", "Marketplace", "ASIN", "SKU", "Name",
    "SalesOrganic", "SalesPPC",
    "UnitsOrganic", "UnitsPPC",
    "UnitsSponsoredProducts", "UnitsSponsoredDisplay",
    "Refunds",
    "GrossProfit", "NetProfit", "Margin",
    "Fulfillment Channel",
]

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
        "AppleWebKit/537.36 Chrome/120 Safari/537.36"
    )
}


def parse_es_number(val):
    """Convierte '1.234,56' → 1234.56"""
    if pd.isna(val):
        return 0.0
    s = str(val).strip().replace(".", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return 0.0


def ensure_usb():
    """Verifica que el USB SSD esté montado antes de continuar."""
    if not os.path.isdir(USB_DIR):
        raise RuntimeError(
            f"USB SSD no encontrado en {USB_DIR}. "
            "Conéctalo antes de ejecutar el script."
        )


def download_report():
    print(f"[{LOG_MARK}] Downloading daily sales report…")
    resp = requests.get(REPORT_URL, headers=HEADERS, timeout=60)
    resp.raise_for_status()

    # Sellerboard devuelve un BOM utf-8; lo eliminamos
    text = resp.content.decode("utf-8-sig")
    df = pd.read_csv(StringIO(text), quotechar='"')

    # Guardar snapshot raw del día en el USB
    raw_path = os.path.join(USB_DIR, f"sellerboard_ventas_{TODAY}.csv")
    df.to_csv(raw_path, index=False)
    print(f"  → Raw snapshot saved: {raw_path}")
    print(f"  → {len(df)} rows downloaded, {df['Date'].nunique()} unique dates")
    return df


def load_history():
    if os.path.exists(HISTORY_FILE):
        return pd.read_csv(HISTORY_FILE)
    return pd.DataFrame()


def save_history(df: pd.DataFrame):
    df.to_csv(HISTORY_FILE, index=False)
    print(f"  → History saved: {len(df)} total rows in {HISTORY_FILE}")


def main():
    ensure_usb()
    raw = download_report()

    # Filtrar columnas de interés (solo las que existen en el CSV)
    available = [c for c in KEEP_COLS if c in raw.columns]
    df_new = raw[available].copy()

    # Normalizar números españoles → float
    for col in ["SalesOrganic", "SalesPPC", "GrossProfit", "NetProfit", "Margin"]:
        if col in df_new.columns:
            df_new[col] = df_new[col].apply(parse_es_number)

    for col in ["UnitsOrganic", "UnitsPPC", "UnitsSponsoredProducts",
                "UnitsSponsoredDisplay", "Refunds"]:
        if col in df_new.columns:
            df_new[col] = pd.to_numeric(
                df_new[col].astype(str).str.replace(",", ""), errors="coerce"
            ).fillna(0).astype(int)

    # Añadir columna de fecha de descarga para trazabilidad
    df_new["downloaded_at"] = LOG_MARK

    # Cargar historial existente y hacer merge/dedup
    history = load_history()

    if history.empty:
        combined = df_new
    else:
        # Dedup por Date + SKU + Marketplace (evita duplicados si se re-ejecuta)
        combined = pd.concat([history, df_new], ignore_index=True)
        combined.drop_duplicates(
            subset=["Date", "SKU", "Marketplace"], keep="last", inplace=True
        )

    save_history(combined)

    # Resumen
    total_units = (
        combined["UnitsOrganic"].sum() + combined["UnitsPPC"].sum()
        if "UnitsOrganic" in combined.columns else 0
    )
    days_covered = combined["Date"].nunique() if "Date" in combined.columns else 0
    print(
        f"  → History covers {days_covered} days | "
        f"{len(combined):,} records | {total_units:,.0f} total units"
    )

    if days_covered >= 60:
        print("  ✅ 60-day threshold reached – sync_fba_report.py will use this file.")
    else:
        print(
            f"  ⏳ {60 - days_covered} more days needed before replacing "
            "'ventas 60 dias.xlsx'."
        )


if __name__ == "__main__":
    main()
