import csv
import os
import time
import requests
import json
import pandas as pd
import re
import glob
from deep_translator import GoogleTranslator

# Cache para traducciones (evitar traduzir el mismo título múltiples veces)
_title_cache = {}

# Palabras comunes en otros idiomas que indican que no está en español
_NON_SPANISH_WORDS = {
    "french": [
        "les",
        "les",
        "pour",
        "avec",
        "Décoration",
        "Ornement",
        "métal",
        "murale",
        "lézard",
        "lézards",
        "cuisine",
        "acier",
        "inox",
        "panier",
        "bois",
        "plastique",
        "céramique",
        "figurine",
        "statue",
    ],
    "german": [
        "der",
        "die",
        "das",
        "mit",
        "für",
        "und",
        "aus",
        " Edelstahl",
        "Küche",
        "Metall",
        "Holz",
        "Keramik",
        "Figur",
        "Deko",
    ],
    "italian": [
        "per",
        "con",
        "in",
        "di",
        "metallo",
        "legno",
        "ceramica",
        "cucina",
        "decorazione",
        "figura",
        "statua",
    ],
    "english": [
        "the",
        "and",
        "with",
        "for",
        "in",
        "of",
        "metal",
        "wood",
        "kitchen",
        "decor",
        "decoration",
        "figure",
        "statue",
        "gift",
        "home",
        "garden",
    ],
}


def is_likely_spanish(title):
    """Detecta rápidamente si el título probablemente está en español."""
    title_lower = title.lower()
    # Palabras que claramente NO son español
    non_spanish_indicators = [
        "lézard",
        "ornement",
        "murale",
        "métal",
        "cuisine",
        "acier",
        "inox",  # Francés
        "edelstahl",
        "küche",
        "keramik",
        "figur",
        "deko",
        "aus metall",  # Alemán
        "metallo",
        "legno",
        "ceramica",
        "cucina",
        "decorazione",  # Italiano
        " ornament",
        " metal",
        " wood ",
        " kitchen",
        " decoration",
        " figurine",  # Inglés
    ]
    for word in non_spanish_indicators:
        if word in title_lower:
            return False
    return True


def translate_to_spanish(title):
    """Traduce el título al español solo si no parece estar en español."""
    if not title or len(title) < 10:
        return title

    # Si ya está en cache, devolverlo
    if title in _title_cache:
        return _title_cache[title]

    # Rápido: primero verificar si probablemente ya está en español
    if is_likely_spanish(title):
        return title

    # Solo traducir si no parece español
    try:
        translated = GoogleTranslator(source="auto", target="es").translate(title)
        if translated and translated != title:
            _title_cache[title] = translated
            return translated
    except Exception as e:
        pass

    return title

    # Si ya está en cache, devolverlo
    if title in _title_cache:
        return _title_cache[title]

    try:
        # Detectar idioma y traducir si no es español
        # GoogleTranslator detectará automáticamente el idioma origen
        translated = GoogleTranslator(source="auto", target="es").translate(title)
        if translated and translated != title:
            _title_cache[title] = translated
            return translated
    except Exception as e:
        print(f"Translation error: {e}")

    return title


def safe_f(val):
    if not val or val == "nan":
        return 0.0
    try:
        return float(str(val).replace(",", "").strip())
    except:
        return 0.0


# Configuration
CSV_URL = "https://app.sellerboard.com/es/automation/reports?id=a1a2f4284b8043c39964edfe3cef86ca&format=csv&t=bbc9d347dff7407dbd01c90884f31121"
OUTPUT_JSON = "/Users/christianvidalwolf/Stock/fba-replenishment/public/data.json"
WORK_DIR = "/Users/christianvidalwolf/Stock"
SELLERBOARD_DIR = "/Volumes/USB SSD/Ficheros sellerboard"

# Sales data URLs from SellerBoard
SALES_URL = "https://app.sellerboard.com/es/automation/reports?id=a258a124dd524541be35028b6a172013&format=csv&t=bbc9d347dff7407dbd01c90884f31121"

VENTAS_FILE = "/Users/christianvidalwolf/Stock/Ventas 365.xlsx"
VENTAS_60_FILE = "/Users/christianvidalwolf/Stock/ventas 60 dias.xlsx"

PROVIDERS = {
    "dcasa": {
        "url": "https://dcasa.es/DataWeb/DataWeb20260503.csv",
        "type": "csv",
        "id_col": "CODIGO",
        "stock_col": "STOCK_DISPONIBLE",
        "key": "DC",
        "sep": ";",
        "encoding": "latin1",
        "local_fallback": None,
        "id_prefix": None,
    },
    "signes": {
        # URL actualizada mayo-2026 (signesconexion.com); fallback al fichero local si falla
        "url": "https://signesconexion.com/stock/STOCK-44880.CSV",
        "type": "csv",
        "id_col": "Codigo",
        "stock_col": "Stock",
        "key": "SG",
        "sep": ";",
        "encoding": "latin1",
        "local_fallback": "/Users/christianvidalwolf/Stock/signes_stock.csv",
        "id_prefix": "SG-",  # El CSV tiene 'SG-10070'; hay que quitarlo para obtener '10070'
    },
    "minerales": {
        "url": "https://vivescortadaimport.com/modules/doofinder/feed2.php?language=ES&currency=EUR",
        "type": "pipe",
        "id_col": 0,
        "stock_col": 7,
        "key": "VC",
        "sep": "|",
        "encoding": "latin1",
        "local_fallback": "/Users/christianvidalwolf/Stock/minerales_feed.xml",
        "id_prefix": None,
    },
}

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
}


def get_latest_dataweb_file():
    files = sorted(glob.glob(os.path.join(WORK_DIR, "DataWeb*.csv")))
    for path in reversed(files):
        if os.path.getsize(path) > 0:
            return path
    return None


def load_pipe_stock_file(path, stock_col, encoding):
    p_stocks = {}
    with open(path, "r", encoding=encoding, errors="ignore") as f:
        for line in f:
            parts = line.split("|")
            if len(parts) <= stock_col:
                continue
            p_id = parts[0].strip().upper()
            if p_id in {"ID", "ID-ID"}:
                continue
            try:
                p_stock = int(float(parts[stock_col].strip()))
            except:
                p_stock = 0
            p_stocks[p_id] = p_stock
    return p_stocks


def download_supplier_data():
    stocks = {}
    for p_name, config in PROVIDERS.items():
        print(f"Downloading {p_name} data...")
        try:
            df = None
            sep = config.get("sep", ";")
            enc = config.get("encoding", "latin1")
            local_fallback = config.get("local_fallback")

            if p_name == "dcasa":
                # Dcasa: use the newest valid DataWeb snapshot from the daily download.
                local_file = get_latest_dataweb_file()
                if local_file:
                    print(f"  Using local DCASA file: {os.path.basename(local_file)}")
                    df = pd.read_csv(
                        local_file, sep=sep, encoding=enc, on_bad_lines="skip"
                    )
                else:
                    local_file = os.path.join(WORK_DIR, os.path.basename(config["url"]))
                    response = requests.get(config["url"], headers=HEADERS, timeout=30)
                    with open(local_file, "wb") as f:
                        f.write(response.content)
                    df = pd.read_csv(
                        local_file, sep=sep, encoding=enc, on_bad_lines="skip"
                    )
            else:
                # Daily update refreshes these local files first; use them as source of truth.
                if local_fallback and os.path.exists(local_fallback):
                    print(f"  Using local {p_name} file: {local_fallback}")
                    if config.get("type") == "pipe":
                        p_stocks = load_pipe_stock_file(
                            local_fallback, config["stock_col"], enc
                        )
                        stocks[config["key"]] = p_stocks
                        print(f"Loaded {len(p_stocks)} {p_name} items from local file.")
                        continue
                    df = pd.read_csv(
                        local_fallback,
                        sep=sep,
                        encoding=enc,
                        on_bad_lines="skip",
                    )
                else:
                    try:
                        response = requests.get(
                            config["url"], headers=HEADERS, timeout=30
                        )
                        if (
                            response.status_code != 200
                            or "<!DOCTYPE" in response.text[:200]
                        ):
                            raise ValueError(f"URL returned {response.status_code}")

                        if config.get("type") == "pipe":
                            from io import StringIO

                            temp_path = StringIO(response.text)
                            p_stocks = {}
                            for line in temp_path:
                                parts = line.split("|")
                                if len(parts) <= config["stock_col"]:
                                    continue
                                p_id = parts[0].strip().upper()
                                if p_id in {"ID", "ID-ID"}:
                                    continue
                                try:
                                    p_stock = int(
                                        float(parts[config["stock_col"]].strip())
                                    )
                                except:
                                    p_stock = 0
                                p_stocks[p_id] = p_stock
                            stocks[config["key"]] = p_stocks
                            print(f"Loaded {len(p_stocks)} {p_name} items from URL.")
                            continue

                        from io import StringIO

                        df = pd.read_csv(
                            StringIO(response.text),
                            sep=sep,
                            encoding=enc,
                            on_bad_lines="skip",
                        )
                    except Exception as url_err:
                        if local_fallback and os.path.exists(local_fallback):
                            print(
                                f"  URL error ({url_err}), using local fallback: {local_fallback}"
                            )
                            if config.get("type") == "pipe":
                                p_stocks = load_pipe_stock_file(
                                    local_fallback, config["stock_col"], enc
                                )
                                stocks[config["key"]] = p_stocks
                                print(
                                    f"Loaded {len(p_stocks)} {p_name} items from fallback."
                                )
                                continue
                            df = pd.read_csv(
                                local_fallback,
                                sep=sep,
                                encoding=enc,
                                on_bad_lines="skip",
                            )
                        else:
                            raise

            print(f"Loaded {len(df)} {p_name} items.")

            id_prefix = config.get("id_prefix") or ""
            p_stocks = {}
            for _, row in df.iterrows():
                raw_id = str(row.get(config["id_col"], "")).strip().upper()
                # Strip provider prefix (e.g. 'SG-10070' → '10070')
                if id_prefix and raw_id.startswith(id_prefix.upper()):
                    raw_id = raw_id[len(id_prefix) :]
                if raw_id.endswith(".0"):
                    raw_id = raw_id[:-2]
                if not raw_id:
                    continue
                try:
                    p_stock = int(
                        float(str(row.get(config["stock_col"], 0)).replace(",", "."))
                    )
                except:
                    p_stock = 0
                p_stocks[raw_id] = p_stock
            stocks[config["key"]] = p_stocks
        except Exception as e:
            print(f"Error downloading {p_name}: {e}")
            stocks[config["key"]] = {}
    return stocks


def get_provider_and_id(sku):
    """
    Detecta el proveedor y extrae el ID numerico del SKU.
    Mapeo correcto (verificado contra ficheros de stock):
      CLM            -> DC  (Dcasa: refs de 7 digitos como 2684315)
      SG/SGR/SGRG/SGAZ -> SG  (Signes: refs de 5 digitos como 31442)
      VC/VCT         -> VC  (Minerales)
    """
    sku = str(sku).upper().strip()
    # Orden: patrones mas especificos primero
    patterns = [
        (r"^(\d+)CLM", "DC"),
        (r"^(\d+)SGAZ", "SG"),
        (r"^(\d+)SGRG", "SG"),
        (r"^(\d+)SGR", "SG"),
        (r"^(\d+)SGFBA", "SG"),
        (r"^(\d+)SG", "SG"),
        (r"^(\d+)VCFBA", "VC"),
        (r"^(\d+)VCT", "VC"),
        (r"^(\d+)VC", "VC"),
        (r"^(\d+)DC", "DC"),
    ]
    for pattern, prov_key in patterns:
        m = re.match(pattern, sku)
        if m:
            return prov_key, m.group(1)
    return None, None


def normalize_sku(sku):
    return re.sub(r"(0|1|A|RG|I)?FBA$", "", str(sku).upper().strip())


def find_latest_sales_file(prefix):
    """Find the most recent sales file in USB SSD."""
    import glob

    pattern = f"{SELLERBOARD_DIR}/{prefix}_*.csv"
    files = sorted(glob.glob(pattern), reverse=True)
    return files[0] if files else None


def get_latest_sellerboard_file(prefix):
    """Return the most recent SellerBoard snapshot (not just today's)."""
    import glob

    pattern = f"{SELLERBOARD_DIR}/{prefix}_*.csv"
    files = sorted(glob.glob(pattern), reverse=True)
    for f in files:
        if os.path.exists(f):
            return f
    return None


def read_local_sellerboard_file(path):
    with open(path, "r", encoding="utf-8-sig", errors="replace") as f:
        return f.read()


def get_sellerboard_report(prefix, url):
    """Use local snapshot if < 12h old; otherwise download fresh."""
    os.makedirs(SELLERBOARD_DIR, exist_ok=True)
    local_file = get_latest_sellerboard_file(prefix)
    if local_file:
        age_hours = (time.time() - os.path.getmtime(local_file)) / 3600
        if age_hours < 12:
            print(f"Using local SellerBoard snapshot ({age_hours:.1f}h old): {local_file}")
            return read_local_sellerboard_file(local_file), local_file
        print(f"Local snapshot too old ({age_hours:.1f}h). Downloading fresh...")

    timestamp = pd.Timestamp.now().strftime("%Y-%m-%d")
    output_file = f"{SELLERBOARD_DIR}/{prefix}_{timestamp}.csv"
    print(f"Downloading {prefix} from SellerBoard...")
    response = requests.get(url, headers=HEADERS, timeout=60)
    response.raise_for_status()
    text = response.content.decode("utf-8-sig", errors="replace")
    with open(output_file, "w", encoding="utf-8") as f:
        f.write(text)
    print(f"Saved {prefix} to {output_file}")
    return text, output_file


def download_and_save_sales():
    """Download today's sales from SellerBoard and save to USB SSD."""
    try:
        _, sales_file = get_sellerboard_report("sellerboard_ventas", SALES_URL)
        return sales_file
    except Exception as e:
        print(f"Exception downloading sales: {e}")
    return None


def get_sales_data():
    """Build sales history: prioritize Excel files (60 & 365 days), fallback to CSV."""
    import glob
    from datetime import datetime, timedelta

    sales_365 = {}
    sales_60 = {}

    # First: Read Excel files (they have full 60/365 day history)
    if os.path.exists(VENTAS_FILE):
        print(f"Reading {VENTAS_FILE}...")
        try:
            df = pd.read_excel(VENTAS_FILE)
            df["ASIN"] = df["ASIN"].astype(str).str.strip()
            sales_365 = df.groupby("ASIN")["Units"].sum().to_dict()
            print(f"  365-day sales from Excel: {len(sales_365)} ASINs")
        except Exception as e:
            print(f"Error reading {VENTAS_FILE}: {e}")

    if os.path.exists(VENTAS_60_FILE):
        print(f"Reading {VENTAS_60_FILE}...")
        try:
            df = pd.read_excel(VENTAS_60_FILE)
            df["ASIN"] = df["ASIN"].astype(str).str.strip()
            sales_60 = df.groupby("ASIN")["Units"].sum().to_dict()
            print(f"  60-day sales from Excel: {len(sales_60)} ASINs")
        except Exception as e:
            print(f"Error reading {VENTAS_60_FILE}: {e}")

    # Second: If no Excel data, try CSV files from USB
    if not sales_365 and not sales_60:
        sales_files = sorted(glob.glob(f"{SELLERBOARD_DIR}/sellerboard_ventas_*.csv"))
        print(f"Found {len(sales_files)} sales files in USB")

        if sales_files:
            all_data = []
            for f in sales_files:
                try:
                    date_str = f.split("_")[-1].replace(".csv", "")
                    file_date = datetime.strptime(date_str, "%Y-%m-%d")

                    df = pd.read_csv(f)
                    df["Date"] = pd.to_datetime(
                        df["Date"], format="%d/%m/%Y", errors="coerce"
                    )
                    df["ASIN"] = df["ASIN"].astype(str).str.strip()
                    all_data.append(df)
                    print(f"  Loaded {f} ({len(df)} rows)")
                except Exception as e:
                    print(f"  Error reading {f}: {e}")

            if all_data:
                combined = pd.concat(all_data, ignore_index=True)
                unit_cols = [col for col in combined.columns if col.startswith("Units")]
                combined["TotalUnits"] = (
                    combined[unit_cols]
                    .apply(pd.to_numeric, errors="coerce")
                    .sum(axis=1)
                )

                cutoff_60 = datetime.now() - timedelta(days=60)
                df_60 = combined[combined["Date"] >= cutoff_60]
                sales_60 = df_60.groupby("ASIN")["TotalUnits"].sum().to_dict()
                print(f"Total 60-day ASINs: {len(sales_60)}")

                cutoff_365 = datetime.now() - timedelta(days=365)
                df_365 = combined[combined["Date"] >= cutoff_365]
                sales_365 = df_365.groupby("ASIN")["TotalUnits"].sum().to_dict()
                print(f"Total 365-day ASINs: {len(sales_365)}")

    return sales_365, sales_60


def sync():
    supplier_stocks = download_supplier_data()

    # Download and save today's sales to build historical record
    today_file = download_and_save_sales()

    # Build sales history from all available CSV files
    sales_365_map, sales_60_map = get_sales_data()

    print(f"Fetching FBA report...")
    try:
        inventory_text, inventory_file = get_sellerboard_report(
            "sellerboard_inventory", CSV_URL
        )
        print(f"Reading inventory from {inventory_file}")

        reader = csv.DictReader(inventory_text.splitlines())
        rows = list(reader)

        # ── ASIN-based lookup maps (built from the FULL inventory, not just FBA rows) ──
        # sku_to_asin: every SKU in the report → its ASIN
        sku_to_asin = {
            row.get("SKU", "").upper(): row.get("ASIN", "")
            for row in rows
            if row.get("SKU") and row.get("ASIN")
        }
        # asin_with_fba: ASINs that already have at least 1 SKU containing 'FBA'
        asin_with_fba = {asin for sku, asin in sku_to_asin.items() if "FBA" in sku}

        # Mapping for titles
        sku_to_title = {
            row.get("SKU", "").upper(): row.get("Title", "")
            for row in rows
            if row.get("SKU") and row.get("Title")
        }

        # Legacy fallback set (for FBM SKUs not present in the inventory report)
        inventory_fba_skus = {sku for sku in sku_to_asin if "FBA" in sku}
        sales_fba_skus = {s.upper() for s in sales_365_map if "FBA" in s.upper()}
        all_fba_skus = sales_fba_skus.union(inventory_fba_skus)

        data = []
        asin_total_stock = {}  # Para verificar si algún SKU del ASIN ya tiene stock
        asin_has_fba_stock = {}  # Se填充 después

        # Primera pasada: acumular stock por ASIN
        for row in rows:
            sku = row.get("SKU", "").upper()
            if not sku.endswith("FBA"):
                continue

            stock_amz = safe_f(row.get("FBA/FBM Stock", "0"))
            asin = row.get("ASIN", "").strip().upper()
            if not asin or asin == "NAN":
                continue
            if asin not in asin_total_stock:
                asin_total_stock[asin] = 0
            asin_total_stock[asin] += stock_amz

        # Crear set de ASINs que ya tienen stock
        asins_with_stock = {
            asin for asin, total in asin_total_stock.items() if total > 0
        }

        # Segunda pasada: generar datos
        for row in rows:
            sku = row.get("SKU", "").upper()
            if not sku.endswith("FBA"):
                continue

            prov_key, clean_id = get_provider_and_id(sku)
            provider_name = {"SG": "Signes", "VC": "Minerales", "DC": "Dcasa"}.get(
                prov_key, "Unknown"
            )
            supp_stock = supplier_stocks.get(prov_key, {}).get(clean_id, 0)

            stock_amz = safe_f(row.get("FBA/FBM Stock", "0"))
            amazon_rec = safe_f(row.get("Recommended quantity for  reordering", "0"))
            sent_to_fba = safe_f(row.get("Sent  to FBA", "0"))
            reserved = safe_f(row.get("Reserved", "0"))
            days_left = safe_f(row.get("Days  of stock  left", "0"))
            roi = safe_f(row.get("ROI, %", "0"))

            asin = row.get("ASIN", "").strip().upper()
            if not asin or asin == "NAN":
                asin = f"MISSING_{idx}"

            transit = sent_to_fba + reserved
            effective_stock = stock_amz + transit

            # REGLA: Si el ASIN ya tiene stock en cualquier SKU, no recomendar reorder
            # a menos que el tránsito sea insuficiente
            if asin in asins_with_stock:
                # El ASIN ya tiene stock, pero verificamos si necesitamos más según tránsito
                needed = max(0, amazon_rec - transit)
                final_rec = min(needed, supp_stock) if needed > 0 else 0
            else:
                # El ASIN no tiene stock, usar lógica normal
                calculated_need = amazon_rec
                if stock_amz < 3:
                    calculated_need = max(amazon_rec, 5)
                final_rec = min(max(0, calculated_need - transit), supp_stock)

            sales_365 = sales_365_map.get(asin, 0)
            sales_60 = sales_60_map.get(asin, 0)

            is_back_in_stock = (
                (stock_amz == 0)
                and (sales_60 == 0)
                and (sales_365 > 8)
                and (supp_stock > 0)
            )
            # Nueva lógica: Stock AMZ > 0 Y (Sin ventas en 60 días O días de stock > 90)
            is_slow_moving = stock_amz > 0 and (sales_60 == 0 or days_left > 90)

            data.append(
                {
                    "asin": asin,
                    "sku": sku,
                    "title": translate_to_spanish(row.get("Title", "")),
                    "roi": roi,
                    "stock_amz": stock_amz,
                    "velocity": safe_f(row.get("Estimated Sales Velocity", "0")),
                    "days_left": days_left,
                    "final_rec": final_rec,
                    "supp_stock": supp_stock,
                    "sent_to_fba": sent_to_fba,
                    "reserved": reserved,
                    "effective_stock": effective_stock,
                    "provider": provider_name,
                    "status": "critical"
                    if (days_left <= 7 or effective_stock == 0)
                    else ("warning" if days_left <= 15 else "ok"),
                    "sales_365": sales_365,
                    "sales_60": sales_60,
                    "is_back_in_stock": is_back_in_stock,
                    "is_slow_moving": is_slow_moving,
                }
            )

        # FBM→FBA recommendations: ASINs con ventas pero sin listing FBA todavía
        fbm_recommendations = []
        for asin, units in sales_365_map.items():
            if units <= 8:
                continue
            if not asin or asin == "nan":
                continue

            # Si el ASIN ya tiene algún SKU con 'FBA' en el inventario → excluir
            if asin in asin_with_fba:
                continue

            # Buscar el SKU base (FBM) para este ASIN para obtener info del proveedor
            # Usamos el primer SKU que encontremos para ese ASIN que no sea FBA
            base_sku = next(
                (s for s, a in sku_to_asin.items() if a == asin and "FBA" not in s),
                None,
            )
            if not base_sku:
                continue  # Si no tenemos el SKU en el reporte, no podemos mapear proveedor

            prov_key, clean_id = get_provider_and_id(base_sku)
            if not prov_key:
                continue

            fbm_recommendations.append(
                {
                    "asin": asin,
                    "sku": base_sku,
                    "title": translate_to_spanish(sku_to_title.get(base_sku, "")),
                    "sales_365": units,
                    "provider": {"SG": "Signes", "VC": "Minerales", "DC": "Dcasa"}.get(
                        prov_key, "Unknown"
                    ),
                }
            )

        fbm_recommendations.sort(key=lambda x: x["sales_365"], reverse=True)

        final_data = {
            "summary": {
                "total_skus": len(data),
                "critical_count": len([p for p in data if p["status"] == "critical"]),
                "warning_count": len([p for p in data if p["status"] == "warning"]),
                "out_of_supplier_stock": len([p for p in data if p["supp_stock"] == 0]),
                "back_in_stock_count": len([p for p in data if p["is_back_in_stock"]]),
                "slow_moving_count": len([p for p in data if p["is_slow_moving"]]),
                "fbm_rec_count": len(fbm_recommendations),
                "last_update": pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S"),
            },
            "products": data,
            "fbm_recommendations": fbm_recommendations,
        }
        with open(OUTPUT_JSON, "w") as f:
            json.dump(final_data, f, indent=2)
        print(
            f"Synced {len(data)} SKUs. FBM Recs: {len(fbm_recommendations)}. Slow Moving: {final_data['summary']['slow_moving_count']}"
        )

    except Exception as e:
        print(f"Error: {e}")


if __name__ == "__main__":
    sync()
