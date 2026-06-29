import csv
import os
import sys
import time
import fcntl
import signal
import subprocess
import requests
import json
import pandas as pd
import numpy as np
import re
import glob
from datetime import datetime, timedelta
from deep_translator import GoogleTranslator

# Prevent duplicate concurrent runs (two LaunchAgents fire at same time)
_LOCK_FILE = "/tmp/sync_fba_report.lock"
_LOCK_MAX_RUNTIME_SECONDS = 6 * 60 * 60


def _process_command(pid):
    try:
        result = subprocess.run(
            ["ps", "-p", str(pid), "-o", "command="],
            capture_output=True,
            text=True,
            check=False,
        )
        return result.stdout.strip()
    except Exception:
        return ""


def _process_age_seconds(pid):
    try:
        result = subprocess.run(
            ["ps", "-p", str(pid), "-o", "etimes="],
            capture_output=True,
            text=True,
            check=False,
        )
        if result.returncode == 0 and result.stdout.strip():
            return int(result.stdout.strip())
    except Exception:
        pass

    try:
        result = subprocess.run(
            ["ps", "-p", str(pid), "-o", "lstart="],
            capture_output=True,
            text=True,
            check=False,
        )
        if result.returncode != 0 or not result.stdout.strip():
            return 0
        started_text = " ".join(result.stdout.split())
        started = datetime.strptime(started_text, "%a %b %d %H:%M:%S %Y")
        return max(0, int((datetime.now() - started).total_seconds()))
    except Exception:
        return 0


def _lock_holder_pids():
    try:
        result = subprocess.run(
            ["lsof", "-t", _LOCK_FILE],
            capture_output=True,
            text=True,
            check=False,
        )
    except Exception:
        return []

    pids = []
    for value in result.stdout.splitlines():
        try:
            pid = int(value.strip())
        except ValueError:
            continue
        if pid != os.getpid():
            pids.append(pid)
    return pids


def _is_our_stale_sync(pid):
    command = _process_command(pid)
    age = _process_age_seconds(pid)
    return "sync_fba_report.py" in command and age > _LOCK_MAX_RUNTIME_SECONDS


def _terminate_stale_holders():
    stale_pids = [pid for pid in _lock_holder_pids() if _is_our_stale_sync(pid)]
    if not stale_pids:
        return False

    for pid in stale_pids:
        print(
            f"Stale sync_fba_report.py process detected (pid {pid}, "
            f"age {_process_age_seconds(pid)}s). Terminating it."
        )
        try:
            os.kill(pid, signal.SIGTERM)
        except ProcessLookupError:
            continue

    deadline = time.time() + 10
    while time.time() < deadline:
        if not any(_process_command(pid) for pid in stale_pids):
            return True
        time.sleep(0.5)

    for pid in stale_pids:
        if _process_command(pid):
            print(f"Stale process {pid} did not stop after SIGTERM. Killing it.")
            try:
                os.kill(pid, signal.SIGKILL)
            except ProcessLookupError:
                pass

    time.sleep(1)
    return True


def _acquire_lock():
    os.makedirs(os.path.dirname(_LOCK_FILE), exist_ok=True)
    lock_fd = open(_LOCK_FILE, "a+")
    for attempt in range(2):
        try:
            fcntl.flock(lock_fd, fcntl.LOCK_EX | fcntl.LOCK_NB)
            lock_fd.seek(0)
            lock_fd.truncate()
            lock_fd.write(f"pid={os.getpid()}\nstarted_at={time.time()}\n")
            lock_fd.flush()
            return lock_fd
        except BlockingIOError:
            if attempt == 0 and _terminate_stale_holders():
                continue
            holders = ", ".join(
                f"{pid} ({_process_age_seconds(pid)}s)" for pid in _lock_holder_pids()
            )
            print(f"Another instance is already running. Exiting. Holders: {holders}")
            sys.exit(0)
    return lock_fd


_lock_fd = _acquire_lock()

# Cache para traducciones (evitar traduzir el mismo título múltiples veces)
_title_cache = {}
_TRANSLATION_TIMEOUT_SECONDS = 10
_ONLINE_TRANSLATION_ENABLED = os.environ.get("ENABLE_TITLE_TRANSLATION") == "1"
_translation_disabled_logged = False


class TranslationTimeout(Exception):
    pass


def _translation_timeout_handler(signum, frame):
    raise TranslationTimeout()


def _translate_with_timeout(title):
    old_handler = signal.getsignal(signal.SIGALRM)
    original_requests_get = requests.get

    def requests_get_with_timeout(*args, **kwargs):
        kwargs.setdefault("timeout", _TRANSLATION_TIMEOUT_SECONDS)
        return original_requests_get(*args, **kwargs)

    requests.get = requests_get_with_timeout
    signal.signal(signal.SIGALRM, _translation_timeout_handler)
    signal.alarm(_TRANSLATION_TIMEOUT_SECONDS)
    try:
        return GoogleTranslator(source="auto", target="es").translate(title)
    finally:
        signal.alarm(0)
        signal.signal(signal.SIGALRM, old_handler)
        requests.get = original_requests_get

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
    global _translation_disabled_logged

    if not title or len(title) < 10:
        return title

    # Si ya está en cache, devolverlo
    if title in _title_cache:
        return _title_cache[title]

    # Rápido: primero verificar si probablemente ya está en español
    if is_likely_spanish(title):
        return title

    if not _ONLINE_TRANSLATION_ENABLED:
        if not _translation_disabled_logged:
            print(
                "Online title translation disabled during sync "
                "(set ENABLE_TITLE_TRANSLATION=1 to enable)."
            )
            _translation_disabled_logged = True
        return title

    # Solo traducir si no parece español
    try:
        translated = _translate_with_timeout(title)
        if translated and translated != title:
            _title_cache[title] = translated
            return translated
    except TranslationTimeout:
        print(f"Translation timeout after {_TRANSLATION_TIMEOUT_SECONDS}s: {title[:80]}")
    except Exception as e:
        print(f"Translation error: {e}")

    return title


def safe_f(val):
    if not val or val == "nan":
        return 0.0
    try:
        clean_val = str(val).replace("\xa0", "").replace(" ", "").strip()
        if "," in clean_val:
            if "." in clean_val:
                clean_val = clean_val.replace(".", "")
            clean_val = clean_val.replace(",", ".")
        return float(clean_val)
    except:
        return 0.0


# Configuration
CSV_URL = "https://app.sellerboard.com/es/automation/reports?id=a1a2f4284b8043c39964edfe3cef86ca&format=csv&t=bbc9d347dff7407dbd01c90884f31121"
OUTPUT_JSON = "/Users/christianvidalwolf/Stock/fba-replenishment/public/data.json"
OUTPUT_JSON_DIST = "/Users/christianvidalwolf/Stock/fba-replenishment/dist/data.json"
WORK_DIR = "/Users/christianvidalwolf/Stock"
LOCAL_SELLERBOARD_DIR = os.path.join(WORK_DIR, "sellerboard_backups")
USB_SELLERBOARD_DIR = "/Volumes/USB SSD/Ficheros sellerboard"
SELLERBOARD_DIRS = [LOCAL_SELLERBOARD_DIR, USB_SELLERBOARD_DIR]
RETENTION_DAYS = 60
RETENTION_SECONDS = RETENTION_DAYS * 24 * 60 * 60

# Sales data URLs from SellerBoard
SALES_URL = "https://app.sellerboard.com/es/automation/reports?id=a258a124dd524541be35028b6a172013&format=csv&t=bbc9d347dff7407dbd01c90884f31121"

VENTAS_FILE = "/Users/christianvidalwolf/Stock/Ventas 365.xlsx"
VENTAS_60_FILE = "/Users/christianvidalwolf/Stock/ventas 60 dias.xlsx"
SALES_HISTORY_FILE = os.path.join(LOCAL_SELLERBOARD_DIR, "sales_history.csv")
USB_SALES_HISTORY_FILE = os.path.join(USB_SELLERBOARD_DIR, "sales_history.csv")

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


def load_existing_titles_by_sku():
    """Reuse titles already written to data.json so sync never depends on translation."""
    if not os.path.exists(OUTPUT_JSON):
        return {}

    try:
        with open(OUTPUT_JSON, "r", encoding="utf-8") as f:
            current_data = json.load(f)
    except Exception as exc:
        print(f"Could not read existing titles from data.json: {exc}")
        return {}

    titles = {}
    for section in ("products", "fbm_recommendations"):
        for item in current_data.get(section, []):
            sku = str(item.get("sku", "")).strip().upper()
            title = item.get("title", "")
            if sku and title:
                titles[sku] = title
    return titles


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
      MD/MDFBA/MDRG/MDCFBA -> MD  (Trediser)
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
        (r"^(\d+)(?:[A-Z]*)MD(?:RGFBA|CFBA|FBA|RG)?$", "MD"),
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
    """Find the most recent sales file in any SellerBoard snapshot folder."""
    files = get_sellerboard_snapshot_files(prefix)
    return files[0] if files else None


def get_sellerboard_snapshot_files(prefix):
    """Return unique SellerBoard snapshots from the local folder and USB mirror."""
    import glob

    latest_by_name = {}
    for directory in SELLERBOARD_DIRS:
        pattern = os.path.join(directory, f"{prefix}_*.csv")
        for path in glob.glob(pattern):
            if not os.path.exists(path):
                continue
            # Skip invalid/empty files (corrupt or "Report not ready" responses)
            if os.path.getsize(path) < 10000:
                continue
            name = os.path.basename(path)
            current = latest_by_name.get(name)
            if current is None or os.path.getmtime(path) > os.path.getmtime(current):
                latest_by_name[name] = path

    return sorted(
        latest_by_name.values(), key=lambda f: (os.path.getmtime(f), f), reverse=True
    )


def get_latest_sellerboard_file(prefix):
    """Return the most recent SellerBoard snapshot from any local folder."""
    files = get_sellerboard_snapshot_files(prefix)
    for f in files:
        if os.path.exists(f):
            return f
    return None


def save_sellerboard_snapshot(prefix, date_str, content):
    """Write a SellerBoard snapshot to local disk first, then mirror to USB."""
    os.makedirs(LOCAL_SELLERBOARD_DIR, exist_ok=True)
    destinations = [LOCAL_SELLERBOARD_DIR]
    if os.path.isdir(USB_SELLERBOARD_DIR):
        destinations.append(USB_SELLERBOARD_DIR)

    last_error = None
    for directory in destinations:
        try:
            output_file = os.path.join(directory, f"{prefix}_{date_str}.csv")
            with open(output_file, "w", encoding="utf-8") as f:
                f.write(content)
            return output_file
        except OSError as exc:
            last_error = exc
            print(f"Could not save {prefix} to {directory}: {exc}")

    raise last_error or RuntimeError(f"Could not save SellerBoard snapshot for {prefix}.")


def read_local_sellerboard_file(path):
    with open(path, "r", encoding="utf-8-sig", errors="replace") as f:
        return f.read()


def cleanup_old_sellerboard_snapshots():
    cutoff = time.time() - RETENTION_SECONDS
    patterns = ["sellerboard_inventory_*.csv", "sellerboard_ventas_*.csv"]
    for directory in SELLERBOARD_DIRS:
        for pattern in patterns:
            for path in glob.glob(os.path.join(directory, pattern)):
                try:
                    if os.path.getmtime(path) < cutoff:
                        os.remove(path)
                        print(f"Deleted old snapshot: {path}")
                except OSError as exc:
                    print(f"Could not delete {path}: {exc}")


def load_sales_history_file():
    for path in [SALES_HISTORY_FILE, USB_SALES_HISTORY_FILE]:
        if os.path.exists(path):
            try:
                print(f"Reading sales history from {path}...")
                return pd.read_csv(path)
            except Exception as exc:
                print(f"Error reading {path}: {exc}")
    return pd.DataFrame()


def get_sellerboard_report(prefix, url):
    """Use local snapshot if < 12h old; otherwise download fresh."""
    os.makedirs(LOCAL_SELLERBOARD_DIR, exist_ok=True)
    local_file = get_latest_sellerboard_file(prefix)
    if local_file:
        age_hours = (time.time() - os.path.getmtime(local_file)) / 3600
        if age_hours < 12:
            content = read_local_sellerboard_file(local_file)
            if "ASIN" in content and "SKU" in content:
                print(
                    f"Using local SellerBoard snapshot ({age_hours:.1f}h old): {local_file}"
                )
                return content, local_file
            else:
                print(f"Local file {local_file} is invalid. Downloading fresh...")
        else:
            print(f"Local snapshot too old ({age_hours:.1f}h). Downloading fresh...")

    timestamp = pd.Timestamp.now().strftime("%Y-%m-%d")
    print(f"Downloading {prefix} from SellerBoard...")
    
    max_retries = 5
    for attempt in range(max_retries):
        try:
            response = requests.get(url, headers=HEADERS, timeout=90)
            response.raise_for_status()
            text = response.content.decode("utf-8-sig", errors="replace")
            
            if "Report not ready" in text or "try again" in text:
                print(f"  Attempt {attempt+1}/{max_retries}: SellerBoard report not ready yet. Waiting 60s...")
                time.sleep(60)
                continue
                
            if "ASIN" not in text or "SKU" not in text:
                raise ValueError("Report content does not contain expected CSV headers (ASIN, SKU)")
                
            output_file = save_sellerboard_snapshot(prefix, timestamp, text)
            print(f"Saved {prefix} to {output_file}")
            return text, output_file
        except Exception as e:
            if attempt == max_retries - 1:
                print(f"All download attempts failed for {prefix}: {e}")
                # Try to fall back to the most recent valid snapshot, regardless of age
                local_file = get_latest_sellerboard_file(prefix)
                if local_file:
                    content = read_local_sellerboard_file(local_file)
                    if "ASIN" in content and "SKU" in content:
                        print(f"Falling back to latest valid local snapshot: {local_file}")
                        return content, local_file
                raise RuntimeError(f"Could not download or find a valid local snapshot for {prefix}: {e}")
            print(f"  Attempt {attempt+1}/{max_retries} failed: {e}. Retrying in 30s...")
            time.sleep(30)



def download_and_save_sales():
    """Download today's sales from SellerBoard and save to a local snapshot."""
    try:
        _, sales_file = get_sellerboard_report("sellerboard_ventas", SALES_URL)
        return sales_file
    except Exception as e:
        print(f"Exception downloading sales: {e}")
    return None


def get_sales_data():
    """Build sales history from Excel plus the accumulated SellerBoard history/snapshot files."""
    from datetime import datetime, timedelta

    sales_365_sku = {}
    sales_60_sku = {}
    sales_30_sku = {}
    sales_7_sku = {}
    sales_365_asin = {}
    sales_7_asin = {}

    cutoff_7 = datetime.now() - timedelta(days=7)
    cutoff_30 = datetime.now() - timedelta(days=30)
    cutoff_60 = datetime.now() - timedelta(days=60)
    cutoff_365 = datetime.now() - timedelta(days=365)

    if os.path.exists(VENTAS_FILE):
        print(f"Reading {VENTAS_FILE}...")
        try:
            df = pd.read_excel(VENTAS_FILE)
            df["ASIN"] = df["ASIN"].astype(str).str.strip().str.upper()
            df["SKU"] = df["SKU"].astype(str).str.strip().str.upper()
            date_col = next((c for c in df.columns if "date" in c.lower()), None)
            if date_col:
                df["ParsedDate"] = pd.to_datetime(df[date_col], dayfirst=True, errors="coerce")
                df_filtered = df[df["ParsedDate"] >= cutoff_365]
            else:
                df_filtered = df
            sales_365_sku = df_filtered.groupby("SKU")["Units"].sum().to_dict()
            sales_365_asin = df_filtered.groupby("ASIN")["Units"].sum().to_dict()
            print(f"  365-day sales from Excel (filtered): {len(sales_365_sku)} SKUs")
        except Exception as e:
            print(f"Error reading {VENTAS_FILE}: {e}")

    if os.path.exists(VENTAS_60_FILE):
        print(f"Reading {VENTAS_60_FILE}...")
        try:
            df = pd.read_excel(VENTAS_60_FILE)
            df["ASIN"] = df["ASIN"].astype(str).str.strip().str.upper()
            df["SKU"] = df["SKU"].astype(str).str.strip().str.upper()
            date_col = next((c for c in df.columns if "date" in c.lower()), None)
            if date_col:
                df["ParsedDate"] = pd.to_datetime(df[date_col], dayfirst=True, errors="coerce")
                df_filtered = df[df["ParsedDate"] >= cutoff_60]
            else:
                df_filtered = df
            sales_60_sku = df_filtered.groupby("SKU")["Units"].sum().to_dict()
            print(f"  60-day sales from Excel (filtered): {len(sales_60_sku)} SKUs")
        except Exception as e:
            print(f"Error reading {VENTAS_60_FILE}: {e}")

    # Build unique daily sales dataframe from history file or snapshot daily files
    daily_sales_df = pd.DataFrame()

    history_df = load_sales_history_file()
    if not history_df.empty:
        try:
            history_df = history_df.copy()
            history_df["Date"] = pd.to_datetime(
                history_df["Date"], dayfirst=True, errors="coerce"
            )
            history_df["ASIN"] = history_df["ASIN"].astype(str).str.strip().str.upper()
            history_df["SKU"] = history_df["SKU"].astype(str).str.strip().str.upper()
            history_df["Marketplace"] = history_df["Marketplace"].astype(str).str.strip()
            
            # Correctly compute TotalUnits as UnitsOrganic + UnitsPPC to avoid ad category double-counting
            u_org = pd.to_numeric(history_df["UnitsOrganic"], errors="coerce").fillna(0) if "UnitsOrganic" in history_df.columns else 0
            u_ppc = pd.to_numeric(history_df["UnitsPPC"], errors="coerce").fillna(0) if "UnitsPPC" in history_df.columns else 0
            history_df["TotalUnits"] = u_org + u_ppc
            
            daily_sales_df = history_df[["Date", "Marketplace", "ASIN", "SKU", "TotalUnits"]]
            print(f"  Loaded {len(daily_sales_df)} daily sales rows from sales_history.csv")
        except Exception as e:
            print(f"Error processing sales history file: {e}")
    else:
        sales_files = get_sellerboard_snapshot_files("sellerboard_ventas")
        if sales_files:
            print(f"Processing {len(sales_files)} daily sales files from local snapshots...")
            all_dfs = []
            for f in sales_files:
                try:
                    df = pd.read_csv(f)
                    df["Date"] = pd.to_datetime(
                        df["Date"], format="%d/%m/%Y", errors="coerce"
                    )
                    df["ASIN"] = df["ASIN"].astype(str).str.strip().str.upper()
                    df["SKU"] = df["SKU"].astype(str).str.strip().str.upper()
                    df["Marketplace"] = df["Marketplace"].astype(str).str.strip()

                    u_org = pd.to_numeric(df["UnitsOrganic"], errors="coerce").fillna(0) if "UnitsOrganic" in df.columns else 0
                    u_ppc = pd.to_numeric(df["UnitsPPC"], errors="coerce").fillna(0) if "UnitsPPC" in df.columns else 0
                    df["TotalUnits"] = u_org + u_ppc

                    all_dfs.append(df[["Date", "Marketplace", "ASIN", "SKU", "TotalUnits"]])
                except Exception as e:
                    print(f"  Error reading {f}: {e}")
            if all_dfs:
                combined_df = pd.concat(all_dfs, ignore_index=True)
                # Deduplicate overlapping daily snapshot entries
                combined_df.drop_duplicates(subset=["Date", "Marketplace", "SKU"], keep="last", inplace=True)
                daily_sales_df = combined_df
                print(f"  Deduplicated snapshots: {len(daily_sales_df)} unique daily records.")

    # Integrate daily sales into the respective maps
    if not daily_sales_df.empty:
        df_7 = daily_sales_df[daily_sales_df["Date"] >= cutoff_7]
        df_30 = daily_sales_df[daily_sales_df["Date"] >= cutoff_30]
        df_60 = daily_sales_df[daily_sales_df["Date"] >= cutoff_60]
        df_365 = daily_sales_df[daily_sales_df["Date"] >= cutoff_365]

        # Integrate SKU level sales
        for sku, group in df_7.groupby("SKU"):
            units = group["TotalUnits"].sum()
            if units > 0:
                sales_7_sku[sku] = sales_7_sku.get(sku, 0) + units

        for sku, group in df_30.groupby("SKU"):
            units = group["TotalUnits"].sum()
            if units > 0:
                sales_30_sku[sku] = sales_30_sku.get(sku, 0) + units

        for sku, group in df_60.groupby("SKU"):
            units = group["TotalUnits"].sum()
            if units > 0:
                sales_60_sku[sku] = sales_60_sku.get(sku, 0) + units

        for sku, group in df_365.groupby("SKU"):
            units = group["TotalUnits"].sum()
            if units > 0:
                sales_365_sku[sku] = sales_365_sku.get(sku, 0) + units

        # Integrate ASIN level sales
        for asin, group in df_7.groupby("ASIN"):
            units = group["TotalUnits"].sum()
            if units > 0:
                sales_7_asin[asin] = sales_7_asin.get(asin, 0) + units

        for asin, group in df_365.groupby("ASIN"):
            units = group["TotalUnits"].sum()
            if units > 0:
                sales_365_asin[asin] = sales_365_asin.get(asin, 0) + units

        print(f"  Integrated daily sales: 7-day={len(sales_7_sku)} SKUs, 30-day={len(sales_30_sku)} SKUs, 60-day={len(sales_60_sku)} SKUs, 365-day={len(sales_365_sku)} SKUs")

    return sales_365_sku, sales_60_sku, sales_30_sku, sales_7_sku, sales_365_asin, sales_7_asin


def sync():
    supplier_stocks = download_supplier_data()

    # Download and save today's sales to build historical record
    today_file = download_and_save_sales()

    # Build sales history from all available CSV files
    sales_365_sku_map, sales_60_sku_map, sales_30_sku_map, sales_7_sku_map, sales_365_asin_map, sales_7_asin_map = get_sales_data()

    print(f"Fetching FBA report...")
    try:
        inventory_text, inventory_file = get_sellerboard_report(
            "sellerboard_inventory", CSV_URL
        )
        print(f"Reading inventory from {inventory_file}")

        reader = csv.DictReader(inventory_text.splitlines())
        rows = list(reader)
        existing_titles_by_sku = load_existing_titles_by_sku()

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
        sales_fba_skus = {s.upper() for s in sales_365_sku_map if "FBA" in s.upper()}
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
            sent_to_fba = safe_f(row.get("Sent  to FBA", "0"))
            reserved = safe_f(row.get("Reserved", "0"))
            transit = sent_to_fba + reserved
            effective_stock = stock_amz + transit
            
            asin = row.get("ASIN", "").strip().upper()
            if not asin or asin == "NAN":
                continue
            if asin not in asin_total_stock:
                asin_total_stock[asin] = 0
            asin_total_stock[asin] += effective_stock

        # Crear set de ASINs que ya tienen stock
        asins_with_stock = {
            asin for asin, total in asin_total_stock.items() if total > 0
        }

        # Segunda pasada: generar datos
        for idx, row in enumerate(rows):
            sku = row.get("SKU", "").upper()
            if not sku.endswith("FBA"):
                continue

            prov_key, clean_id = get_provider_and_id(sku)
            provider_name = {
                "SG": "Signes",
                "VC": "Minerales",
                "DC": "Dcasa",
                "MD": "Trediser",
            }.get(prov_key, "Unknown")
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
            velocity = safe_f(row.get("Estimated Sales Velocity", "0"))
            # Recalculate days_left using effective_stock when transit covers the gap
            if transit > 0 and days_left <= 7:
                if velocity > 0:
                    effective_days = round(effective_stock / velocity)
                    days_left = max(days_left, effective_days)
                else:
                    # No velocity data — assume stock is covered, use sentinel
                    days_left = 999

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

            sales_365 = int(sales_365_sku_map.get(sku, 0) or 0)
            sales_60 = int(sales_60_sku_map.get(sku, 0) or 0)
            sales_30 = int(sales_30_sku_map.get(sku, 0) or 0)
            sales_7 = int(sales_7_sku_map.get(sku, 0) or 0)

            is_back_in_stock = (
                (stock_amz == 0)
                and (transit == 0)
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
                    "title": existing_titles_by_sku.get(sku)
                    or translate_to_spanish(row.get("Title", "")),
                    "roi": roi,
                    "stock_amz": stock_amz,
                    "velocity": velocity,
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
                    "sales_30": sales_30,
                    "sales_7": sales_7,
                    "is_back_in_stock": is_back_in_stock,
                    "is_slow_moving": is_slow_moving,
                }
            )

        # FBM→FBA recommendations: ASINs con ventas pero sin listing FBA todavía
        fbm_recommendations = []
        for asin, units in sales_365_asin_map.items():
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
                    "title": existing_titles_by_sku.get(base_sku)
                    or translate_to_spanish(sku_to_title.get(base_sku, "")),
                    "sales_365": int(units or 0),
                    "sales_7": int(sales_7_asin_map.get(asin, 0) or 0),
                    "provider": {
                        "SG": "Signes",
                        "VC": "Minerales",
                        "DC": "Dcasa",
                        "MD": "Trediser",
                    }.get(prov_key, "Unknown"),
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
        json_serializer = lambda x: (
            bool(x) if isinstance(x, (bool, np.bool_))
            else str(x) if isinstance(x, np.integer)
            else float(x) if isinstance(x, np.floating)
            else x
        )
        output_paths = [OUTPUT_JSON]
        if os.path.exists(os.path.dirname(OUTPUT_JSON_DIST)):
            output_paths.append(OUTPUT_JSON_DIST)
        for out_path in output_paths:
            with open(out_path, "w") as f:
                json.dump(final_data, f, indent=2, default=json_serializer)
        print(
            f"Synced {len(data)} SKUs. FBM Recs: {len(fbm_recommendations)}. Slow Moving: {final_data['summary']['slow_moving_count']}"
        )
        cleanup_old_sellerboard_snapshots()

        # Auto-commit, push, and deploy so Vercel always serves fresh data
        import subprocess
        from datetime import date
        repo_root = os.path.dirname(os.path.abspath(__file__))
        fba_dir = os.path.dirname(OUTPUT_JSON)  # fba-replenishment/
        today = date.today().isoformat()
        try:
            subprocess.run(["git", "add", OUTPUT_JSON], cwd=repo_root, check=True)
            result = subprocess.run(
                ["git", "diff", "--cached", "--quiet"],
                cwd=repo_root,
            )
            if result.returncode != 0:  # there are staged changes
                subprocess.run(
                    ["git", "commit", "-m", f"data: auto-sync FBA data {today}"],
                    cwd=repo_root,
                    check=True,
                )
                subprocess.run(["git", "push", "origin", "main"], cwd=repo_root, check=True)
                print(f"Git: committed and pushed data.json ({today})")
            else:
                print("Git: data.json unchanged, no commit needed")
        except subprocess.CalledProcessError as git_err:
            print(f"Git push failed (data still saved locally): {git_err}")

        # Deploy to Vercel directly (bypasses GitHub webhook)
        # Retry up to 3 times — network may not be ready at 7AM Mac wakeup
        try:
            vercel_bin = subprocess.run(["which", "vercel"], capture_output=True, text=True).stdout.strip()
            if not vercel_bin:
                print("Vercel: CLI not found, skipping deploy")
            else:
                deployed = False
                for attempt in range(1, 4):
                    try:
                        subprocess.run(
                            [vercel_bin, "deploy", "--prod", "--yes",
                             "--scope", "christians-projects-dd62b5fc"],
                            cwd=repo_root,
                            check=True,
                            timeout=120,
                        )
                        print(f"Vercel: deployed to production (attempt {attempt})")
                        deployed = True
                        break
                    except Exception as e:
                        print(f"Vercel deploy attempt {attempt}/3 failed: {e}")
                        if attempt < 3:
                            time.sleep(30)  # wait 30s for network to stabilize
                if not deployed:
                    print("Vercel deploy failed after 3 attempts (data still in git)")
        except Exception as vercel_err:
            print(f"Vercel deploy error: {vercel_err}")

    except Exception as e:
        print(f"Error: {e}")


if __name__ == "__main__":
    sync()
