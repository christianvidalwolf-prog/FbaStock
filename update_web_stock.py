import pandas as pd
import os
import re
import glob
from datetime import datetime
import json

# --- CONFIGURATION ---
WORK_DIR = "/Users/christianvidalwolf/Stock"
STOCK_WEB_XLSX = os.path.join(WORK_DIR, "Stock Web.xlsx")
STOCK_WEB_CSV = os.path.join(WORK_DIR, "Stock Web.csv")
STOCK_WEB_HISTORY_CSV = os.path.join(WORK_DIR, f"Stock Web {datetime.now().strftime('%Y%m%d')}.csv")
MINERALES_FEED = os.path.join(WORK_DIR, "minerales_feed.xml")
SIGNES_STOCK = os.path.join(WORK_DIR, "signes_stock.csv")
CATALOG_FILE = os.path.join(WORK_DIR, "catalog_robust.json")
BASE_STOCKS_FILE = os.path.join(WORK_DIR, "base_stocks.json")

def to_num(val):
    if val is None or val == "":
        return 0
    try:
        if isinstance(val, (int, float)):
            return float(val)
        val_str = str(val).replace(",", ".")
        return float(val_str)
    except:
        return 0

def get_latest_dcasa_file():
    local_files = sorted([f for f in os.listdir(WORK_DIR) if f.startswith("DataWeb") and f.endswith(".csv")])
    if not local_files:
        return None
    return os.path.join(WORK_DIR, local_files[-1])

def load_suppliers_data():
    stocks = {"VC": {}, "SG": {}, "DC": {}, "MD": {}}
    
    # 1. Minerales (VC)
    if os.path.exists(MINERALES_FEED):
        try:
            # Minerales usually uses pipe separator |
            with open(MINERALES_FEED, "r", encoding="utf-8", errors="ignore") as f:
                for line in f:
                    parts = line.split("|")
                    if len(parts) > 7:
                        p_id = parts[0].strip().upper()
                        try:
                            p_stock = int(float(parts[7].strip()))
                            stocks["VC"][p_id] = p_stock
                        except:
                            pass
        except Exception as e:
            print(f"Error loading Minerales: {e}")

    # 2. Signes (SG)
    if os.path.exists(SIGNES_STOCK):
        try:
            df = pd.read_csv(SIGNES_STOCK, sep=";", encoding="latin-1", header=None, on_bad_lines="skip")
            for _, row in df.iterrows():
                # Signes refs can be 'SG-12345' or '12345'
                raw_id = str(row[0]).replace("SG-", "").strip().upper()
                if raw_id.endswith(".0"): raw_id = raw_id[:-2]
                try:
                    p_stock = int(float(str(row[2]).replace(",", ".")))
                    stocks["SG"][raw_id] = p_stock
                except:
                    pass
        except Exception as e:
            print(f"Error loading Signes: {e}")

    # 3. Dcasa (DC/CLM)
    dcasa_file = get_latest_dcasa_file()
    if dcasa_file:
        try:
            df = pd.read_csv(dcasa_file, sep=";", encoding="latin-1", on_bad_lines="skip")
            df.columns = [c.strip() for c in df.columns]
            for _, row in df.iterrows():
                raw_id = str(row["CODIGO"]).strip().upper()
                if raw_id.endswith(".0"): raw_id = raw_id[:-2]
                try:
                    p_stock = int(float(str(row["STOCK_DISPONIBLE"]).replace(",", ".")))
                    stocks["DC"][raw_id] = p_stock
                except:
                    pass
        except Exception as e:
            print(f"Error loading Dcasa: {e}")

    # 4. Madelcar (MD) - from base_stocks.json
    if os.path.exists(BASE_STOCKS_FILE):
        try:
            with open(BASE_STOCKS_FILE, "r") as f:
                md_stocks = json.load(f)
                stocks["MD"] = md_stocks
        except Exception as e:
            print(f"Error loading Madelcar base stocks: {e}")

    return stocks

def get_provider_and_id(sku):
    """
    Detects provider and extracts ID.
    User rules: VC=Minerales, SG=Signes, DC/CLM=Dcasa.
    MD is Madelcar (from other scripts).
    """
    sku_orig = str(sku).upper().strip()
    
    # 1. Regex based on sync_fba_report.py patterns
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
        (r"^(\d+)MD", "MD"),
    ]
    for pattern, prov_key in patterns:
        m = re.match(pattern, sku_orig)
        if m:
            return prov_key, m.group(1)
    
    # 2. Simple suffix check for variety
    if sku_orig.endswith("VC"): return "VC", sku_orig[:-2]
    if sku_orig.endswith("SG"): return "SG", sku_orig[:-2]
    if sku_orig.endswith("DC"): return "DC", sku_orig[:-2]
    if sku_orig.endswith("CLM"): return "DC", sku_orig[:-3]
    if sku_orig.endswith("MD"): return "MD", sku_orig[:-2]
    
    # 3. Contains check as requested by user
    if "VC" in sku_orig: 
        # Extract digits
        digits = "".join(filter(str.isdigit, sku_orig))
        return "VC", digits
    if "SG" in sku_orig:
        digits = "".join(filter(str.isdigit, sku_orig))
        return "SG", digits
    if "DC" in sku_orig or "CLM" in sku_orig:
        digits = "".join(filter(str.isdigit, sku_orig))
        return "DC", digits
        
    return None, None

def main():
    print(f"--- Starting Web Stock Update @ {datetime.now()} ---")
    
    # Priority: CSV file if exists, otherwise XLSX for first-time migration
    source_file = None
    is_csv = False
    if os.path.exists(STOCK_WEB_CSV):
        source_file = STOCK_WEB_CSV
        is_csv = True
    elif os.path.exists(STOCK_WEB_XLSX):
        source_file = STOCK_WEB_XLSX
        is_csv = False
    
    if not source_file:
        print(f"Error: Neither {STOCK_WEB_CSV} nor {STOCK_WEB_XLSX} found.")
        return

    print(f"Reading from {os.path.basename(source_file)}...")
    supplier_stocks = load_suppliers_data()
    print(f"Loaded stocks: VC:{len(supplier_stocks['VC'])}, SG:{len(supplier_stocks['SG'])}, DC:{len(supplier_stocks['DC'])}, MD:{len(supplier_stocks['MD'])}")
    
    try:
        if is_csv:
            # Use semicolon as it is common in the user's files
            df_web = pd.read_csv(source_file, header=None, sep=";")
        else:
            df_web = pd.read_excel(source_file, header=None)
    except Exception as e:
        print(f"Error reading source file: {e}")
        return

    updated_count = 0
    not_found_count = 0
    
    for idx in df_web.index:
        sku = str(df_web.at[idx, 0]).strip()
        if sku == "nan" or not sku:
            continue
            
        prov, clean_id = get_provider_and_id(sku)
        
        if prov and clean_id:
            raw_stock = 0
            if prov == "MD":
                raw_stock = supplier_stocks.get("MD", {}).get(sku, 0)
            else:
                raw_stock = supplier_stocks.get(prov, {}).get(clean_id, 0)
            
            final_stock = 0
            if prov == "VC":
                if raw_stock >= 5:
                    final_stock = raw_stock
            elif prov == "SG":
                if raw_stock > 3:
                    final_stock = raw_stock
            elif prov == "DC":
                if raw_stock > 3:
                    final_stock = raw_stock
            elif prov == "MD":
                final_stock = raw_stock
            
            df_web.at[idx, 1] = float(final_stock)
            updated_count += 1
        else:
            not_found_count += 1

    print(f"Processed {len(df_web)} rows. Updated: {updated_count}, Provider not matched: {not_found_count}")
    
    try:
        # Always save to CSV from now on
        df_web.to_csv(STOCK_WEB_CSV, index=False, header=False, sep=";")
        df_web.to_csv(STOCK_WEB_HISTORY_CSV, index=False, header=False, sep=";")
        print(f"Successfully saved updated stock to {STOCK_WEB_CSV} and {STOCK_WEB_HISTORY_CSV}")
    except Exception as e:
        print(f"Error saving CSV: {e}")

if __name__ == "__main__":
    main()
