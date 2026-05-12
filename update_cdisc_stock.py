import pandas as pd
import os
import re
import glob
from datetime import datetime

# --- CONFIGURATION ---
WORK_DIR = "/Users/christianvidalwolf/Stock"
CDISC_FILE = os.path.join(WORK_DIR, "cdisc precio.xlsx")
CDISC_HISTORY_FILE = os.path.join(WORK_DIR, f"cdisc precio {datetime.now().strftime('%Y%m%d')}.xlsx")
MINERALES_FEED = os.path.join(WORK_DIR, "minerales_feed.xml")
SIGNES_STOCK = os.path.join(WORK_DIR, "signes_stock.csv")
CATALOG_FILE = os.path.join(WORK_DIR, "catalog_robust.json")
BASE_STOCKS_FILE = os.path.join(WORK_DIR, "base_stocks.json")
STOCK_TREDISER_FILE = os.path.join(WORK_DIR, "STOCK TREDISER.xls")

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

    # 4. trEDISER (MD) from Excel — only source for MD stock
    if os.path.exists(STOCK_TREDISER_FILE):
        try:
            df_t = pd.read_excel(STOCK_TREDISER_FILE)
            df_t = df_t.dropna(subset=['Código'])
            for _, row in df_t.iterrows():
                code = str(row['Código']).strip().replace(".0", "")
                if not code: continue
                # Store by numeric code
                stock_val = to_num(row['Unnamed: 3'])
                stocks["MD"][code] = stock_val
        except Exception as e:
            print(f"Error loading trEDISER Excel: {e}")

    return stocks

def get_provider_and_id(sku):
    sku_orig = str(sku).upper().strip()
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
    
    if sku_orig.endswith("VC"): return "VC", sku_orig[:-2]
    if sku_orig.endswith("SG"): return "SG", sku_orig[:-2]
    if sku_orig.endswith("DC"): return "DC", sku_orig[:-2]
    if sku_orig.endswith("CLM"): return "DC", sku_orig[:-3]
    if sku_orig.endswith("MD"): return "MD", sku_orig[:-2]
    
    if "VC" in sku_orig: 
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
    print(f"--- Starting Cdiscount Stock Update @ {datetime.now()} ---")
    if not os.path.exists(CDISC_FILE):
        print(f"Error: {CDISC_FILE} not found.")
        return

    supplier_stocks = load_suppliers_data()
    print(f"Loaded stocks: VC:{len(supplier_stocks['VC'])}, SG:{len(supplier_stocks['SG'])}, DC:{len(supplier_stocks['DC'])}, MD:{len(supplier_stocks['MD'])}")
    
    try:
        # Load the whole file without header to preserve structure (4 rows of headers)
        df = pd.read_excel(CDISC_FILE, header=None)
    except Exception as e:
        print(f"Error reading Excel: {e}")
        return

    updated_count = 0
    not_found_count = 0
    
    # Data starts at row 4 (index 4)
    # Column 0 is SKU, Column 2 is Stock
    for idx in range(4, len(df)):
        sku = str(df.at[idx, 0]).strip()
        if sku == "nan" or not sku:
            continue
            
        prov, clean_id = get_provider_and_id(sku)
        
        if prov and clean_id:
            raw_stock = 0
            if prov == "MD":
                # Only use numeric ID from TREDISER Excel
                raw_stock = supplier_stocks.get("MD", {}).get(clean_id, 0)
            else:
                raw_stock = supplier_stocks.get(prov, {}).get(clean_id, 0)
            
            final_stock = 0
            if prov == "VC":
                if raw_stock >= 5: final_stock = raw_stock
            elif prov == "SG":
                if raw_stock > 3: final_stock = raw_stock
            elif prov == "DC":
                if raw_stock > 3: final_stock = raw_stock
            elif prov == "MD":
                final_stock = raw_stock
            
            df.at[idx, 2] = int(final_stock)
            updated_count += 1
        else:
            not_found_count += 1

    print(f"Processed {len(df)-4} data rows. Updated: {updated_count}, Provider not matched: {not_found_count}")
    
    try:
        df.to_excel(CDISC_FILE, index=False, header=False)
        df.to_excel(CDISC_HISTORY_FILE, index=False, header=False)
        print(f"Successfully saved updated stock to {CDISC_FILE} and {CDISC_HISTORY_FILE}")
    except Exception as e:
        print(f"Error saving Excel: {e}")

if __name__ == "__main__":
    main()
