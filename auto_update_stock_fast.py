
import os
import json
import csv
import time
import pandas as pd
from ftplib import FTP
import datetime
import subprocess

# --- CONFIGURATION ---
WORK_DIR = '/Users/christianvidalwolf/Stock'
CATALOG_FILE = f'{WORK_DIR}/catalog_robust.json'
EXCEL_PRICES_FILE = f'{WORK_DIR}/excel_prices_final.json'
OUTPUT_FILE = f'{WORK_DIR}/STOCK AMZ.txt'
PRECIOS_FILE = f'{WORK_DIR}/precios ES.xlsx'
SKUS_FORZAR_CERO = {
    '2450VC', '2450VCI',  # ASIN B010TN6SXU
    # VC SKUs (User request)
    '010VC', '3019VC', '3020VC', '3027VC', '3030VC', '3035VC',
    '3298VC0', '4462VC', '4695VC', '4697VC', '2088VC0',
}

def to_num(val):
    if val is None or val == "": return 0
    try:
        if isinstance(val, (int, float)): return float(val)
        val_str = str(val).replace(',', '.')
        return float(val_str)
    except: return 0

def download_files():
    providers = {
        'dcasa': {'type': 'ftp', 'host': 'data.dcasacollection.com', 'user': 'sek4283', 'pass': 'Rx34z5m6_pER'},
        'mina': {'type': 'url', 'url': 'https://vivescortadaimport.com/modules/doofinder/feed2.php?language=ES&currency=EUR'},
        'signes': {'type': 'url', 'url': 'https://signesconexion.com/stock/STOCK-44880.CSV'}
    }
    paths = {}
    
    # DCASA
    MIN_DCASA_ROWS = 1000
    
    # Check if we already have today's file (downloaded by curl)
    today_str = datetime.datetime.now().strftime('%Y%m%d')
    local = sorted([f for f in os.listdir(WORK_DIR) if f.startswith('DataWeb') and f.endswith('.csv')])
    
    if local and today_str in local[-1]:
        print(f"DCASA file for today ({local[-1]}) already exists. Skipping FTP download.")
        paths['dcasa'] = f"{WORK_DIR}/{local[-1]}"
    else:
        for attempt in range(1, 7):
            try:
                print(f"Downloading DCASA from FTP (attempt {attempt}/6)...")
                ftp = FTP(providers['dcasa']['host'], timeout=180)
                ftp.set_pasv(True)
                ftp.login(user=providers['dcasa']['user'], passwd=providers['dcasa']['pass'])
                files = ftp.nlst()
                dfiles = sorted([f for f in files if f.startswith('DataWeb') and f.endswith('.csv')])
                if dfiles:
                    path = f"{WORK_DIR}/{dfiles[-1]}"
                    # If local file already exists and is the same name, skip if you want, 
                    # but let's re-download to be sure it's complete if it wasn't today's.
                    with open(path, 'wb') as fp: ftp.retrbinary(f'RETR {dfiles[-1]}', fp.write)
                    with open(path, encoding='latin-1') as chk:
                        row_count = sum(1 for _ in chk) - 1
                    if row_count < MIN_DCASA_ROWS:
                        print(f"WARNING: DCASA file has only {row_count} rows (expected >{MIN_DCASA_ROWS}).")
                    else:
                        paths['dcasa'] = path
                ftp.quit()
                break  # success
            except Exception as e:
                print(f"WARNING: FTP attempt {attempt} failed ({e}).")
                if attempt < 6:
                    wait_time = 15 * attempt
                    print(f"Waiting {wait_time}s before next attempt...")
                    time.sleep(wait_time)
        local = sorted([f for f in os.listdir(WORK_DIR) if f.startswith('DataWeb') and f.endswith('.csv')])
        # Skip files with < MIN_DCASA_ROWS rows
        for fname in reversed(local):
            fpath = f"{WORK_DIR}/{fname}"
            with open(fpath, encoding='latin-1') as chk:
                rc = sum(1 for _ in chk) - 1
            if rc >= MIN_DCASA_ROWS:
                paths['dcasa'] = fpath
                break

    if 'dcasa' in paths:
        print(f"Using DCASA file: {os.path.basename(paths['dcasa'])}")

    # MINA & SIGNES
    subprocess.run(['curl', '-L', providers['mina']['url'], '-o', f"{WORK_DIR}/minerales_feed.xml"], capture_output=True)
    paths['mina'] = f"{WORK_DIR}/minerales_feed.xml"
    subprocess.run(['curl', '-L', providers['signes']['url'], '-o', f"{WORK_DIR}/signes_stock.csv"], capture_output=True)
    paths['signes'] = f"{WORK_DIR}/signes_stock.csv"
    
    return paths

def run_fast_update():
    print(f"--- Fast Stock Update @ {datetime.datetime.now()} ---")
    
    if not os.path.exists(CATALOG_FILE) or not os.path.exists(EXCEL_PRICES_FILE):
        print("Error: JSON files missing. Run extract_full.py first.")
        return
        
    with open(CATALOG_FILE, 'r') as f: catalog = json.load(f)
    with open(EXCEL_PRICES_FILE, 'r') as f: excel_prices = json.load(f)
    print(f"Loaded {len(catalog)} SKUs mapping.")
    
    # Load Base Stocks for MD (from STOCK ES)
    BASE_STOCKS_FILE = f'{WORK_DIR}/base_stocks.json'
    base_stocks = {}
    if os.path.exists(BASE_STOCKS_FILE):
        with open(BASE_STOCKS_FILE, 'r') as f: base_stocks = json.load(f)
        print(f"Loaded {len(base_stocks)} base stocks.")
    
    # Load Prices from precios ES.xlsx
    print(f"Loading prices from {PRECIOS_FILE}...")
    try:
        df_p = pd.read_excel(PRECIOS_FILE)
        # Ensure SKU is string and strip whitespace
        df_p['sku'] = df_p['sku'].astype(str).str.strip()
        # Clean price: convert to string, replace comma, then to numeric
        df_p['price_clean'] = df_p['price'].astype(str).str.replace(',', '.').str.strip()
        df_p['price_num'] = pd.to_numeric(df_p['price_clean'], errors='coerce').fillna(0)
        prices_map = df_p.set_index('sku')['price_num'].to_dict()
        print(f"Loaded {len(prices_map)} prices.")
    except Exception as e:
        print(f"Error loading prices from Excel: {e}")
        prices_map = {}

    paths = download_files()
    data = {}
    
    # Load Feeds
    if 'dcasa' in paths:
        try:
            df = pd.read_csv(paths['dcasa'], sep=';', encoding='latin-1', on_bad_lines='skip')
            df.columns = [c.strip() for c in df.columns]
            df['id_str'] = df['CODIGO'].astype(str).str.replace('.0', '', regex=False)
            data['DC'] = df.set_index('id_str')
        except: pass
    if 'mina' in paths:
        try:
            df = pd.read_csv(paths['mina'], sep='|', encoding='utf-8', header=None, on_bad_lines='skip')
            df[0] = df[0].astype(str).str.replace('.0', '', regex=False)
            data['VC'] = df.set_index(0)
        except: pass
    if 'signes' in paths:
        try:
            df = pd.read_csv(paths['signes'], sep=';', encoding='latin-1', header=None, on_bad_lines='skip')
            df[0] = df[0].astype(str).str.replace('SG-', '').str.strip().str.replace('.0', '', regex=False)
            data['SG'] = df.set_index(0)
        except: pass

    output_rows = [ ['sku', 'price', 'minimum-seller-allowed-price', 'maximum-seller-allowed-price', 'quantity', 'fulfillment-channel', 'handling-time'] ]
    
    print("Calculating updates...")
    for entry in catalog:
        sku = entry['sku']
        provider = entry['provider']
        l_id = entry['id']
        
        sheet_map = {'VC':'INICIOVC', 'DC':'DcasaWeb', 'SG':'Signes', 'MD':'Madelcar'}
        sheet_name = sheet_map.get(provider)
        
        # Prepare IDs and secondary info
        lookup_id = l_id.replace('.0', '')
        provider_info = excel_prices.get(sheet_name, {}).get(l_id, {})
        divisor = provider_info.get('divisor', 1.0) if isinstance(provider_info, dict) else 1.0
        
        # FINAL PRICE: Always from precios ES.xlsx, ensuring it's a number
        final_price = to_num(prices_map.get(sku, 0))
        
        # If not found in precios ES, use a default or 0 (User said prices ARE in that file)
        final_stock = 0 
        
        try:
            if provider == 'VC' and 'VC' in data and lookup_id in data['VC'].index:
                row = data['VC'].loc[lookup_id]
                if isinstance(row, pd.DataFrame): row = row.iloc[0]
                raw_stock = to_num(row[7])
                if raw_stock >= 5:
                    final_stock = raw_stock
                else: final_stock = 0
                    
            elif provider == 'DC' and 'DC' in data and lookup_id in data['DC'].index:
                row = data['DC'].loc[lookup_id]
                if isinstance(row, pd.DataFrame): row = row.iloc[0]
                s_raw, p_cost = to_num(row['STOCK_DISPONIBLE']), to_num(row['Tarifa A'])
                # Stock filter for DC (using divisor from excel_prices_final if needed for the stock threshold)
                if s_raw > 3 and (s_raw >= 20 or (p_cost * divisor) >= 20):
                    final_stock = 99 if s_raw > 1 else (1 if s_raw == 1 else 0)
                else: final_stock = 0

            elif provider == 'SG' and 'SG' in data and lookup_id in data['SG'].index:
                row = data['SG'].loc[lookup_id]
                if isinstance(row, pd.DataFrame): row = row.iloc[0]
                s_raw = to_num(row[2])
                # Stock filter for SG
                if s_raw > 3:
                    # Filter '83627' (User request)
                    if '83627' in sku:
                        final_stock = 0
                    else:
                        p_cost = to_num(row[3])
                        if s_raw >= 20 or p_cost >= 20:
                            final_stock = s_raw
                        else: final_stock = 0
                else: final_stock = 0
                
            elif provider == 'MD':
                # Take stock from STOCK ES extract (inicio plus 2023)
                final_stock = base_stocks.get(sku, 0)
                
        except: pass

        # Force stock=0 for SKUs in SKUS_FORZAR_CERO (any provider)
        if sku in SKUS_FORZAR_CERO:
            final_stock = 0

        # Final Safety Check
        if final_price <= 0: final_price = 0

        # Min/Max (Col C, D)
        min_p_val = round(final_price / 2, 2) if final_price > 0 else 0
        max_p_val = round(final_price * 2, 2) if final_price > 0 else 0
        
        # Format prices with COMMA as decimal separator for Amazon ES
        final_price_str = f"{final_price:g}".replace('.', ',') if final_price > 0 else "0"
        min_p_str = f"{min_p_val:g}".replace('.', ',') if min_p_val > 0 else ""
        max_p_str = f"{max_p_val:g}".replace('.', ',') if max_p_val > 0 else ""

        output_rows.append([
            sku,
            final_price_str,
            min_p_str,
            max_p_str,
            str(int(final_stock)),
            "", ""
        ])

    print(f"Writing {len(output_rows)} rows to {OUTPUT_FILE}...")
    with open(OUTPUT_FILE, 'w', encoding='utf-8', newline='') as f:
        writer = csv.writer(f, delimiter='\t')
        writer.writerows(output_rows)
    print("Export finished.")

if __name__ == "__main__":
    run_fast_update()
