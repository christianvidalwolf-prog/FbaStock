import pandas as pd
import os
import csv
import json

WORK_DIR = '/Users/christianvidalwolf/Stock'
NUEVOS_SG1_FILE = f"{WORK_DIR}/Nuevos SG1.xlsx"
SIGNES_FEED = f"{WORK_DIR}/signes_stock.csv"

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

def dry_run():
    print("Loading Signes feed...")
    df_sg = pd.read_csv(SIGNES_FEED, sep=";", encoding="latin-1", header=None, on_bad_lines="skip")
    df_sg[0] = df_sg[0].astype(str).str.replace("SG-", "").str.strip().str.replace(".0", "", regex=False)
    sg_data = df_sg.set_index(0)

    print(f"Loading {NUEVOS_SG1_FILE}...")
    df_n = pd.read_excel(NUEVOS_SG1_FILE)
    if 'item_sku' not in df_n.columns:
        df_n.columns = df_n.iloc[0]
        df_n = df_n[1:]
    
    sg_new = df_n[df_n['item_sku'].astype(str).str.endswith('SGRI')]
    
    results = []
    for _, row in sg_new.head(10).iterrows():
        sku = str(row['item_sku']).strip()
        price = to_num(row['standard_price'])
        prod_num = sku[:-4]
        
        stock = 0
        if prod_num in sg_data.index:
            s_row = sg_data.loc[prod_num]
            if isinstance(s_row, pd.DataFrame):
                s_row = s_row.iloc[0]
            s_raw = to_num(s_row[2])
            p_cost = to_num(s_row[3])
            
            if s_raw > 3:
                if s_raw >= 20 or p_cost >= 20:
                    stock = s_raw
        
        min_p = round(price / 2, 2)
        max_p = round(price * 2, 2)
        
        results.append([sku, price, min_p, max_p, int(stock)])
        
    print("\nDry run results (First 10):")
    print("SKU | Price | Min | Max | Stock")
    for r in results:
        print(f"{r[0]} | {r[1]} | {r[2]} | {r[3]} | {r[4]}")

dry_run()
