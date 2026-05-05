import pandas as pd
import requests
import io
import os
from datetime import datetime, timedelta

# Configuration
VENTAS_FILE = 'Ventas 365.xlsx'
SELLERBOARD_URL = "https://app.sellerboard.com/es/automation/reports?id=a1a2f4284b8043c39964edfe3cef86ca&format=csv&t=bbc9d347dff7407dbd01c90884f31121"
OUTPUT_FILE = 'Ventas 365 Analizadas.xlsx'

def to_num(val):
    if val is None or val == "": return 0.0
    try:
        if isinstance(val, (int, float)): return float(val)
        val_str = str(val).replace(".", "").replace(",", ".")
        return float(val_str)
    except:
        return 0.0

def analyze():
    print(f"Reading {VENTAS_FILE}...")
    df_orders = pd.read_excel(VENTAS_FILE)
    
    # Clean SKU and Date
    df_orders['SKU'] = df_orders['SKU'].astype(str).str.strip()
    df_orders['Order date'] = pd.to_datetime(df_orders['Order date'])
    
    print("Grouping sales by SKU...")
    # Aggregate sales
    sku_summary = df_orders.groupby('SKU').agg({
        'Units': 'sum',
        'Order date': ['max', 'min', 'count']
    }).reset_index()
    
    sku_summary.columns = ['SKU', 'Ventas 365 días', 'Last Sale', 'First Sale', 'Order Count']
    
    print(f"Fetching current stock from Sellerboard...")
    try:
        r = requests.get(SELLERBOARD_URL)
        r.raise_for_status()
        content = r.content.decode('utf-8-sig')
        df_stock = pd.read_csv(io.StringIO(content))
        df_stock['SKU'] = df_stock['SKU'].astype(str).str.strip()
        
        # We need 'FBA/FBM Stock'
        # Let's map stock to sku_summary
        stock_map = {}
        for _, row in df_stock.iterrows():
            s = str(row.get('SKU', '')).strip()
            qty = to_num(row.get('FBA/FBM Stock', '0'))
            stock_map[s] = qty
            
        sku_summary['Current Stock'] = sku_summary['SKU'].map(stock_map).fillna(0)
    except Exception as e:
        print(f"Error fetching stock: {e}")
        sku_summary['Current Stock'] = 0

    # Logic 1: FBA products back in stock
    # "productos fba estan agotados hace tiempo y vuelven a estar en stock, y que hayan vendido mas de 8 unidades"
    # Criteria: SKU ends in FBA, Ventas > 8, Current Stock > 0.
    # To detect "out of stock for a time", we look for a gap in the order history.
    
    def detect_gap(sku):
        sku_orders = df_orders[df_orders['SKU'] == sku].sort_values('Order date')
        if len(sku_orders) < 2: return False
        # Calculate max gap between consecutive orders
        gaps = sku_orders['Order date'].diff().dt.days
        max_gap = gaps.max()
        return max_gap > 30 # More than 30 days gap

    sku_summary['Is FBA'] = sku_summary['SKU'].str.upper().str.endswith('FBA')
    
    print("Detecting gaps for FBA products...")
    # Only check gaps for FBA with > 8 sales
    fba_candidates = sku_summary[(sku_summary['Is FBA']) & (sku_summary['Ventas 365 días'] > 8) & (sku_summary['Current Stock'] > 0)]
    
    back_in_stock_skus = []
    for sku in fba_candidates['SKU']:
        if detect_gap(sku):
            back_in_stock_skus.append(sku)
            
    sku_summary['Back in Stock Alert'] = sku_summary['SKU'].isin(back_in_stock_skus)

    # Logic 2: FBM recommendation for FBA
    # "detectar prodcutos fbm (no terminacion FBA) que no existen en FBA y hayan vendido mas de 8"
    fba_skus = set(sku_summary[sku_summary['Is FBA']]['SKU'].str.upper())
    
    def needs_fba_recommendation(row):
        sku = str(row['SKU']).upper()
        if row['Is FBA']: return False
        if row['Ventas 365 días'] <= 8: return False
        
        # Check if SKU + FBA exists
        fba_version = sku + "FBA"
        if fba_version in fba_skus: return False
        return True

    sku_summary['Recommend FBA'] = sku_summary.apply(needs_fba_recommendation, axis=1)

    # Now, the user wants to add "Ventas 365 días" to the original file context or see it clearly.
    # We will create a summary sheet and also a detailed sheet if needed.
    
    # Save results
    with pd.ExcelWriter(OUTPUT_FILE) as writer:
        sku_summary.to_excel(writer, sheet_name='Summary', index=False)
        
        # Highlights
        back_in_stock_df = sku_summary[sku_summary['Back in Stock Alert']]
        back_in_stock_df.to_excel(writer, sheet_name='Back in Stock FBA', index=False)
        
        recommend_fba_df = sku_summary[sku_summary['Recommend FBA']]
        recommend_fba_df.to_excel(writer, sheet_name='FBM to FBA Recommendations', index=False)

    print(f"Analysis complete. Results saved to {OUTPUT_FILE}")

if __name__ == "__main__":
    analyze()
