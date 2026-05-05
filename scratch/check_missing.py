import pandas as pd
import json
import os

WORK_DIR = '/Users/christianvidalwolf/Stock'

def check():
    df = pd.read_excel(os.path.join(WORK_DIR, 'Nuevos SG1.xlsx'))
    # The columns shown earlier were 'Unnamed: 0' (item_sku) and 'Unnamed: 1' (standard_price)
    # But wait, head() showed row 0 has 'item_sku' and 'standard_price'
    # pandas might have read them as data if I didn't specify header.
    # Let's try to find headers correctly.
    if 'item_sku' not in df.columns:
        df.columns = df.iloc[0]
        df = df[1:]
    
    sgri_skus = df[df['item_sku'].astype(str).str.endswith('SGRI')]
    print(f"Total SGRI items in Excel: {len(sgri_skus)}")
    
    with open(os.path.join(WORK_DIR, 'catalog_robust.json'), 'r') as f:
        catalog = json.load(f)
    
    catalog_skus = {e['sku'] for e in catalog}
    
    missing = []
    for _, row in sgri_skus.iterrows():
        sku = str(row['item_sku']).strip()
        if sku not in catalog_skus:
            missing.append(sku)
            
    print(f"Items NOT in catalog_robust.json: {len(missing)}")
    if missing:
        print("First 10 missing:", missing[:10])

check()
