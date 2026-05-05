import pandas as pd
import requests
import io
import os
import re

CSV_URL = "https://app.sellerboard.com/es/automation/reports?id=a1a2f4284b8043c39964edfe3cef86ca&format=csv&t=bbc9d347dff7407dbd01c90884f31121"
VENTAS_FILE = "Ventas 365.xlsx"

def normalize_sku(sku):
    sku = str(sku).upper().strip()
    # Remove FBA suffixes
    sku = re.sub(r'(0|1|A|RG|I)?FBA$', '', sku)
    return sku

def audit():
    print("Fetching FBA Inventory...")
    r = requests.get(CSV_URL)
    df_inv = pd.read_csv(io.StringIO(r.content.decode('utf-8-sig')))
    fba_skus = set(df_inv['SKU'].astype(str).str.upper().str.strip())
    
    print(f"Reading {VENTAS_FILE}...")
    df_orders = pd.read_excel(VENTAS_FILE)
    df_orders['SKU'] = df_orders['SKU'].astype(str).str.strip().str.upper()
    
    # Get total sales per SKU
    summary = df_orders.groupby('SKU')['Units'].sum().reset_index()
    summary.columns = ['SKU', 'Ventas']
    
    # Identify FBM candidates (> 8 sales, not FBA)
    fbm_candidates = summary[(summary['Ventas'] > 8) & (~summary['SKU'].str.endswith('FBA'))].copy()
    
    print(f"Auditing {len(fbm_candidates)} candidates...")
    
    false_positives = []
    
    for _, row in fbm_candidates.iterrows():
        fbm_sku = row['SKU']
        
        # 1. Exact match with startswith logic (my current logic)
        exists_current = any(s.startswith(fbm_sku) and s.endswith('FBA') for s in fba_skus)
        
        # 2. Fuzzy match: Check if FBM SKU is a substring of any FBA SKU (excluding FBA suffix)
        # Or if they share the same normalized root
        norm_fbm = normalize_sku(fbm_sku)
        
        exists_fuzzy = False
        potential_matches = []
        for s in fba_skus:
            if not s.endswith('FBA'): continue
            norm_fba = normalize_sku(s)
            
            # If norm roots match exactly
            if norm_fbm == norm_fba:
                exists_fuzzy = True
                potential_matches.append(s)
            # Or if FBM SKU is inside the FBA SKU (e.g. 245910DC in 245910DCAFBA)
            elif fbm_sku in s:
                exists_fuzzy = True
                potential_matches.append(s)
                
        if exists_fuzzy and not exists_current:
            # This is a case where startswith failed but fuzzy found it
            # Actually, if exists_fuzzy is true, we should NOT recommend it.
            pass
            
        if exists_fuzzy:
            # Already has FBA, should NOT be recommended
            continue
        else:
            # Truly needs FBA recommendation
            pass

    # Let's find cases where the current logic might be FAILING (i.e. recommending when it shouldn't)
    # Re-evaluating the recommendation logic
    recommendations = []
    for _, row in fbm_candidates.iterrows():
        fbm_sku = row['SKU']
        
        # Enhanced logic:
        # A) Startswith + FBA
        # B) Contains + FBA
        # C) Normalized root match
        
        norm_fbm = normalize_sku(fbm_sku)
        
        match = None
        for s in fba_skus:
            if not s.endswith('FBA'): continue
            if fbm_sku in s: 
                match = s
                break
            if normalize_sku(s) == norm_fbm:
                match = s
                break
        
        if not match:
            recommendations.append(fbm_sku)
        else:
            # Already exists in FBA as 'match'
            pass

    print(f"Final recommendations count: {len(recommendations)}")
    return recommendations

if __name__ == "__main__":
    audit()
