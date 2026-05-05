import pandas as pd

try:
    df = pd.read_excel('Nuevos SG1.xlsx')
    print("Columns:", df.columns.tolist())
    print("\nFirst 5 rows:")
    print(df.head())
    
    # Filter for Signes products (SKU ends in SGRI or something similar?)
    # The user said SKU ends in SGRI.
    # Let's check if there's an SKU column.
    sku_col = None
    for col in df.columns:
        if 'sku' in col.lower() or 'referencia' in col.lower() or 'producto' in col.lower():
            sku_col = col
            break
    
    if sku_col:
        print(f"\nFiltering by SKU column: {sku_col}")
        signes_df = df[df[sku_col].astype(str).str.endswith('SGRI')]
        print(f"Found {len(signes_df)} items ending in SGRI")
        if not signes_df.empty:
            print("\nFirst 5 Signes rows:")
            print(signes_df.head())
    else:
        print("\nCould not find SKU column automatically.")

except Exception as e:
    print(f"Error: {e}")
