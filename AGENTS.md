# Stock Management System

## Project Structure

- **fba-replenishment/** - React frontend (Vite + React 19 + TypeScript + Tailwind)
- **Python scripts** - Data processing in root directory
- **Data sources**: CSV/JSON files, supplier feeds (Signes FTP, Minerales XML, DCASA curl)

## Developer Commands

### React App (fba-replenishment/)
```bash
cd fba-replenishment
npm run dev      # Start dev server
npm run build    # Production build (tsc -b && vite build)
npm run lint     # ESLint check
```

### Python Scripts (run from Stock root)
```bash
# Full daily update (includes data download + processing)
./run_stock_update.sh

# Run sync script (main data processor)
python3 sync_fba_report.py

# Translate titles in existing data.json
python3 translate_titles.py
```

### Data Files
- `fba-replenishment/public/data.json` - Frontend reads this
- `fba-replenishment/public/` - Static assets served by Vite
- `/Volumes/USB SSD/ficheros sellerboard/` - Historical data storage

## Historical Data (USB SSD)

The script automatically downloads and saves:
- **Inventario**: `sellerboard_inventory_YYYY-MM-DD.csv`
- **Ventas**: `sellerboard_ventas_YYYY-MM-DD.csv`

Sales history is built progressively:
- 60-day sales: aggregated from all available CSV files
- 365-day sales: aggregated from all available CSV files

Run `sync_fba_report.py` daily to build the historical record.

## Key Logic Rules

- **No reorder if ASIN has stock**: If any SKU of an ASIN already has FBA stock, other SKUs of that ASIN won't be marked for reorder (see `sync_fba_report.py` asins_with_stock logic)
- **Title translation**: Automatically translates non-Spanish titles to Spanish using deep-translator. Runs on sync or via `translate_titles.py`
- **Excluded SKUs**: Products can be excluded in any tab. State persists in localStorage key `excluded_skus`

## Data Pipeline

1. `sync_fba_report.py` fetches from multiple sources:
   - SellerBoard CSV (FBA inventory) - saved to USB SSD
   - SellerBoard sales (daily) - saved to USB SSD
   - DCASA, Signes, Minerales supplier feeds
2. Generates `data.json` with products + fbm_recommendations + summary
3. React app reads `data.json` for all tabs

## Important Notes

- Python scripts use system Python3 (`/Library/Developer/CommandLineTools/usr/bin/python3`)
- Cron job runs daily via `run_stock_update.sh`
- Translation requires: `pip install deep-translator`
- Tables support column sorting (click header) and filtering (input under header name)