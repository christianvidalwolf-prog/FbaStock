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
- `/Volumes/USB SSD/Ficheros sellerboard/` - Primary SellerBoard snapshot storage
- `/Users/christianvidalwolf/Stock/sellerboard_backups/` - Primary local SellerBoard snapshots and history
- `/Users/christianvidalwolf/Stock/logs/` - Local cron logs

## Historical Data

Primary storage is the local `sellerboard_backups/` folder. If the USB SSD is mounted, the scripts also mirror snapshots there.

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
   - SellerBoard CSV (FBA inventory) - saved to local storage, mirrored to USB if available
   - SellerBoard sales (daily) - saved to local storage, mirrored to USB if available
   - DCASA, Signes, Minerales supplier feeds
2. Generates `data.json` with products + fbm_recommendations + summary
3. React app reads `data.json` for all tabs

## Important Notes

- Python scripts use system Python3 (`/Library/Developer/CommandLineTools/usr/bin/python3`)
- Cron job runs daily via `run_stock_update.sh`
- Cron output lives in `/Users/christianvidalwolf/Stock/logs/` and does not depend on the USB SSD
- Translation requires: `pip install deep-translator`
- Tables support column sorting (click header) and filtering (input under header name)
