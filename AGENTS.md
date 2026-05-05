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

# Or run directly
python3 auto_update_stock_fast.py
```

### Data Files
- `fba-replenishment/public/data.json` - Frontend reads this
- `STOCK AMZ.txt` - Generated stock data
- `catalog_robust.json` - Product catalog
- `signes_stock.csv`, `minerales_feed.xml` - Supplier data

## Key Notes

- Python scripts use system Python3 (`/Library/Developer/CommandLineTools/usr/bin/python3`)
- Cron job runs daily via `run_stock_update.sh` (see plist in `com.christianvidalwolf.stockupdate.plist`)
- The React app displays FBA inventory data with filtering/sorting on all table columns
- `fbm_recommendations` array in data.json is currently empty - needs a separate analysis process to populate