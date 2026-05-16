#!/bin/bash

export PATH="/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin:$PATH"
PYTHON=/usr/bin/python3
LOG_FILE="/Users/christianvidalwolf/Stock/logs/cron_update.log"

cd /Users/christianvidalwolf/Stock
mkdir -p /Users/christianvidalwolf/Stock/logs

echo "--- CRON JOB STARTED AT $(date) ---" >> "$LOG_FILE"

# 1. Download SellerBoard daily reports to USB SSD
echo "[$(date)] Downloading SellerBoard reports..." >> "$LOG_FILE"
$PYTHON /Users/christianvidalwolf/Stock/download_sellerboard_reports.py >> "$LOG_FILE" 2>&1

# 2. Download DCASA via curl (Native, more robust)
echo "[$(date)] Downloading DCASA files with curl..." >> "$LOG_FILE"
bash /Users/christianvidalwolf/Stock/download_dcasa_curl.sh >> "$LOG_FILE" 2>&1

# 3. Run Python Update
$PYTHON /Users/christianvidalwolf/Stock/auto_update_stock_fast.py >> "$LOG_FILE" 2>&1

# 4. Update FBA app data from today's SellerBoard snapshots
echo "[$(date)] Syncing FBA app data..." >> "$LOG_FILE"
$PYTHON /Users/christianvidalwolf/Stock/sync_fba_report.py >> "$LOG_FILE" 2>&1

# 5. Update Web Stock Excel
echo "[$(date)] Updating Web Stock Excel..." >> "$LOG_FILE"
$PYTHON /Users/christianvidalwolf/Stock/update_web_stock.py >> "$LOG_FILE" 2>&1

# 6. Update Cdiscount Stock Excel
echo "[$(date)] Updating Cdiscount Stock Excel..." >> "$LOG_FILE"
$PYTHON /Users/christianvidalwolf/Stock/update_cdisc_stock.py >> "$LOG_FILE" 2>&1

# 7. Copy daily stock files to Dropbox
echo "[$(date)] Backing up stock files to Dropbox..." >> "$LOG_FILE"
$PYTHON /Users/christianvidalwolf/Stock/backup_stock_files.py >> "$LOG_FILE" 2>&1

echo "--- CRON JOB FINISHED AT $(date) ---" >> "$LOG_FILE"
