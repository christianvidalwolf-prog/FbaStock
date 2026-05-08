#!/bin/bash

export PATH="/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin:$PATH"

cd /Users/christianvidalwolf/Stock

echo "--- CRON JOB STARTED AT $(date) ---" >> /Users/christianvidalwolf/Stock/cron_update.log

# 1. Download SellerBoard daily reports to USB SSD
echo "[$(date)] Downloading SellerBoard reports..." >> /Users/christianvidalwolf/Stock/cron_update.log
/Library/Developer/CommandLineTools/usr/bin/python3 /Users/christianvidalwolf/Stock/download_sellerboard_reports.py >> /Users/christianvidalwolf/Stock/cron_update.log 2>&1

# 2. Download DCASA via curl (Native, more robust)
echo "[$(date)] Downloading DCASA files with curl..." >> /Users/christianvidalwolf/Stock/cron_update.log
bash /Users/christianvidalwolf/Stock/download_dcasa_curl.sh >> /Users/christianvidalwolf/Stock/cron_update.log 2>&1

# 3. Run Python Update
/Library/Developer/CommandLineTools/usr/bin/python3 /Users/christianvidalwolf/Stock/auto_update_stock_fast.py >> /Users/christianvidalwolf/Stock/cron_update.log 2>&1

# 4. Update FBA app data from today's SellerBoard snapshots
echo "[$(date)] Syncing FBA app data..." >> /Users/christianvidalwolf/Stock/cron_update.log
/Library/Developer/CommandLineTools/usr/bin/python3 /Users/christianvidalwolf/Stock/sync_fba_report.py >> /Users/christianvidalwolf/Stock/cron_update.log 2>&1

# 5. Update Web Stock Excel
echo "[$(date)] Updating Web Stock Excel..." >> /Users/christianvidalwolf/Stock/cron_update.log
/Library/Developer/CommandLineTools/usr/bin/python3 /Users/christianvidalwolf/Stock/update_web_stock.py >> /Users/christianvidalwolf/Stock/cron_update.log 2>&1

# 6. Update Cdiscount Stock Excel
echo "[$(date)] Updating Cdiscount Stock Excel..." >> /Users/christianvidalwolf/Stock/cron_update.log
/Library/Developer/CommandLineTools/usr/bin/python3 /Users/christianvidalwolf/Stock/update_cdisc_stock.py >> /Users/christianvidalwolf/Stock/cron_update.log 2>&1

# 7. Copy daily stock files to Dropbox
echo "[$(date)] Backing up stock files to Dropbox..." >> /Users/christianvidalwolf/Stock/cron_update.log
/Library/Developer/CommandLineTools/usr/bin/python3 /Users/christianvidalwolf/Stock/backup_stock_files.py >> /Users/christianvidalwolf/Stock/cron_update.log 2>&1

echo "--- CRON JOB FINISHED AT $(date) ---" >> /Users/christianvidalwolf/Stock/cron_update.log
