#!/bin/bash

export PATH="/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin:$PATH"

cd /Users/christianvidalwolf/Stock

echo "--- CRON JOB STARTED AT $(date) ---" >> /Users/christianvidalwolf/Stock/cron_update.log

# 1. Download DCASA via curl (Native, more robust)
echo "[$(date)] Downloading DCASA files with curl..." >> /Users/christianvidalwolf/Stock/cron_update.log
bash /Users/christianvidalwolf/Stock/download_dcasa_curl.sh >> /Users/christianvidalwolf/Stock/cron_update.log 2>&1

# 2. Run Python Update
/Library/Developer/CommandLineTools/usr/bin/python3 /Users/christianvidalwolf/Stock/auto_update_stock_fast.py >> /Users/christianvidalwolf/Stock/cron_update.log 2>&1

echo "--- CRON JOB FINISHED AT $(date) ---" >> /Users/christianvidalwolf/Stock/cron_update.log
