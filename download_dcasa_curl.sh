#!/bin/bash
# Robust DCASA downloader using curl (native to macOS)
# No sudo or lftp installation required.

HOST="data.dcasacollection.com"
USER="sek4283"
PASS="Rx34z5m6_pER"
WORK_DIR="/Users/christianvidalwolf/Stock"

cd "$WORK_DIR" || exit 1

echo "[$(date)] Syncing DCASA data via curl..."

# 1. Fetch the file list to find the latest DataWeb*.csv
# We use --connect-timeout to avoid hanging and sort to get the latest by date
echo "Connecting to $HOST..."
LATEST=$(curl -s --connect-timeout 60 -u "$USER:$PASS" -l "ftp://$HOST/" | grep "DataWeb" | grep "\.csv$" | sort | tail -n 1)

if [ -z "$LATEST" ]; then
    echo "Error: Could not find any DataWeb CSV files on DCASA FTP (Connection timeout or server down)."
    exit 1
fi

echo "Target file identified: $LATEST"

# 2. Download the file with robust retry and resume support
# --retry 10: try up to 10 times on connection failure
# --retry-delay 10: wait 10s between retries
# -C -: resume partial downloads if the connection breaks
# --connect-timeout 60: don't wait forever to connect
echo "Starting download: $LATEST"
curl -u "$USER:$PASS" \
    --retry 10 \
    --retry-delay 10 \
    --connect-timeout 60 \
    --max-time 600 \
    -C - \
    -O "ftp://$HOST/$LATEST"

if [ $? -eq 0 ]; then
    echo "Successfully downloaded/synced: $LATEST"
else
    echo "Failed to download $LATEST after retries."
    exit 1
fi
