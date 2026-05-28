import os
import glob
from datetime import datetime
import time

import requests


LOCAL_SELLERBOARD_DIR = "/Users/christianvidalwolf/Stock/sellerboard_backups"
USB_SELLERBOARD_DIR = "/Volumes/USB SSD/Ficheros sellerboard"
RETENTION_DAYS = 60
RETENTION_SECONDS = RETENTION_DAYS * 24 * 60 * 60

REPORTS = [
    (
        "sellerboard_inventory",
        "https://app.sellerboard.com/es/automation/reports?id=a1a2f4284b8043c39964edfe3cef86ca&format=csv&t=bbc9d347dff7407dbd01c90884f31121",
    ),
    (
        "sellerboard_ventas",
        "https://app.sellerboard.com/es/automation/reports?id=a258a124dd524541be35028b6a172013&format=csv&t=bbc9d347dff7407dbd01c90884f31121",
    ),
]

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
        "AppleWebKit/537.36 Chrome/120 Safari/537.36"
    )
}


def ensure_destination():
    os.makedirs(LOCAL_SELLERBOARD_DIR, exist_ok=True)
    return LOCAL_SELLERBOARD_DIR


def write_snapshot(prefix, date_str, content, destination_dir):
    output_path = os.path.join(destination_dir, f"{prefix}_{date_str}.csv")
    with open(output_path, "wb") as f:
        f.write(content)
    return output_path


def cleanup_old_snapshots():
    cutoff = time.time() - RETENTION_SECONDS
    patterns = ["sellerboard_inventory_*.csv", "sellerboard_ventas_*.csv"]
    for directory in [LOCAL_SELLERBOARD_DIR, USB_SELLERBOARD_DIR]:
        for pattern in patterns:
            for path in glob.glob(os.path.join(directory, pattern)):
                try:
                    if os.path.getmtime(path) < cutoff:
                        os.remove(path)
                        print(f"Deleted old snapshot: {path}")
                except OSError as exc:
                    print(f"Could not delete {path}: {exc}")


def download_report(prefix, url, date_str):
    print(f"Downloading {prefix}...")

    max_retries = 5
    for attempt in range(max_retries):
        try:
            response = requests.get(url, headers=HEADERS, timeout=90)
            response.raise_for_status()

            content = response.content
            text = content.decode("utf-8-sig", errors="replace")

            # Check if SellerBoard report is ready
            if "Report not ready" in text or "try again" in text:
                print(f"  Attempt {attempt+1}/{max_retries}: SellerBoard report not ready yet. Waiting 60s...")
                time.sleep(60)
                continue

            if not content.strip():
                raise RuntimeError(f"SellerBoard returned an empty file for {prefix}.")

            # Validate that this is a valid CSV report by checking for headers
            if "ASIN" not in text or "SKU" not in text:
                raise RuntimeError(f"SellerBoard report does not contain expected CSV headers (ASIN, SKU). Preview: {text[:200]}")

            destinations = [LOCAL_SELLERBOARD_DIR]
            if os.path.isdir(USB_SELLERBOARD_DIR):
                destinations.append(USB_SELLERBOARD_DIR)

            last_error = None
            for destination_dir in destinations:
                try:
                    output_path = write_snapshot(prefix, date_str, content, destination_dir)
                    print(f"Saved {prefix}: {output_path} ({len(content)} bytes)")
                except OSError as exc:
                    last_error = exc
                    print(f"Could not save {prefix} to {destination_dir}: {exc}")

            if last_error:
                raise last_error
            return output_path

        except Exception as e:
            if attempt == max_retries - 1:
                raise
            print(f"  Attempt {attempt+1}/{max_retries} failed: {e}. Retrying in 30s...")
            time.sleep(30)



def main():
    print(f"--- Downloading SellerBoard reports @ {datetime.now()} ---")
    ensure_destination()

    date_str = datetime.now().strftime("%Y-%m-%d")
    for prefix, url in REPORTS:
        download_report(prefix, url, date_str)

    cleanup_old_snapshots()
    print("SellerBoard downloads finished.")


if __name__ == "__main__":
    main()
