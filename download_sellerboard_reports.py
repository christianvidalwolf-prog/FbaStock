import os
from datetime import datetime

import requests


USB_SELLERBOARD_DIR = "/Volumes/USB SSD/Ficheros sellerboard"
BACKUP_SELLERBOARD_DIR = "/Users/christianvidalwolf/Stock/sellerboard_backups"

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
    os.makedirs(BACKUP_SELLERBOARD_DIR, exist_ok=True)
    if os.path.isdir(USB_SELLERBOARD_DIR):
        return USB_SELLERBOARD_DIR
    print(
        f"USB SSD no disponible en {USB_SELLERBOARD_DIR}. "
        f"Usando respaldo local en {BACKUP_SELLERBOARD_DIR}."
    )
    return BACKUP_SELLERBOARD_DIR


def write_snapshot(prefix, date_str, content, destination_dir):
    output_path = os.path.join(destination_dir, f"{prefix}_{date_str}.csv")
    with open(output_path, "wb") as f:
        f.write(content)
    return output_path


def download_report(prefix, url, date_str):
    print(f"Downloading {prefix}...")

    response = requests.get(url, headers=HEADERS, timeout=90)
    response.raise_for_status()

    content = response.content
    if not content.strip():
        raise RuntimeError(f"SellerBoard returned an empty file for {prefix}.")

    destinations = []
    if os.path.isdir(USB_SELLERBOARD_DIR):
        destinations.append(USB_SELLERBOARD_DIR)
    destinations.append(BACKUP_SELLERBOARD_DIR)

    last_error = None
    for destination_dir in destinations:
        try:
            output_path = write_snapshot(prefix, date_str, content, destination_dir)
            print(f"Saved {prefix}: {output_path} ({len(content)} bytes)")
            return output_path
        except OSError as exc:
            last_error = exc
            print(f"Could not save {prefix} to {destination_dir}: {exc}")

    raise last_error or RuntimeError(f"Could not save SellerBoard snapshot for {prefix}.")


def main():
    print(f"--- Downloading SellerBoard reports @ {datetime.now()} ---")
    ensure_destination()

    date_str = datetime.now().strftime("%Y-%m-%d")
    for prefix, url in REPORTS:
        download_report(prefix, url, date_str)

    print("SellerBoard downloads finished.")


if __name__ == "__main__":
    main()
