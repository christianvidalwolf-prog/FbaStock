import os
from datetime import datetime

import requests


SELLERBOARD_DIR = "/Volumes/USB SSD/Ficheros sellerboard"

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
    if not os.path.isdir(SELLERBOARD_DIR):
        raise RuntimeError(
            f"No encuentro la carpeta de SellerBoard: {SELLERBOARD_DIR}. "
            "Comprueba que el USB SSD este montado."
        )


def download_report(prefix, url, date_str):
    output_path = os.path.join(SELLERBOARD_DIR, f"{prefix}_{date_str}.csv")
    print(f"Downloading {prefix}...")

    response = requests.get(url, headers=HEADERS, timeout=90)
    response.raise_for_status()

    content = response.content
    if not content.strip():
        raise RuntimeError(f"SellerBoard returned an empty file for {prefix}.")

    with open(output_path, "wb") as f:
        f.write(content)

    print(f"Saved {prefix}: {output_path} ({len(content)} bytes)")
    return output_path


def main():
    print(f"--- Downloading SellerBoard reports @ {datetime.now()} ---")
    ensure_destination()

    date_str = datetime.now().strftime("%Y-%m-%d")
    for prefix, url in REPORTS:
        download_report(prefix, url, date_str)

    print("SellerBoard downloads finished.")


if __name__ == "__main__":
    main()
