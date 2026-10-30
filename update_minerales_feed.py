#!/usr/bin/env python3
"""
update_minerales_feed.py
Converts manual CSV files for Minerales (VC) into the standard minerales_feed.xml
used across auto_update_stock_fast.py, sync_fba_report.py, update_web_stock.py, and update_cdisc_stock.py.
"""
import os
import glob
import csv
import shutil

WORK_DIR = "/Users/christianvidalwolf/Stock"
MINERALES_FEED = os.path.join(WORK_DIR, "minerales_feed.xml")


def find_latest_manual_file():
    """Finds the latest manual Minerales CSV file in WORK_DIR or Downloads."""
    candidates = []

    # Check WORK_DIR for manual files
    patterns = [
        os.path.join(WORK_DIR, "minerales_stock.csv"),
        os.path.join(WORK_DIR, "minerales_stock_*.csv"),
        os.path.join(WORK_DIR, "minerales_*.csv"),
        os.path.join(WORK_DIR, "MINERALES*.csv"),
        os.path.join(WORK_DIR, "request_sql*.csv"),
    ]
    for pattern in patterns:
        for f in glob.glob(pattern):
            if os.path.isfile(f) and os.path.getsize(f) > 500:
                candidates.append((os.path.getmtime(f), f))

    # Also check Downloads for recent request_sql files
    downloads_dir = os.path.expanduser("~/Downloads")
    if os.path.exists(downloads_dir):
        for f in glob.glob(os.path.join(downloads_dir, "request_sql*.csv")):
            if os.path.isfile(f) and os.path.getsize(f) > 500:
                candidates.append((os.path.getmtime(f), f))

    if not candidates:
        return None

    # Return the one with highest mtime
    candidates.sort(key=lambda x: x[0], reverse=True)
    return candidates[0][1]


def convert_manual_csv_to_feed(manual_file, feed_target=MINERALES_FEED):
    """
    Reads the manual CSV and writes out minerales_feed.xml in pipe format.
    Format:
      parts[0] = Referencia (ID)
      parts[1] = Nombre
      parts[2..6] = ''
      parts[7] = Cantidad (Stock)
      parts[8..11] = ''
      parts[12] = Precio (imp. excl.)
    """
    if not os.path.exists(manual_file):
        print(f"Error: Manual file does not exist: {manual_file}")
        return False

    print(f"Reading manual Minerales file: {manual_file}")

    # Detect delimiter and encoding
    encoding = "utf-8"
    try:
        with open(manual_file, "r", encoding=encoding) as f:
            sample = f.read(4096)
    except UnicodeDecodeError:
        encoding = "latin-1"
        with open(manual_file, "r", encoding=encoding) as f:
            sample = f.read(4096)

    delimiter = ";" if ";" in sample else ("," if "," in sample else "\t")

    items_by_ref = {}
    with open(manual_file, "r", encoding=encoding, errors="ignore") as f:
        reader = csv.reader(f, delimiter=delimiter)
        header = next(reader, None)

        # Determine column indexes
        ref_idx = 1
        qty_idx = 4
        name_idx = 0
        price_idx = 3

        if header:
            norm_header = [col.strip().lower() for col in header]
            for idx, col in enumerate(norm_header):
                if "referencia" in col or col == "ref" or col == "id":
                    ref_idx = idx
                elif "cantidad" in col or "stock" in col or "qty" in col:
                    qty_idx = idx
                elif "nombre" in col or "title" in col or "name" in col:
                    name_idx = idx
                elif "precio" in col or "price" in col:
                    price_idx = idx

        for row in reader:
            if len(row) <= max(ref_idx, qty_idx):
                continue

            ref_val = row[ref_idx].strip()
            if not ref_val or ref_val.lower() == "referencia":
                continue

            # Clean reference
            clean_ref = ref_val.replace(".0", "").strip()

            # Clean quantity
            qty_raw = row[qty_idx].strip().replace(",", ".")
            try:
                stock = int(float(qty_raw))
            except ValueError:
                stock = 0

            if clean_ref == "2162" and stock == 91:
                stock = 0

            # Clean name and price if available
            name = row[name_idx].strip() if len(row) > name_idx else ""
            price = row[price_idx].strip().replace(",", ".") if len(row) > price_idx else "0.00"

            # Build standard 14-field pipe line
            # parts[0]=ref, parts[1]=name, parts[7]=stock, parts[12]=price
            parts = [clean_ref, name, "", "", "", "", "", str(stock), "", "", "", "", price, ""]
            items_by_ref[clean_ref] = "|".join(parts)

    if not items_by_ref:
        print("Warning: No valid items found in manual file!")
        return False

    # Write out to feed_target safely via tmp file
    tmp_target = feed_target + ".tmp"
    with open(tmp_target, "w", encoding="utf-8") as f:
        for line in items_by_ref.values():
            f.write(line + "\n")

    shutil.move(tmp_target, feed_target)

    stock_gt_0 = sum(1 for line in items_by_ref.values() if int(line.split("|")[7]) > 0)
    stock_ge_5 = sum(1 for line in items_by_ref.values() if int(line.split("|")[7]) >= 5)

    print(f"Successfully generated {feed_target}:")
    print(f"  Total items: {len(items_by_ref)}")
    print(f"  Items with stock > 0: {stock_gt_0}")
    print(f"  Items with stock >= 5 (Amazon threshold): {stock_ge_5}")

    # Ensure a copy is in WORK_DIR as minerales_stock.csv
    work_stock_csv = os.path.join(WORK_DIR, "minerales_stock.csv")
    if os.path.abspath(manual_file) != os.path.abspath(work_stock_csv):
        try:
            shutil.copy2(manual_file, work_stock_csv)
            print(f"  Synchronized copy to {work_stock_csv}")
        except Exception as e:
            print(f"  Could not copy to {work_stock_csv}: {e}")

    return True


if __name__ == "__main__":
    import sys
    manual_file = sys.argv[1] if len(sys.argv) > 1 else find_latest_manual_file()
    if not manual_file:
        print("No manual Minerales file found in Stock/ or Downloads/.")
        sys.exit(1)

    success = convert_manual_csv_to_feed(manual_file)
    sys.exit(0 if success else 1)
