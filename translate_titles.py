#!/usr/bin/env python3
"""Translate titles in data.json to Spanish."""

import json
import time
from sync_fba_report import translate_to_spanish, is_likely_spanish

DATA_FILE = "/Users/christianvidalwolf/Stock/fba-replenishment/public/data.json"


def main():
    print("Loading data.json...")
    with open(DATA_FILE) as f:
        data = json.load(f)

    products = data.get("products", [])
    print(f"Processing {len(products)} products...")

    translated = 0
    skipped = 0

    for i, p in enumerate(products):
        title = p.get("title", "")
        if not title or len(title) < 10:
            skipped += 1
            continue

        if not is_likely_spanish(title):
            original = title
            p["title"] = translate_to_spanish(title)
            translated += 1
            if translated % 10 == 0:
                print(f"  Translated {translated} titles... ({i + 1}/{len(products)})")

        if (i + 1) % 100 == 0:
            print(f"  Progress: {i + 1}/{len(products)}")

    print(f"\nDone: {translated} translated, {skipped} skipped")

    # Save
    with open(DATA_FILE, "w") as f:
        json.dump(data, f, indent=2)

    print(f"Saved to {DATA_FILE}")


if __name__ == "__main__":
    main()
