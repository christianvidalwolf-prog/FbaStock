import csv
import tempfile
import unittest
from pathlib import Path

import auto_update_stock_fast as stock


def write_catalog(rows):
    path = Path(tempfile.mkdtemp()) / "md.csv"
    with open(path, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["sku", "price"])
        writer.writerows(rows)
    return path


class TrediserMdSkuTests(unittest.TestCase):
    def test_split_returns_numeric_code_for_md_and_mdcar(self):
        self.assertEqual(stock.split_trediser_md_sku("42608MD"), ("MD", "42608"))
        self.assertEqual(stock.split_trediser_md_sku("30227MDCAR"), ("MDCAR", "30227"))

    def test_split_rejects_other_suffixes(self):
        for sku in ("42608MDRG", "42608", "MD42608", "42608XX"):
            self.assertEqual(stock.split_trediser_md_sku(sku), (None, None))


class TrediserMdCatalogTests(unittest.TestCase):
    def test_loads_approved_prices(self):
        path = write_catalog([["42608MD", "24.4"], ["30227MDCAR", "71,4"]])
        self.assertEqual(stock.load_trediser_md_prices(path),
                         {"42608MD": 24.4, "30227MDCAR": 71.4})

    def test_rejects_invalid_sku_duplicate_and_non_positive_price(self):
        for rows in ([["42608XX", "10"]],
                     [["42608MD", "10"], ["42608MD", "12"]],
                     [["42608MD", "0"]],
                     [["42608MD", "-1"]]):
            with self.assertRaises(ValueError):
                stock.load_trediser_md_prices(write_catalog(rows))

    def test_rejects_header_only_catalog(self):
        with self.assertRaises(ValueError):
            stock.load_trediser_md_prices(write_catalog([]))

    def test_missing_file_keeps_pipeline_working(self):
        self.assertEqual(stock.load_trediser_md_prices("/nonexistent/md.csv"), {})


if __name__ == "__main__":
    unittest.main()
