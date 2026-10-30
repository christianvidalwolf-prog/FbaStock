import tempfile
import unittest
from pathlib import Path
from zipfile import ZipFile

import openpyxl
from amazon_stock_workbook import export_amazon_workbook, TEMPLATE_FILE, SKU, QUANTITY

CHANNEL = 'fulfillment_availability#1.fulfillment_channel_code'
PRICE = 'purchasable_offer[marketplace_id=A1RKKUPIHCS9HS][audience=ALL]#1.our_price#1.schedule#1.value_with_tax'


class WorkbookExportTests(unittest.TestCase):
    def test_add_remove_reorder_and_preserve_template(self):
        with tempfile.TemporaryDirectory() as d:
            out = Path(d)/'daily.xlsm'
            # Source column order deliberately differs from the Amazon workbook.
            headers = [QUANTITY, PRICE, SKU, CHANNEL]
            rows = [['0', '12,95', 'new<&SKU', 'Logística por parte del vendedor (predeterminado)'],
                    ['42', '109,8', '231VC', 'Logística por parte del vendedor (predeterminado)']]
            self.assertEqual(export_amazon_workbook(rows, headers, out), {'skus': 2, 'with_stock': 1})
            wb = openpyxl.load_workbook(out, read_only=True, data_only=True)
            sheet = wb['Plantilla']
            self.assertEqual(sheet.max_row, 8)
            self.assertEqual(sheet['D7'].value, 'new<&SKU')
            self.assertEqual(sheet['F7'].value, 0)
            self.assertEqual(sheet['J7'].value, 12.95)
            self.assertEqual(sheet['D8'].value, '231VC')
            self.assertEqual(sheet['F8'].value, 42)
            self.assertIsNone(sheet['A8'].value)
            self.assertEqual(sheet['D6'].value, 'ABC123')
            wb.close()
            with ZipFile(TEMPLATE_FILE) as original, ZipFile(out) as generated:
                self.assertEqual(original.namelist(), generated.namelist())
                self.assertEqual([n for n in original.namelist() if original.read(n) != generated.read(n)],
                                 ['xl/worksheets/sheet6.xml'])
            before = out.read_bytes()
            for invalid in [[], rows + [rows[0]], [['-1', '12', 'bad', 'DEFAULT']]]:
                with self.assertRaises(ValueError):
                    export_amazon_workbook(invalid, headers, out)
                self.assertEqual(out.read_bytes(), before)


if __name__ == '__main__':
    unittest.main()
