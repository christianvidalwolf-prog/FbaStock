import xlwings as xw
import datetime
import time

today_str = datetime.datetime.now().strftime('%Y%m%d')
master_file = '/Users/christianvidalwolf/Stock/INICIO PLUS 2023.xlsx'
output_file = f'/Users/christianvidalwolf/Stock/STOCK AMZ {today_str}.txt'
sheet_name = 'STOCK ES'

print(f"Opening {master_file} via Excel (visible, full recalc)...")
app = xw.App(visible=True)
app.display_alerts = False
try:
    wb = app.books.open(master_file, update_links=True)

    # Force full recalculation of all sheets
    app.calculation = 'automatic'
    app.calculate()
    time.sleep(3)  # Give Excel time to finish

    ws = wb.sheets[sheet_name]

    print(f"Reading values from '{sheet_name}'...")
    data = ws.used_range.value

    rows = []
    for row in data:
        if isinstance(row, list):
            if any(v is not None and v != '' for v in row):
                rows.append(row)
        elif row is not None and row != '':
            rows.append([row])

    print(f"Writing {len(rows)} rows to {output_file}...")
    with open(output_file, 'w', encoding='utf-8') as f:
        for row in rows:
            line = '\t'.join('' if v is None else str(v) for v in row)
            f.write(line + '\n')

    print(f"Done: {output_file}")
    # Show first few lines
    with open(output_file) as f:
        for i, line in enumerate(f):
            if i >= 5:
                break
            print(line.rstrip())
finally:
    wb.close()
    app.quit()
