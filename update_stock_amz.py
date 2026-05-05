import re

input_file = "/Users/christianvidalwolf/Stock/STOCK AMZ.txt"
output_file = "/Users/christianvidalwolf/Stock/STOCK AMZ.txt"


def parse_price(price_str):
    return float(price_str.replace(",", "."))


def format_price(price):
    return str(price).replace(".", ",")


def round_to_95(price):
    new_price = price * 1.05
    whole = int(new_price)
    decimal = new_price - whole
    if decimal <= 0.50:
        rounded = whole + 0.95
    else:
        rounded = whole + 1.95
    return round(rounded, 2)


with open(input_file, "r", encoding="utf-8") as f:
    lines = f.readlines()

header = lines[0].strip()
updated_count = 0
output_lines = [header]

for line in lines[1:]:
    line = line.strip()
    if not line:
        continue

    parts = line.split("\t")
    if len(parts) < 2:
        continue

    sku = parts[0]
    if not sku.endswith("VC"):
        output_lines.append(line.replace("\t", "\t"))
        continue

    original_price = parse_price(parts[1])

    if original_price < 30:
        new_price = round_to_95(original_price)
        min_price = round(new_price * 0.5, 2)
        max_price = round(new_price * 2, 2)

        parts[1] = format_price(new_price)
        parts[2] = format_price(min_price)
        parts[3] = format_price(max_price)

        updated_count += 1
        print(f"Updated {sku}: {original_price} -> {new_price}")

    output_lines.append("\t".join(parts))

with open(output_file, "w", encoding="utf-8") as f:
    f.write("\n".join(output_lines))

print(f"\nTotal products updated: {updated_count}")
