#!/usr/bin/env python3
import csv

# Read the file
with open(
    "/Users/christianvidalwolf/Stock/minerales_feed.xml", "r", encoding="utf-8"
) as f:
    content = f.read()

# Split into lines (skip header)
lines = content.strip().split("\n")
header = lines[0]
data_lines = lines[1:]

# Process each line
updated_count = 0
updates = []

for line in data_lines:
    parts = line.split("|")
    if len(parts) >= 13:
        product_id = parts[0]
        title = parts[1]
        price_str = parts[12].strip()

        # Check if price is a valid number
        if price_str and price_str.replace(".", "").isdigit():
            current_price = float(price_str)

            # Only update if price < 30
            if current_price < 30:
                # Apply 5% increase
                new_price = current_price * 1.05

                # Round to end in .95
                integer_part = int(new_price)
                new_price_rounded = integer_part + 0.95

                # Update the price
                parts[12] = str(new_price_rounded)
                updated_count += 1
                updates.append((product_id, title, current_price, new_price_rounded))

# Reconstruct the file
new_content = (
    header
    + "\n"
    + "\n".join(
        [
            "|".join(parts)
            for parts in [header.split("|")] + [line.split("|") for line in data_lines]
        ]
    )
)

# Better approach - rebuild line by line
output_lines = [header]
for i, line in enumerate(data_lines):
    parts = line.split("|")
    if len(parts) >= 13:
        product_id = parts[0]
        price_str = parts[12].strip()

        if price_str and price_str.replace(".", "").isdigit():
            current_price = float(price_str)

            if current_price < 30:
                new_price = current_price * 1.05
                integer_part = int(new_price)
                new_price_rounded = integer_part + 0.95
                parts[12] = str(round(new_price_rounded, 2))

    output_lines.append("|".join(parts))

# Write back
with open(
    "/Users/christianvidalwolf/Stock/minerales_feed.xml", "w", encoding="utf-8"
) as f:
    f.write("\n".join(output_lines))

print(f"Total de productos actualizados: {updated_count}")
print("\nDetalles de actualizaciones:")
for u in updates:
    print(f"  ID {u[0]}: {u[1][:50]}... | {u[2]:.2f} -> {u[3]:.2f}")
