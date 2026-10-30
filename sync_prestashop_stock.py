"""Sube el stock de Stock Web.csv a PrestaShop (www.vidalregals.com).

Se ejecuta cada mañana desde run_stock_update.sh justo después de update_web_stock.py.
Cruza SKU (columna 0) con la referencia de productos y combinaciones de la tienda y solo
actualiza los stock_availables cuya cantidad ha cambiado. Los productos de la tienda que no
están en Stock Web.csv no se tocan.

Uso:
  python3 sync_prestashop_stock.py            # aplica cambios
  python3 sync_prestashop_stock.py --dry-run  # solo cuenta lo que cambiaría
"""
import csv
import json
import os
import sys
import time
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime
from xml.sax.saxutils import escape

import requests

WORK_DIR = "/Users/christianvidalwolf/Stock"
STOCK_WEB_CSV = os.path.join(WORK_DIR, "Stock Web.csv")
ENV_FILE = "/Users/christianvidalwolf/AMAZON AWS SP-API/.env"  # PS_URL, PS_API_KEY
LOG_JSONL = os.path.join(WORK_DIR, "logs", "prestashop_stock_sync.jsonl")
PAGE = 1000
WORKERS = 4


def load_env():
    env = {}
    with open(ENV_FILE, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                k, v = line.split("=", 1)
                env[k.strip()] = v.strip().strip('"').strip("'")
    return env["PS_URL"].rstrip("/"), env["PS_API_KEY"]


PS_URL, PS_KEY = load_env()
session = requests.Session()  # el hosting quita la cabecera Authorization: key por ?ws_key=


def ps(method, path, params=None, data=None):
    for attempt in range(5):
        try:
            r = session.request(method, f"{PS_URL}/api/{path}", params={"output_format": "JSON", "ws_key": PS_KEY, **(params or {})},
                                data=data, headers={"Content-Type": "application/xml"} if data else None, timeout=120)
            if r.status_code < 500:
                r.raise_for_status()
                return r.json() if r.text.strip() not in ("", "[]") else {}
        except requests.RequestException:
            if attempt == 4:
                raise
        time.sleep(2 ** attempt)
    raise RuntimeError(f"{method} {path} failed")


def paged(resource, display):
    rows, offset = [], 0
    while True:  # paginado: la consulta completa supera el timeout del hosting
        batch = ps("GET", resource, {"display": display, "limit": f"{offset},{PAGE}", "sort": "[id_ASC]"})
        batch = batch.get(resource, [])
        rows += batch
        if len(batch) < PAGE:
            return rows
        offset += PAGE


def load_stock_web():
    stock = {}
    with open(STOCK_WEB_CSV, encoding="utf-8") as f:
        for row in csv.reader(f, delimiter=";"):
            if len(row) >= 2 and row[0].strip() and row[0].strip() != "nan":
                try:
                    stock[row[0].strip().upper()] = int(float(row[1] or 0))
                except ValueError:
                    pass
    return stock


def put_stock(sa, qty):
    fields = {**sa, "quantity": qty}
    body = "".join(f"<{k}><![CDATA[{'' if v is None else v}]]></{k}>" for k, v in fields.items())
    xml = ('<?xml version="1.0" encoding="UTF-8"?><prestashop xmlns:xlink="http://www.w3.org/1999/xlink">'
           f"<stock_available>{body}</stock_available></prestashop>")
    ps("PUT", f"stock_availables/{sa['id']}", data=xml.encode("utf-8"))


def main():
    dry = "--dry-run" in sys.argv
    print(f"--- PrestaShop stock sync @ {datetime.now()} {'(dry-run)' if dry else ''} ---", flush=True)
    stock = load_stock_web()
    products = paged("products", "[id,reference]")
    combos = paged("combinations", "[id,id_product,reference]")
    availables = paged("stock_availables", "full")
    print(f"Stock Web: {len(stock)} SKUs | tienda: {len(products)} productos, {len(combos)} combinaciones, "
          f"{len(availables)} stock_availables", flush=True)

    # (id_product, id_product_attribute) -> SKU
    key_ref = {}
    for p in products:
        if p.get("reference"):
            key_ref[(str(p["id"]), "0")] = p["reference"].strip().upper()
    for c in combos:
        if c.get("reference"):
            key_ref[(str(c["id_product"]), str(c["id"]))] = c["reference"].strip().upper()

    changes, matched = [], 0
    for sa in availables:
        ref = key_ref.get((str(sa["id_product"]), str(sa["id_product_attribute"])))
        if ref is None or ref not in stock:
            continue
        matched += 1
        if int(sa["quantity"]) != stock[ref]:
            changes.append((sa, ref, stock[ref]))
    print(f"Cruzados: {matched} | a cambiar: {len(changes)}", flush=True)
    if dry:
        for sa, ref, qty in changes[:15]:
            print(f"  {ref} (producto {sa['id_product']}): {sa['quantity']} -> {qty}")
        return 0

    errors = 0
    os.makedirs(os.path.dirname(LOG_JSONL), exist_ok=True)

    def one(item):
        sa, ref, qty = item
        try:
            put_stock(sa, qty)
            return {"ref": ref, "id_product": sa["id_product"], "from": int(sa["quantity"]), "to": qty}
        except Exception as exc:  # seguimos con el resto
            return {"ref": ref, "id_product": sa["id_product"], "error": str(exc)[:300]}

    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with ThreadPoolExecutor(max_workers=WORKERS) as pool, open(LOG_JSONL, "a", encoding="utf-8") as log:
        for n, res in enumerate(pool.map(one, changes), 1):
            res["ts"] = ts
            log.write(json.dumps(res, ensure_ascii=False) + "\n")
            errors += "error" in res
            if n % 1000 == 0:
                print(f"  [{n}/{len(changes)}] errores={errors}", flush=True)
    print(f"Actualizados: {len(changes) - errors} | errores: {errors}", flush=True)
    return 1 if errors else 0


if __name__ == "__main__":
    sys.exit(main())
