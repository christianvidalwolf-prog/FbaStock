import os
import json
import csv
import time
import shutil
import pandas as pd
from ftplib import FTP
import datetime
import subprocess

# --- CONFIGURATION ---
WORK_DIR = "/Users/christianvidalwolf/Stock"
CATALOG_FILE = f"{WORK_DIR}/catalog_robust.json"
EXCEL_PRICES_FILE = f"{WORK_DIR}/excel_prices_final.json"
OUTPUT_FILE = f"{WORK_DIR}/STOCK AMZ.txt"
PRECIOS_FILE = f"{WORK_DIR}/precios ES.xlsx"
STOCK_TREDISER_FILE = f"{WORK_DIR}/STOCK TREDISER.xls"

AMAZON_NEW_TEMPLATE_HEADERS = [
    ['settings=feedType=256&timestamp=2026-09-04T11%3A57%3A59.677Z&contributorId=amzn1.cr.o.A3RY0L9OY3TPHI&primaryMarketplaceId=amzn1.mp.o.A1RKKUPIHCS9HS&contentLanguageTag=es_ES&templateIdentifier=e4bbc050-a76e-4b77-acd9-e91ac5e8be26&headerLanguageTag=es_ES&labelRow=4&attributeRow=5&dataRow=7&flavor=seller-price-quantity&isProcessingSummary=false&isEdit=false&productTypeRequirement=LISTING_OFFER_ONLY&listingsItemRequirement=LISTING_OFFER_ONLY&reportProvenance=false&settingsHasAllDelocalizationData=true&upgradedFrom=priceandqty&ptds=UFJPRFVDVA%3D%3D&ptdToNamespaceMap=eyJQUk9EVUNUIjoiaW5nZXN0aW9uIn0%3D&browseClassifications=W3sicHJvZHVjdFR5cGUiOiJQUk9EVUNUIiwiYnJvd3NlQ2xhc3NpZmljYXRpb25LZXlzIjpbXX1d&vendorCodes=W10%3D&AttributeDefaultValues=eyJwcm9kdWN0X3R5cGUjMS52YWx1ZSI6IlBST0RVQ1QiLCJyZWNvcmRfYWN0aW9uIzEudmFsdWUiOiJwYXJ0aWFsX3VwZGF0ZSJ9&attributeSettings=W3siYXR0cmlidXRlIjoicHVyY2hhc2FibGVfb2ZmZXJbbWFya2V0cGxhY2VfaWQ9QTFSS0tVUElIQ1M5SFNdW2F1ZGllbmNlPUFMTF0jMS5vdXJfcHJpY2UjMS5zY2hlZHVsZSMxLnZhbHVlX3dpdGhfdGF4IiwiYWxpYXNlcyI6eyJFbGltaW5hciBvZmVydGEgKFZlbmRlciBlbiBBbWF6b24pIjoiYW16bjEudm9sdC5jdi5kZWxldGVfdW1wX3ZhcmlhbnQifX0seyJhdHRyaWJ1dGUiOiJwdXJjaGFzYWJsZV9vZmZlclttYXJrZXRwbGFjZV9pZD1BMVJLS1VQSUhDUzlIU11bYXVkaWVuY2U9QUxMXSMxLm1heGltdW1fc2VsbGVyX2FsbG93ZWRfcHJpY2UjMS5zY2hlZHVsZSMxLnZhbHVlX3dpdGhfdGF4IiwiYWxpYXNlcyI6eyJFbGltaW5hciBQcmVjaW8gbcOheGltbyBwZXJtaXRpZG8gYWwgdmVuZGVkb3IgKFZlbmRlciBlbiBBbWF6b24sIEVTKSI6ImFtem4xLnZvbHQuY3YuZGVsZXRlX3VtcF90b3BfbGV2ZWxfZmllbGQifX0seyJhdHRyaWJ1dGUiOiJtZXJjaGFudF9zaGlwcGluZ19ncm91cCMxLnZhbHVlIiwiYWxpYXNlcyI6eyJFTlZJTyA0IjoibGVnYWN5LXRlbXBsYXRlLWlkIiwiUGxhbnRpbGxhIG1pZ3JhZGEgcGFyYSBtw7pzaWNhIjoibGVnYWN5LW11c2ljLXRlbXBsYXRlLWlkIiwiR1JVUE8gMCI6ImY4MzczNDFlLTk3MjUtNDUyNC05YjhlLTk1YmM2OWFkMjUzZCIsIlByaW1lIFZlbmRlZG9yIjoiYWRmYTdhZmEtY2QyMS00NTNhLWJmNjUtMDFiYWY0MjBlYmJkIiwiUGxhbnRpbGxhIG1pZ3JhZGEgcGFyYSB2aWRlbyI6ImxlZ2FjeS12aWRlby10ZW1wbGF0ZS1pZCIsIkVudmlvcyAzIGRpYXMiOiIxZWZhODE5Zi1kOTEyLTRkZDAtODYwNy04NzUwY2IwZjZlNGEiLCJQbGFudGlsbGEgbWlncmFkYSBwYXJhIGxpYnJvcyI6ImxlZ2FjeS1ib29rLXRlbXBsYXRlLWlkIiwiUHJpbWUgTmFjaW9uYWwiOiI3MDM0MjFiYy03ODQzLTRjZTgtYjJiMC1kYmYxYWMxYzY5YjYiLCJQZW5pbnN1bGEzIjoiYWM4ZGY4ZDgtNzlkZC00ZGVkLThjMTYtNTMyZWRkZDEyOTYyIiwiUGxhbnRpbGxhIG1pZ3JhZGEgcGFyYSBEVkQiOiJsZWdhY3ktZHZkLXRlbXBsYXRlLWlkIn19LHsiYXR0cmlidXRlIjoicHVyY2hhc2FibGVfb2ZmZXJbbWFya2V0cGxhY2VfaWQ9QTFSS0tVUElIQ1M5SFNdW2F1ZGllbmNlPUFMTF0jMS5jdXJyZW5jeSIsImFsaWFzZXMiOnsiRVVSIjoiRVVSIn19LHsiYXR0cmlidXRlIjoicHVyY2hhc2FibGVfb2ZmZXJbbWFya2V0cGxhY2VfaWQ9QTFSS0tVUElIQ1M5SFNdW2F1ZGllbmNlPUFMTF0jMS5taW5pbXVtX3NlbGxlcl9hbGxvd2VkX3ByaWNlIzEuc2NoZWR1bGUjMS52YWx1ZV93aXRoX3RheCIsImFsaWFzZXMiOnsiRWxpbWluYXIgUHJlY2lvIG3DrW5pbW8gcGVybWl0aWRvIGFsIHZlbmRlZG9yIChWZW5kZXIgZW4gQW1hem9uLCBFUykiOiJhbXpuMS52b2x0LmN2LmRlbGV0ZV91bXBfdG9wX2xldmVsX2ZpZWxkIn19LHsiYXR0cmlidXRlIjoicHVyY2hhc2FibGVfb2ZmZXJbYXVkaWVuY2U9QjJCXSMxLnF1YW50aXR5X2Rpc2NvdW50X3BsYW4jMS5zY2hlZHVsZSMxLmRpc2NvdW50X3R5cGUiLCJhbGlhc2VzIjp7ImZpeGVkIjoiZml4ZWQiLCJwZXJjZW50IjoicGVyY2VudCJ9fSx7ImF0dHJpYnV0ZSI6InB1cmNoYXNhYmxlX29mZmVyW21hcmtldHBsYWNlX2lkPUExUktLVVBJSENTOUhTXVthdWRpZW5jZT1CMkJdIzEuY3VycmVuY3kiLCJhbGlhc2VzIjp7IkVVUiI6IkVVUiJ9fSx7ImF0dHJpYnV0ZSI6InB1cmNoYXNhYmxlX29mZmVyW21hcmtldHBsYWNlX2lkPUExUktLVVBJSENTOUhTXVthdWRpZW5jZT1CMkJdIzEucXVhbnRpdHlfZGlzY291bnRfcGxhbiMxLnNjaGVkdWxlIzEuZGlzY291bnRfdHlwZSIsImFsaWFzZXMiOnsiZml4ZWQiOiJmaXhlZCIsIkVsaW1pbmFyIGRlc2N1ZW50b3MgcG9yIGNhbnRpZGFkIjoiYW16bjEudm9sdC5jdi5kZWxldGVfdW1wX3RvcF9sZXZlbF9maWVsZCIsInBlcmNlbnQiOiJwZXJjZW50In19LHsiYXR0cmlidXRlIjoicHVyY2hhc2FibGVfb2ZmZXJbYXVkaWVuY2U9QUxMXSMxLmF1dG9tYXRlZF9wcmljaW5nX21lcmNoYW5kaXNpbmdfcnVsZV9wbGFuIzEubWVyY2hhbmRpc2luZ19ydWxlLnJ1bGVfaWQiLCJhbGlhc2VzIjp7IkdyYWRlIEFuZCBSZXNlbGwgQXV0b21hdGVkIFByaWNpbmcgUnVsZSI6Ijk2MjAxMzY1NDAwMi1HUkFERV9BTkRfUkVTRUxMX0FVVE9NQVRFRCIsIkNvbnRyb2wiOiJkMGYzNGY5OC0xNmMzLTRmZGMtYjU2OS05MTRiMmUzNDg2NTYiLCJCdXNpbmVzcyI6ImEzYjdmOTA4LTlhOTUtNGM3My1iOGZlLWUzOGI3YmZkM2M5OCIsIkNvbXBldGl0aXZlIFByaWNlIFJ1bGUgYnkgQW1hem9uIjoiMTg0MzY2MzcxMDAyLUNPTVBFVElUSVZFX0JVWUJPWCJ9fSx7ImF0dHJpYnV0ZSI6Im1lcmNoYW50X3NoaXBwaW5nX2dyb3VwW21hcmtldHBsYWNlX2lkPUExUktLVVBJSENTOUhTXSMxLnZhbHVlIiwiYWxpYXNlcyI6eyJFTlZJTyA0IjoibGVnYWN5LXRlbXBsYXRlLWlkIiwiUGxhbnRpbGxhIG1pZ3JhZGEgcGFyYSBtw7pzaWNhIjoibGVnYWN5LW11c2ljLXRlbXBsYXRlLWlkIiwiR1JVUE8gMCI6ImY4MzczNDFlLTk3MjUtNDUyNC05YjhlLTk1YmM2OWFkMjUzZCIsIlByaW1lIFZlbmRlZG9yIjoiYWRmYTdhZmEtY2QyMS00NTNhLWJmNjUtMDFiYWY0MjBlYmJkIiwiUGxhbnRpbGxhIG1pZ3JhZGEgcGFyYSB2aWRlbyI6ImxlZ2FjeS12aWRlby10ZW1wbGF0ZS1pZCIsIkVudmlvcyAzIGRpYXMiOiIxZWZhODE5Zi1kOTEyLTRkZDAtODYwNy04NzUwY2IwZjZlNGEiLCJQbGFudGlsbGEgbWlncmFkYSBwYXJhIGxpYnJvcyI6ImxlZ2FjeS1ib29rLXRlbXBsYXRlLWlkIiwiUHJpbWUgTmFjaW9uYWwiOiI3MDM0MjFiYy03ODQzLTRjZTgtYjJiMC1kYmYxYWMxYzY5YjYiLCJQZW5pbnN1bGEzIjoiYWM4ZGY4ZDgtNzlkZC00ZGVkLThjMTYtNTMyZWRkZDEyOTYyIiwiUGxhbnRpbGxhIG1pZ3JhZGEgcGFyYSBEVkQiOiJsZWdhY3ktZHZkLXRlbXBsYXRlLWlkIn19LHsiYXR0cmlidXRlIjoicHVyY2hhc2FibGVfb2ZmZXJbbWFya2V0cGxhY2VfaWQ9QTFSS0tVUElIQ1M5SFNdW2F1ZGllbmNlPUIyQl0jMS5taW5pbXVtX3NlbGxlcl9hbGxvd2VkX3ByaWNlIzEuc2NoZWR1bGUjMS52YWx1ZV93aXRoX3RheCIsImFsaWFzZXMiOnsiRWxpbWluYXIgUHJlY2lvIG3DrW5pbW8gcGVybWl0aWRvIGFsIHZlbmRlZG9yIChCMkIsIEVTKSI6ImFtem4xLnZvbHQuY3YuZGVsZXRlX3VtcF90b3BfbGV2ZWxfZmllbGQifX0seyJhdHRyaWJ1dGUiOiJwdXJjaGFzYWJsZV9vZmZlclttYXJrZXRwbGFjZV9pZD1BMVJLS1VQSUhDUzlIU11bYXVkaWVuY2U9QUxMXSMxLmF1dG9tYXRlZF9wcmljaW5nX21lcmNoYW5kaXNpbmdfcnVsZV9wbGFuIzEubWVyY2hhbmRpc2luZ19ydWxlLnJ1bGVfaWQiLCJhbGlhc2VzIjp7IkdyYWRlIEFuZCBSZXNlbGwgQXV0b21hdGVkIFByaWNpbmcgUnVsZSI6Ijk2MjAxMzY1NDAwMi1HUkFERV9BTkRfUkVTRUxMX0FVVE9NQVRFRCIsIkNvbnRyb2wiOiJkMGYzNGY5OC0xNmMzLTRmZGMtYjU2OS05MTRiMmUzNDg2NTYiLCJCdXNpbmVzcyI6ImEzYjdmOTA4LTlhOTUtNGM3My1iOGZlLWUzOGI3YmZkM2M5OCIsIkNvbXBldGl0aXZlIFByaWNlIFJ1bGUgYnkgQW1hem9uIjoiMTg0MzY2MzcxMDAyLUNPTVBFVElUSVZFX0JVWUJPWCJ9fSx7ImF0dHJpYnV0ZSI6InByb2R1Y3RfdHlwZSMxLnZhbHVlIiwiYWxpYXNlcyI6eyJQUk9EVUNUIjoiUFJPRFVDVCJ9fSx7ImF0dHJpYnV0ZSI6ImZ1bGZpbGxtZW50X2F2YWlsYWJpbGl0eSMxLmlzX2ludmVudG9yeV9hdmFpbGFibGUiLCJhbGlhc2VzIjp7IkhhYmlsaXRhZG8iOiJ0cnVlIiwiRGVzaGFiaWxpdGFkbyI6ImZhbHNlIn19LHsiYXR0cmlidXRlIjoicHVyY2hhc2FibGVfb2ZmZXJbbWFya2V0cGxhY2VfaWQ9QTFSS0tVUElIQ1M5SFNdW2F1ZGllbmNlPUIyQl0jMS5tYXhpbXVtX3NlbGxlcl9hbGxvd2VkX3ByaWNlIzEuc2NoZWR1bGUjMS52YWx1ZV93aXRoX3RheCIsImFsaWFzZXMiOnsiRWxpbWluYXIgUHJlY2lvIG3DoXhpbW8gcGVybWl0aWRvIGFsIHZlbmRlZG9yIChCMkIsIEVTKSI6ImFtem4xLnZvbHQuY3YuZGVsZXRlX3VtcF90b3BfbGV2ZWxfZmllbGQifX0seyJhdHRyaWJ1dGUiOiJwdXJjaGFzYWJsZV9vZmZlclthdWRpZW5jZT1CMkJdIzEuY3VycmVuY3kiLCJhbGlhc2VzIjp7IkVVUiI6IkVVUiJ9fSx7ImF0dHJpYnV0ZSI6InB1cmNoYXNhYmxlX29mZmVyW21hcmtldHBsYWNlX2lkPUExUktLVVBJSENTOUhTXVthdWRpZW5jZT1CMkJdIzEub3VyX3ByaWNlIzEuc2NoZWR1bGUjMS52YWx1ZV93aXRoX3RheCIsImFsaWFzZXMiOnsiRWxpbWluYXIgb2ZlcnRhIChCMkIpIjoiYW16bjEudm9sdC5jdi5kZWxldGVfdW1wX3ZhcmlhbnQifX0seyJhdHRyaWJ1dGUiOiJmdWxmaWxsbWVudF9hdmFpbGFiaWxpdHkjMS5mdWxmaWxsbWVudF9jaGFubmVsX2NvZGUiLCJhbGlhc2VzIjp7IkFNQVpPTl9FVV9SQUZOIjoiQU1BWk9OX0VVX1JBRk4iLCJBTUFaT05fRVUyVFJfUkFGTiI6IkFNQVpPTl9FVTJUUl9SQUZOIiwiTG9nw61zdGljYSBwb3IgcGFydGUgZGVsIHZlbmRlZG9yIChwcmVkZXRlcm1pbmFkbykiOiJERUZBVUxUIiwiQU1BWk9OX1VLX1JBRk4iOiJBTUFaT05fVUtfUkFGTiIsIkFNQVpPTl9VSzJOTF9SQUZOIjoiQU1BWk9OX1VLMk5MX1JBRk4iLCJBTUFaT05fRVVfVkNTIjoiQU1BWk9OX0VVX1ZDUyIsIkxvZ8Otc3RpY2EgZGUgQW1hem9uIChVRSkiOiJBTUFaT05fRVUiLCJBTUFaT05fRVUyU0FfUkFGTiI6IkFNQVpPTl9FVTJTQV9SQUZOIn19LHsiYXR0cmlidXRlIjoicHVyY2hhc2FibGVfb2ZmZXJbYXVkaWVuY2U9QUxMXSMxLmN1cnJlbmN5IiwiYWxpYXNlcyI6eyJFVVIiOiJFVVIifX1d&TemplateType=unified&Version=2026.0904&TemplateSignature=UFJPRFVDVA==&umpVersion=MS41NC40NQ==', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', ''],
    ['     Debes rellenar esta plantilla en ESPAÑOL. NO modifiques ni elimines las filas de encabezado coloreadas. Para ampliar todas las columnas opcionales, haz clic en el botón "2" en la parte superior izquierda.', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', ''],
    ['Identidad de listing', 'Oferta (ES) - (Vender en Amazon), (ES) - (B2B)', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', ''],
    ['SKU', 'Cumplimiento de código de canal (ES)', 'Cantidad (ES)', 'Tiempo de tramitación (ES)', 'Fecha de reposición (ES)', 'Inventario siempre disponible (ES)', 'Tu precio EUR (Vender en Amazon, ES)', 'Regla de precios (Vender en Amazon, ES)', 'Precio mínimo permitido al vendedor (Vender en Amazon, ES)', 'Precio máximo permitido al vendedor (Vender en Amazon, ES)', 'Fecha de finalización de la venta. (Vender en Amazon, ES)', 'Fecha de comienzo de la venta. (Vender en Amazon, ES)', 'Precio de venta. EUR (Vender en Amazon, ES)', 'Fecha de inicio de la oferta (Vender en Amazon, ES)', 'Fecha de finalización de venta (Vender en Amazon, ES)', 'Tu precio EUR (B2B, ES)', 'Precio mínimo permitido al vendedor (B2B, ES)', 'Precio máximo permitido al vendedor (B2B, ES)', 'Fecha de inicio de la oferta (B2B, ES)', 'Fecha de finalización de venta (B2B, ES)', 'Tipo de precio por cantidad (B2B, ES)', 'Umbral de cantidad (límite inferior, B2B, ES)', 'Precio por cantidad (precio fijo/descuento porcentual, B2B, ES)', 'Umbral de cantidad (límite inferior, B2B, ES)', 'Precio por cantidad (precio fijo/descuento porcentual, B2B, ES)', 'Umbral de cantidad (límite inferior, B2B, ES)', 'Precio por cantidad (precio fijo/descuento porcentual, B2B, ES)', 'Umbral de cantidad (límite inferior, B2B, ES)', 'Precio por cantidad (precio fijo/descuento porcentual, B2B, ES)', 'Umbral de cantidad (límite inferior, B2B, ES)', 'Precio por cantidad (precio fijo/descuento porcentual, B2B, ES)', 'Plantilla de envío (ES)'],
    ['contribution_sku#1.value', 'fulfillment_availability#1.fulfillment_channel_code', 'fulfillment_availability#1.quantity', 'fulfillment_availability#1.lead_time_to_ship_max_days', 'fulfillment_availability#1.restock_date', 'fulfillment_availability#1.is_inventory_available', 'purchasable_offer[marketplace_id=A1RKKUPIHCS9HS][audience=ALL]#1.our_price#1.schedule#1.value_with_tax', 'purchasable_offer[marketplace_id=A1RKKUPIHCS9HS][audience=ALL]#1.automated_pricing_merchandising_rule_plan#1.merchandising_rule.rule_id', 'purchasable_offer[marketplace_id=A1RKKUPIHCS9HS][audience=ALL]#1.minimum_seller_allowed_price#1.schedule#1.value_with_tax', 'purchasable_offer[marketplace_id=A1RKKUPIHCS9HS][audience=ALL]#1.maximum_seller_allowed_price#1.schedule#1.value_with_tax', 'purchasable_offer[marketplace_id=A1RKKUPIHCS9HS][audience=ALL]#1.discounted_price#1.schedule#1.end_at', 'purchasable_offer[marketplace_id=A1RKKUPIHCS9HS][audience=ALL]#1.discounted_price#1.schedule#1.start_at', 'purchasable_offer[marketplace_id=A1RKKUPIHCS9HS][audience=ALL]#1.discounted_price#1.schedule#1.value_with_tax', 'purchasable_offer[marketplace_id=A1RKKUPIHCS9HS][audience=ALL]#1.start_at.value', 'purchasable_offer[marketplace_id=A1RKKUPIHCS9HS][audience=ALL]#1.end_at.value', 'purchasable_offer[marketplace_id=A1RKKUPIHCS9HS][audience=B2B]#1.our_price#1.schedule#1.value_with_tax', 'purchasable_offer[marketplace_id=A1RKKUPIHCS9HS][audience=B2B]#1.minimum_seller_allowed_price#1.schedule#1.value_with_tax', 'purchasable_offer[marketplace_id=A1RKKUPIHCS9HS][audience=B2B]#1.maximum_seller_allowed_price#1.schedule#1.value_with_tax', 'purchasable_offer[marketplace_id=A1RKKUPIHCS9HS][audience=B2B]#1.start_at.value', 'purchasable_offer[marketplace_id=A1RKKUPIHCS9HS][audience=B2B]#1.end_at.value', 'purchasable_offer[marketplace_id=A1RKKUPIHCS9HS][audience=B2B]#1.quantity_discount_plan#1.schedule#1.discount_type', 'purchasable_offer[marketplace_id=A1RKKUPIHCS9HS][audience=B2B]#1.quantity_discount_plan#1.schedule#1.levels#1.lower_bound', 'purchasable_offer[marketplace_id=A1RKKUPIHCS9HS][audience=B2B]#1.quantity_discount_plan#1.schedule#1.levels#1.value', 'purchasable_offer[marketplace_id=A1RKKUPIHCS9HS][audience=B2B]#1.quantity_discount_plan#1.schedule#1.levels#2.lower_bound', 'purchasable_offer[marketplace_id=A1RKKUPIHCS9HS][audience=B2B]#1.quantity_discount_plan#1.schedule#1.levels#2.value', 'purchasable_offer[marketplace_id=A1RKKUPIHCS9HS][audience=B2B]#1.quantity_discount_plan#1.schedule#1.levels#3.lower_bound', 'purchasable_offer[marketplace_id=A1RKKUPIHCS9HS][audience=B2B]#1.quantity_discount_plan#1.schedule#1.levels#3.value', 'purchasable_offer[marketplace_id=A1RKKUPIHCS9HS][audience=B2B]#1.quantity_discount_plan#1.schedule#1.levels#4.lower_bound', 'purchasable_offer[marketplace_id=A1RKKUPIHCS9HS][audience=B2B]#1.quantity_discount_plan#1.schedule#1.levels#4.value', 'purchasable_offer[marketplace_id=A1RKKUPIHCS9HS][audience=B2B]#1.quantity_discount_plan#1.schedule#1.levels#5.lower_bound', 'purchasable_offer[marketplace_id=A1RKKUPIHCS9HS][audience=B2B]#1.quantity_discount_plan#1.schedule#1.levels#5.value', 'merchant_shipping_group[marketplace_id=A1RKKUPIHCS9HS]#1.value'],
    ['ABC123', 'AMAZON_NA', '152', '5', '2020-05-05', 'Deshabilitado', '50.00', 'Regla de precios competitivos de Amazon', '259.99', '259.99', '2017-07-01', '2017-06-30', '219.99', '2017-06-30', '2017-07-01', '50.00', '259.99', '259.99', '2017-06-30', '2017-07-01', 'Fijo, porcentaje', '5, 10', '10, 20', '5, 10', '10, 20', '5, 10', '10, 20', '5, 10', '10, 20', '5, 10', '10, 20', 'Productos pesados y voluminosos, entrega de electrodomésticos grandes de NCR'],
]

SKUS_FORZAR_CERO = {
    "4085VC",
    "4085VCI",
    "2598VCI",
    "2598VC",
    "5324VC",
    "4100VCO",
    "4122VC0",
    "3635VC0",
    "2523VC",
    "4686VC",
    "2450VC",
    "2450VCI",  # ASIN B010TN6SXU
    "11631VC",
    "1237VC",
    "1238VC",
    "1652VC",
    "1653VC",
    "1684VC",
    "1688VC",
    "1717VC",
    "1718VC",
    "180VC",
    "1832VC",
    "2025VC",
    "2027VC",
    "2180VC",
    "2181VC",
    "2210VC",
    "2260VC",
    "2355VC",
    "2356VC",
    "2360VC",
    "2400VC",
    "2401VC",
    "2402VC",
    "2405VC",
    "2406VC",
    "2414VC",
    "2415VC",
    "2425VC",
    "2447VC",
    "2510VC",
    "2517VC",
    "2522VC",
    "2611VC",
    "2648VC",
    "2660VC",
    "2665VC",
    "2857VC",
    "2858VC",
    "2990VC",
    "3011VC",
    "3012VC",
    "3013VC",
    "3015VC",
    "3018VC",
    "3019VC",
    "3020VC",
    "3021VC",
    "3026VC",
    "3027VC",
    "3030VC",
    "3035VC",
    "3038VC",
    "3041VC",
    "3043VC",
    "3044VC",
    "3045VC",
    "3046VC",
    "3048VC",
    "3052VC",
    "3053VC",
    "3054VC",
    "3055VC",
    "3062VC",
    "3063VC",
    "3064VC",
    "3069VC",
    "3070VC",
    "3071VC",
    "3072VC",
    "3074VC",
    "3089VC",
    "3116VC",
    "3117VC",
    "3118VC",
    "3119VC",
    "3120VC",
    "3121VC",
    "3123VC",
    "3125VC",
    "3126VC",
    "3127VC",
    "3130VC",
    "3133VC",
    "3134VC",
    "3135VC",
    "3200VC",
    "3201VC",
    "3230VC",
    "3311VC",
    "3314VC",
    "3315VC",
    "3354VC",
    "3372VC",
    "3380VC",
    "3391VC",
    "3392VC",
    "3423VC",
    "3432VC",
    "3437VC",
    "3484VC",
    "3487VC",
    "3555VC",
    "3557VC",
    "3558VC",
    "3559VC",
    "3593VC",
    "3593VC0",
    "3712VC",
    "3790VC",
    "3824VC",
    "4107VC",
    "4108VC",
    "4110VC",
    "4115VC",
    "4125VC",
    "4130VC",
    "4134VC",
    "4140VC",
    "4155VC",
    "4158VC",
    "4165VC",
    "4231VC",
    "4232VC",
    "4255VC",
    "4259VC",
    "4261VC",
    "4263VC",
    "4264VC",
    "4265VC",
    "4285VC",
    "4285VC0",
    "4288VC",
    "4330VC",
    "4361VC",
    "4442VC",
    "4445VC",
    "4462VC",
    "4528VC",
    "4531VC",
    "4534VC",
    "4558VC",
    "4657VC",
    "4721VC",
    "4722VC",
    "4727VC",
    "4735VC",
    "4737VC",
    "4745VC",
    "4752VC",
    "4758VC",
    "4765VC",
    "4785VC",
    "4790VC",
    "4801VC",
    "4815VC",
    "4821VC",
    "4874VC",
    "4882VC",
    "5032VC",
    "5088VC",
    "5090VC",
    "5096VC",
    "5113VC",
    "5129VC",
    "5382VC",
    "5383VC",
    "5440VC",
    "5442VC",
    "5456VC",
    "5540VC",
    "5571VC",
    "5572VC",
    "5574VC",
    "5575VC",
    "5588VC",
    "5592VC",
    "5685VC",
    "5712VC",
    "5721VC",
    "5807VC",
    "5823VC",
    "5825VC",
    "5835VC",
    "5852VC",
    "5871VC",
    "5872VC",
    "5873VC",
    "5874VC",
    "5882VC",
    "5883VC",
    "5885VC",
    "5987VC",
    "5991VC",
    "5995VC",
    "6001VC",
    "6007VC",
    "6040VC",
    "6151VC",
    "6252VC",
    "6255VC",
    "6270VC",
    "6471VC",
    "5658VCI",
    "2915244CLM",
    "30512912CLM",
    "2861262CLM",
    "41683MDRG",
    "41891MDRG",
    "11329VCT",
    "11329VC",
}

SKUS_ELIMINADOS = {
    "277222CLM",
    "22473SG",
    "193SG",
}

PRECIOS_FIJOS = {
    "2515VCI": 59.0,
    "4746VC": 22.0,
    "31612SG": 25.0,
    "3097101CLM": 15.95,
    "3097100CLM": 15.95,
    "12359VC": 14.95,
    "1016VCI": 10.99,
    "14165VCI": 8.99,
    "40552MD": 22.99,
    "42772MD": 26.99,
    "4359VCI": 10.99,
    "309766CLM": 9.99,
    "31469MD": 13.99,
    "2388VCI": 15.99,
    "2185651CLM": 9.99,
    "23779SG": 15.99,
    "17385VCI": 8.99,
    "31181MD": 21.99,
    "1304VC": 10.99,
    "15846VCI": 12.99,
    "11302VC": 16.99,
    "43047SGRG": 39.99,
    "4148VCI": 39.0,
}

# SKUs excluidos de la automatización (se actualizan manualmente)
SKUS_MANUALES = {
    "41170MDRG",
    "41171MDRG",
    "40907MDRG",
    "41244MDRG",
    "41116MDRG",
    "41113MDRG",
    "41114MDRG",
    "41932MDRG",
    "41936MDRG",
    "41931MDRG",
    "41930MDRG",
    "41929MDRG",
    "30751MDRG",
    "30752MDRG",
    "41058MDRG",
    "41059MDRG",
    "41060MDRG",
    "30517MDRG",
    "30048MDRG",
    "30047MDRG",
    "30050MDRG",
    "30052MDRG",
    "30053MDRG",
    "30055MDRG",
    "30057MDRG",
    "30058MDRG",
    "30060MDRG",
    "41603MDRG",
    "41591MDRG",
    "41594MDRG",
    "41592MDRG",
    "41733MDRG",
    "41726MDRG",
    "41732MDRG",
    "41887MDRG",
    "41714MDRG",
    "30934MDRG",
    "41490MDRG",
    "41489MDRG",
    "41713MDRG",
    "41486MDRG",
    "41717MDRG",
    "30933MDRG",
    "31008MDRG",
    "40064MDRG",
    "41396MDRG",
    "40918MDRG",
    "40921MDRG",
    "41925MDRG",
    "41927MDRG",
    "41815MDRG",
    "41858MDRG",
    "41517MDRG",
    "41041MDRG",
    "41050MDRG",
    "41805MDRG",
    "41804MDRG",
    "41518MDRG",
    "41048MDRG",
    "41039MDRG",
    "41807MDRG",
    "31025MDRG",
    "41966MDRG",
    "41349MDRG",
    "41647MDRG",
    "41348MDRG",
    "41505MDRG",
    "41503MDRG",
    "41623MDRG",
    "41622MDRG",
    "41665MDRG",
    "30127MDRG",
    "30125MDRG",
    "41940MDRG",
    "41939MDRG",
    "41938MDRG",
    "41937MDRG",
    "41944MDRG",
    "41943MDRG",
    "41942MDRG",
    "41941MDRG",
    "30886MDRG",
    "41201MDRG",
    "41230MDRG",
    "30982MDRG",
    "40053MDRG",
    "41964MDRG",
    "41963MDRG",
    "30939MDRG",
    "30139MDRG",
    "30940MDRG",
    "40953MDRG",
    "41669MDRG",
    "41650MDRG",
    "41193MDRG",
    "40822MDRG",
    "41835MDRG",
    "41735MDRG",
    "41653MDRG",
    "41654MDRG",
    "41331MDRG",
    "41324MDRG",
    "40057MDRG",
    "41105MDRG",
    "41338MDRG",
    "41084MDRG",
    "41083MDRG",
    "41079MDRG",
    "41646MDRG",
    "41673MDRG",
    "41377MDRG",
    "40027MDRG",
    "41432MDRG",
    "41548MDRG",
    "41973MDRG",
    "41808MDRG",
    "41411MDRG",
    "20004MDRG",
    "20005MDRG",
    "20002MDRG",
    "20030MDRG",
    "20880MDRG",
    "20626MDRG",
    "20612MDRG",
    "20710MDRG",
    "20725MDRG",
    "20726MDRG",
    "20006MDRG",
    "20008MDRG",
    "20606MDRG",
    "20733MDRG",
    "20929MDRG",
    "20900MDRG",
    "20908MDRG",
    "20895MDRG",
    "20892MDRG",
    "20890MDRG",
    "30937MDRG",
    "30938MDRG",
    "40934MDRG",
    "40933MDRG",
    "41096MDRG",
    "41094MDRG",
    "41874MDRG",
    "41875MDRG",
    "41683MDRG",
    "41889MDRG",
    "41891MDRG",
    "41508MDRG",
    "41221MDRG",
    "41854MDRG",
    "41852MDRG",
    "41498MDRG",
    "41495MDRG",
    "41494MDRG",
    "41965MDRG",
    "41663MDRG",
    "41659MDRG",
    "41660MDRG",
    "41658MDRG",
    "41657MDRG",
    "30039MDRG",
    "30038MDRG",
    "40043MDRG",
    "41426MDRG",
    "2915244CLM",
    "30512912CLM",
    "2861262CLM",
    "41683MDRG",
    "41891MDRG",
    "41429MDRG",
    "40040MDRG",
    "40038MDRG",
    "40039MDRG",
    "41423MDRG",
    "41424MDRG",
    "41425MDRG",
    "41430MDRG",
    "41007MDRG",
    "41006MDRG",
    "41009MDRG",
    "41014MDRG",
    "41015MDRG",
    "41008MDRG",
    "41000MDRG",
    "40997MDRG",
    "30919MDRG",
    "31086MDRG",
    "31087MDRG",
    "40294MDRG",
    "41661MDRG",
    "30102MDRG",
    "30111MDRG",
    "30103MDRG",
    "30104MDRG",
    "30105MDRG",
    "30108MDRG",
    "30109MDRG",
    "30114MDRG",
    "41197MDRG",
    "41198MDRG",
    "41085MDRG",
    "41617MDRG",
    "41613MDRG",
    "41614MDRG",
    "41618MDRG",
    "41615MDRG",
    "41678MDRG",
    "41679MDRG",
    "41401MDRG",
    "41674MDRG",
    "41400MDRG",
    "40956MDRG",
    "41154MDRG",
    "40254MD",
    "40255MD",
    "20325MD",
    "30020MD",
    "30021MD",
    "30022MD",
    "30023MD",
    "30270MD",
    "30324MD",
    "40205MD",
    "30325MD",
    "30169MD",
    "30168MD",
    "30172MD",
    "30171MD",
    "30056MD",
    "30059MD",
    "30166MD",
    "30167MD",
    "30051MD",
    "30007MD",
    "30003MD",
    "30008MD",
    "30004MD",
    "30002MD",
    "30000MD",
    "40034MD",
    "30037MD",
    "30025MD",
    "30026MD",
    "30163MD",
    "30164MD",
    "30173MD",
    "30036MD",
    "30035MD",
    "40067MD",
    "20062MD",
    "40044MD",
    "40073MD",
    "40210MD",
    "40076MD",
    "20156MD",
    "20155MD",
    "40080MD",
    "30024MD",
    "20152MD",
    "40093MD",
    "30054MD",
    "20323MD",
    "30260MD",
    "40293MD",
    "30253MD",
    "20153MD",
    "30062MD",
    "40106MD",
    "40105MD",
    "30274MD",
    "40071MD",
    "40243MD",
    "40258MD",
    "40035MD",
    "40229MD",
    "40239MD",
    "20322MD",
    "40082MD",
    "30300MD",
    "30189MD",
    "30267MD",
    "40228MD",
    "30126MD",
    "30271MD",
    "40256MD",
    "20154MD",
    "40260MD",
    "30268MD",
    "20321MD",
    "40081MD",
    "20151MD",
    "30273MD",
    "30084MD",
    "40074MD",
    "40185MD",
    "40318MD",
    "30370MD",
    "40399MD",
    "30366MD",
    "30369MD",
    "30367MD",
    "40320MD",
    "30357MD",
    "30397MD",
    "40188MD",
    "30362MD",
    "30377MD",
    "40264MD",
    "30358MD",
    "30371MD",
    "30363MD",
    "30414MD",
    "40697MD",
    "40535MD",
    "30540MD",
    "40511MD",
    "30537MD",
    "40457MD",
    "30628MD",
    "40560MD",
    "30629MD",
    "30625MD",
    "40512MD",
    "40667MD",
    "30544MD",
    "40504MD",
    "40533MD",
    "40661MD",
    "30547MD",
    "30624MD",
    "40655MD",
    "30631MD",
    "40506MD",
    "40657MD",
    "40278MD",
    "40686MD",
    "30626MD",
    "40696MD",
    "40539MD",
    "40660MD",
    "40280MD",
    "40550MD",
    "40510MD",
    "40668MD",
    "40509MD",
    "40513MD",
    "40393MD",
    "40698MD",
    "40419MD",
    "40813MD",
    "40817MD",
    "40838MD",
    "40440MD",
    "40720MD",
    "40802MD",
    "40803MD",
    "40816MD",
    "30755MD",
    "40703MD",
    "30754MD",
    "40683MD",
    "40788MD",
    "40805MD",
    "40704MD",
    "40319MD",
    "41181MDRG",
    "41648MDRG",
    "41205MDRG",
    "41967MDRG",
    "41355MDRG",
    "40914MDRG",
    "42378MD",
    "42943MD",
    "42594MD",
    "42382MD",
    "42380MD",
    "42381MD",
    "42593MD",
    "41950MD",
    "41949MD",
    "42640MD",
    "42945MD",
    "42642MD",
    "42944MD",
    "41243MD",
    "40521MD",
    "40565MD",
    "40079MD",
    "40077MD",
    "40111MD",
    "40734MD",
    "41112MD",
    "41846MD",
    "41371MD",
    "40835MD",
    "41577MD",
    "41184MD",
    "42698MD",
    "31191MD",
    "42809MD",
    "40973MD",
    "42633MD",
    "42636MD",
    "42686MD",
    "41935MD",
    "42326MD",
    "40395MD",
    "42699MD",
    "42328MD",
    "42685MD",
    "31470MD",
    "42618MD",
    "42958MD",
    "42959MD",
    "42794MD",
    "42957MD",
    "42955MD",
    "42956MD",
    "42401MD",
    "31331MD",
    "30049MD",
    "31333MD",
    "31332MD",
    "31334MD",
    "31298MD",
    "31299MD",
    "31366MD",
    "31365MD",
    "31363MD",
    "31364MD",
    "31362MD",
    "31186MD",
    "42099MD",
    "41593MD",
    "42591MD",
    "42641MD",
    "41605MD",
    "31263MD",
    "31112MD",
    "30936MD",
    "30935MD",
    "30124MD",
    "31315MD",
    "30925MD",
    "31179MD",
    "42941MD",
    "42940MD",
    "42942MD",
    "41989MD",
    "41827MD",
    "42181MD",
    "42564MD",
    "42563MD",
    "41826MD",
    "42180MD",
    "42289MD",
    "42061MD",
    "41190MD",
    "42688MD",
    "42689MD",
    "42690MD",
    "40388MD",
    "42939MD",
    "42693MD",
    "42331MD",
    "42692MD",
    "42691MD",
    "42700MD",
    "42926MD",
    "40540MD",
    "42929MD",
    "42927MD",
    "42694MD",
    "42324MD",
    "40798MD",
    "42930MD",
    "42697MD",
    "42886MD",
    "42842MD",
    "42837MD",
    "42607MD",
    "42852MD",
    "42844MD",
    "42609MD",
    "42614MD",
    "42481MD",
    "42839MD",
    "42846MD",
    "42841MD",
    "42843MD",
    "42479MD",
    "42853MD",
    "42610MD",
    "42480MD",
    "42611MD",
    "42616MD",
    "42848MD",
    "42847MD",
    "42486MD",
    "42850MD",
    "42838MD",
    "42388MD",
    "42389MD",
    "42391MD",
    "30001MD",
    "30005MD",
    "31104MD",
    "31103MD",
    "31225MD",
    "31447MD",
    "31448MD",
    "31317MD",
    "31295MD",
    "31316MD",
    "31294MD",
    "42445MD",
    "31322MD",
    "31318MD",
    "42812MD",
    "42745MD",
    "42813MD",
    "42514MD",
    "42510MD",
    "41668MD",
    "41670MD",
    "41672MD",
    "42977MD",
    "42979MD",
    "42978MD",
    "42523MD",
    "40069MD",
    "42919MD",
    "41178MD",
    "40700MD",
    "31130MD",
    "31122MD",
    "31267MD",
    "42824MD",
    "42254MD",
    "42695MD",
    "42917MD",
    "42918MD",
    "42140MD",
    "42142MD",
    "40722MD",
    "40721MD",
    "42116MD",
    "42117MD",
    "42833MD",
    "42505MD",
    "41200MD",
    "41199MD",
    "42504MD",
    "42517MD",
    "42515MD",
    "42516MD",
    "42518MD",
    "42834MD",
    "41253MD",
    "41254MD",
    "41288MD",
    "42575MD",
    "42637MD",
    "42574MD",
    "42310MD",
    "42306MD",
    "42219MD",
    "42217MD",
    "42218MD",
    "42346MD",
    "42589MD",
    "42228MD",
    "42588MD",
    "42530MD",
    "42344MD",
    "42345MD",
    "42343MD",
    "41762MD",
    "41767MD",
    "41760MD",
    "41765MD",
    "41769MD",
    "41770MD",
    "42595MD",
    "42804MD",
    "42810MD",
    "42905MD",
    "42906MD",
    "42665MD",
    "42904MD",
    "42308MD",
    "42150MD",
    "41406MD",
    "42318MD",
    "42316MD",
    "42317MD",
    "42315MD",
    "42747MD",
    "42746MD",
    "42658MD",
    "42898MD",
    "42899MD",
    "42900MD",
    "42901MD",
    "42902MD",
    "42903MD",
    "42892MD",
    "42657MD",
    "42895MD",
    "42893MD",
    "42894MD",
    "41703MD",
    "41358MD",
    "42652MD",
    "42887MD",
    "42888MD",
    "40784MD",
    "42661MD",
    "42662MD",
    "42896MD",
    "42897MD",
    "42279MD",
    "42889MD",
    "42890MD",
    "42891MD",
    "42653MD",
    "42654MD",
    "41704MD",
    "42314MD",
    "42319MD",
    "42749MD",
    "42748MD",
    "42750MD",
    "41353MD",
    "41347MD",
    "41354MD",
    "41195MD",
    "41179MD",
    "41344MD",
    "42194MD",
    "41849MD",
    "41185MD",
    "42298MD",
    "42297MD",
    "42242MD",
    "31353MD",
    "41734MD",
    "42341MD",
    "42098MD",
    "42248MD",
    "42340MD",
    "42752MD",
    "41323MD",
    "42755MD",
    "42751MD",
    "42753MD",
    "42426MD",
    "41026MD",
    "41028MD",
    "42555MD",
    "41337MD",
    "42411MD",
    "41643MD",
    "42417MD",
    "42419MD",
    "42418MD",
    "42416MD",
    "40447MD",
    "42414MD",
    "42415MD",
    "42114MD",
    "42413MD",
    "42422MD",
    "41642MD",
    "41641MD",
    "42420MD",
    "42112MD",
    "42113MD",
    "42910MD",
    "42911MD",
    "42679MD",
    "42678MD",
    "42118MD",
    "42597MD",
    "42292MD",
    "42203MD",
    "42482MD",
    "42953MD",
    "42447MD",
    "41057MD",
    "42840MD",
    "42529MD",
    "41840MD",
    "40812MD",
    "42363MD",
    "42252MD",
    "42253MD",
    "42250MD",
    "21057MD",
    "21003MD",
    "21014MD",
    "21054MD",
    "21024MD",
    "21092MD",
    "21052MD",
    "21053MD",
    "20003MD",
    "21080MD",
    "21081MD",
    "31184MD",
    "31183MD",
    "30406MD",
    "30019MD",
    "30889MD",
    "30888MD",
    "31337MD",
    "31338MD",
    "31453MD",
    "31384MD",
    "42801MD",
    "42823MD",
    "42309MD",
    "41482MD",
    "31387MD",
    "31385MD",
    "31388MD",
    "31386MD",
    "42202MD",
    "42201MD",
    "42541MD",
    "42370MD",
    "42476MD",
    "42475MD",
    "42868MD",
    "42683MD",
    "42682MD",
    "42520MD",
    "42519MD",
    "42742MD",
    "42724MD",
    "42728MD",
    "42743MD",
    "42146MD",
    "42740MD",
    "42723MD",
    "42739MD",
    "42738MD",
    "42722MD",
    "42737MD",
    "42307MD",
    "42832MD",
    "42831MD",
    "42830MD",
    "42829MD",
    "42827MD",
    "42828MD",
    "42934MD",
    "21082MD",
    "40710MD",
    "42806MD",
    "40465MD",
    "42643MD",
    "41780MD",
    "42527MD",
    "42036MD",
    "42035MD",
    "42358MD",
    "42357MD",
    "41262MD",
    "42208MD",
    "41579MD",
    "42506MD",
    "42865MD",
    "42671MD",
    "42851MD",
    "42933MD",
    "41675MD",
    "42222MD",
    "42771MD",
    "42915MD",
    "42135MD",
    "42136MD",
    "42916MD",
    "42802MD",
    "42312MD",
    "42313MD",
    "40391MD",
    "40965MD",
    "41020MD",
    "41077MD",
    "41109MD",
    "41110MD",
    "41182MD",
    "41183MD",
    "41206MD",
    "41255MD",
    "41504MD",
    "41506MD",
    "41671MD",
    "41763MD",
    "41843MD",
    "42119MD",
    "42120MD",
    "42141MD",
    "42196MD",
    "42200MD",
    "42207MD",
    "42209MD",
    "42214MD",
    "42215MD",
    "42230MD",
    "42232MD",
    "42249MD",
    "42290MD",
    "42342MD",
    "42366MD",
    "42472MD",
    "42478MD",
    "42522MD",
    "42540MD",
    "42568MD",
    "42571MD",
    "42576MD",
    "42645MD",
    "42677MD",
    "42765MD",
    "42766MD",
    "42773MD",
    "42803MD",
    "42820MD",
    "42825MD",
    "42872MD",
    "42874MD",
    "42875MD",
    "42885MD",
    "42907MD",
    "42908MD",
    "42909MD",
    "42912MD",
    "42913MD",
    "42920MD",
    "42921MD",
    "42922MD",
    "42925MD",
    "42931MD",
    "42947MD",
    "42948MD",
    "42949MD",
    "83637VD-AM",
    "83627SGR",
    "83637SG",
    "83637SGRGAZ-VDE",
    "83637SGRJ-LI",
    "83637SGLI-AZ",
    "83637SGFBA",
    "83637SGAZ-VD",
    "83637SGAZ-RJ",
    "83637RGAZ-RJ",
    "83627SGVD-AM",
    "83627SGRO-AM",
    "83627SGAZ-VE",
    "83627SGAZ-VD",
    "83627SGAZ-RJ",
    "83627SG",
    "83627AZ-BLSG",
    "83627AZ-AMSG",
    "42576MD",
    "42571MD",
    "42568MD",
    "307105CLM",
    "307104CLM",
}

VC_PRICE_THRESHOLD = 30.0


def apply_vc_price_increase(price):
    if price <= 0 or price >= VC_PRICE_THRESHOLD:
        return price
    new_price = price * 1.05
    whole = int(new_price)
    decimal = new_price - whole
    if decimal <= 0.50:
        rounded = whole + 0.95
    else:
        rounded = whole + 1.95
    return round(rounded, 2)


def to_num(val):
    if val is None or val == "":
        return 0
    try:
        if isinstance(val, (int, float)):
            return float(val)
        val_str = str(val).replace(",", ".")
        return float(val_str)
    except:
        return 0


def load_trediser_stocks(path):
    try:
        df_t = pd.read_excel(path)
        if "Código" in df_t.columns:
            df_t = df_t.dropna(subset=["Código"])
            stocks = {}
            stock_column = "Unnamed: 3" if "Unnamed: 3" in df_t.columns else df_t.columns[-1]
            for _, row in df_t.iterrows():
                code = str(row["Código"]).strip().replace(".0", "")
                if not code:
                    continue
                stocks[code] = to_num(row[stock_column])
            return stocks

        raw_df = pd.read_excel(path, header=None)
        code_col = None
        stock_col = None
        data_start = 0

        for idx, row in raw_df.iterrows():
            for col_idx, value in enumerate(row.tolist()):
                value_str = str(value).strip() if pd.notna(value) else ""
                if value_str == "Código" and code_col is None:
                    code_col = col_idx
                    data_start = max(data_start, idx + 1)
                elif value_str == "Unidades" and stock_col is None:
                    stock_col = col_idx
                    data_start = max(data_start, idx + 1)

        if code_col is None or stock_col is None:
            raise KeyError("Could not locate 'Código'/'Unidades' columns in Trediser file")

        stocks = {}
        for _, row in raw_df.iloc[data_start:].iterrows():
            code_val = row.iloc[code_col]
            if pd.isna(code_val):
                continue
            code = str(code_val).strip().replace(".0", "")
            if not code or code.lower() == "código":
                continue
            stocks[code] = to_num(row.iloc[stock_col])

        return stocks
    except Exception:
        raise


def download_files():
    providers = {
        "dcasa": {
            "type": "ftp",
            "host": "data.dcasacollection.com",
            "user": "sek4283",
            "pass": "Rx34z5m6_pER",
        },
        "mina": {
            "type": "url",
            "url": "https://vivescortadaimport.com/modules/doofinder/feed2.php?language=ES&currency=EUR",
        },
        "signes": {
            "type": "url",
            "url": "https://signesconexion.com/stock/STOCK-44880.CSV",
        },
    }
    paths = {}

    # DCASA
    MIN_DCASA_ROWS = 1000

    # Check if we already have today's file (downloaded by curl)
    today_str = datetime.datetime.now().strftime("%Y%m%d")
    local = sorted(
        [
            f
            for f in os.listdir(WORK_DIR)
            if f.startswith("DataWeb") and f.endswith(".csv")
        ]
    )

    if local and today_str in local[-1]:
        print(
            f"DCASA file for today ({local[-1]}) already exists. Skipping FTP download."
        )
        paths["dcasa"] = f"{WORK_DIR}/{local[-1]}"
    else:
        for attempt in range(1, 3):
            try:
                print(f"Downloading DCASA from FTP (attempt {attempt}/2)...")
                ftp = FTP(providers["dcasa"]["host"], timeout=10)
                ftp.set_pasv(True)
                ftp.login(
                    user=providers["dcasa"]["user"], passwd=providers["dcasa"]["pass"]
                )
                files = ftp.nlst()
                dfiles = sorted(
                    [f for f in files if f.startswith("DataWeb") and f.endswith(".csv")]
                )
                if dfiles:
                    path = f"{WORK_DIR}/{dfiles[-1]}"
                    # If local file already exists and is the same name, skip if you want,
                    # but let's re-download to be sure it's complete if it wasn't today's.
                    with open(path, "wb") as fp:
                        ftp.retrbinary(f"RETR {dfiles[-1]}", fp.write)
                    with open(path, encoding="latin-1") as chk:
                        row_count = sum(1 for _ in chk) - 1
                    if row_count < MIN_DCASA_ROWS:
                        print(
                            f"WARNING: DCASA file has only {row_count} rows (expected >{MIN_DCASA_ROWS})."
                        )
                    else:
                        paths["dcasa"] = path
                ftp.quit()
                break  # success
            except Exception as e:
                print(f"WARNING: FTP attempt {attempt} failed ({e}).")
                if attempt < 2:
                    wait_time = 5
                    print(f"Waiting {wait_time}s before next attempt...")
                    time.sleep(wait_time)
        local = sorted(
            [
                f
                for f in os.listdir(WORK_DIR)
                if f.startswith("DataWeb") and f.endswith(".csv")
            ]
        )
        # Skip files with < MIN_DCASA_ROWS rows
        for fname in reversed(local):
            fpath = f"{WORK_DIR}/{fname}"
            with open(fpath, encoding="latin-1") as chk:
                rc = sum(1 for _ in chk) - 1
            if rc >= MIN_DCASA_ROWS:
                paths["dcasa"] = fpath
                break

    if "dcasa" in paths:
        print(f"Using DCASA file: {os.path.basename(paths['dcasa'])}")

    # MINA & SIGNES
    mina_tmp = f"{WORK_DIR}/minerales_feed_tmp.xml"
    mina_target = f"{WORK_DIR}/minerales_feed.xml"
    subprocess.run(
        [
            "curl",
            "-f",
            "-L",
            providers["mina"]["url"],
            "-o",
            mina_tmp,
        ],
        capture_output=True,
    )
    if os.path.exists(mina_tmp) and os.path.getsize(mina_tmp) > 1000:
        with open(mina_tmp, "r", encoding="latin1", errors="ignore") as f:
            first_chunk = f.read(500).lower()
        if "<!doctype html" not in first_chunk and "<html" not in first_chunk:
            shutil.move(mina_tmp, mina_target)
            print("Successfully updated minerales_feed.xml")
        else:
            print("Downloaded minerales file is HTML error, keeping existing minerales_feed.xml")
            try:
                os.remove(mina_tmp)
            except OSError:
                pass
    else:
        if os.path.exists(mina_tmp):
            try:
                os.remove(mina_tmp)
            except OSError:
                pass
        print("Download of minerales feed failed or returned error, keeping existing minerales_feed.xml")
    paths["mina"] = mina_target

    signes_tmp = f"{WORK_DIR}/signes_stock_tmp.csv"
    signes_target = f"{WORK_DIR}/signes_stock.csv"
    subprocess.run(
        [
            "curl",
            "-f",
            "-L",
            providers["signes"]["url"],
            "-o",
            signes_tmp,
        ],
        capture_output=True,
    )
    if os.path.exists(signes_tmp) and os.path.getsize(signes_tmp) > 1000:
        shutil.move(signes_tmp, signes_target)
    elif os.path.exists(signes_tmp):
        try:
            os.remove(signes_tmp)
        except OSError:
            pass
    paths["signes"] = signes_target

    return paths


def run_fast_update():
    print(f"--- Fast Stock Update @ {datetime.datetime.now()} ---")

    if not os.path.exists(CATALOG_FILE) or not os.path.exists(EXCEL_PRICES_FILE):
        print("Error: JSON files missing. Run extract_full.py first.")
        return

    with open(CATALOG_FILE, "r") as f:
        catalog = json.load(f)
    with open(EXCEL_PRICES_FILE, "r") as f:
        excel_prices = json.load(f)
    print(f"Loaded {len(catalog)} SKUs mapping.")

    # Load Base Stocks for MD (from STOCK ES)
    BASE_STOCKS_FILE = f"{WORK_DIR}/base_stocks.json"
    base_stocks = {}
    if os.path.exists(BASE_STOCKS_FILE):
        with open(BASE_STOCKS_FILE, "r") as f:
            base_stocks = json.load(f)
        print(f"Loaded {len(base_stocks)} base stocks.")

    # Load trEDISER Stocks (MD) from Excel
    trediser_stocks = {}
    if os.path.exists(STOCK_TREDISER_FILE):
        print(f"Loading trEDISER stocks from {STOCK_TREDISER_FILE}...")
        try:
            trediser_stocks = load_trediser_stocks(STOCK_TREDISER_FILE)
            print(f"Loaded {len(trediser_stocks)} trEDISER stocks.")
        except Exception as e:
            print(f"Error loading trEDISER Excel: {e}")

    # Load Prices from precios ES.xlsx
    print(f"Loading prices from {PRECIOS_FILE}...")
    try:
        df_p = pd.read_excel(PRECIOS_FILE)
        # Ensure SKU is string and strip whitespace
        df_p["sku"] = df_p["sku"].astype(str).str.strip()
        # Clean price: convert to string, replace comma, then to numeric
        df_p["price_clean"] = (
            df_p["price"].astype(str).str.replace(",", ".").str.strip()
        )
        df_p["price_num"] = pd.to_numeric(df_p["price_clean"], errors="coerce").fillna(
            0
        )
        prices_map = df_p.set_index("sku")["price_num"].to_dict()
        print(f"Loaded {len(prices_map)} prices.")
    except Exception as e:
        print(f"Error loading prices from Excel: {e}")
        prices_map = {}

    paths = download_files()
    data = {}

    # Load Feeds
    if "dcasa" in paths:
        try:
            df = pd.read_csv(
                paths["dcasa"], sep=";", encoding="latin-1", on_bad_lines="skip"
            )
            df.columns = [c.strip() for c in df.columns]
            df["id_str"] = df["CODIGO"].astype(str).str.replace(".0", "", regex=False)
            data["DC"] = df.set_index("id_str")
        except:
            pass
    if "mina" in paths:
        try:
            df = pd.read_csv(
                paths["mina"],
                sep="|",
                encoding="utf-8",
                header=None,
                on_bad_lines="skip",
            )
            df[0] = df[0].astype(str).str.replace(".0", "", regex=False)
            data["VC"] = df.set_index(0)
        except:
            pass
    if "signes" in paths:
        try:
            df = pd.read_csv(
                paths["signes"],
                sep=";",
                encoding="latin-1",
                header=None,
                on_bad_lines="skip",
            )
            df[0] = (
                df[0]
                .astype(str)
                .str.replace("SG-", "")
                .str.strip()
                .str.replace(".0", "", regex=False)
            )
            data["SG"] = df.set_index(0)
        except:
            pass

    # Load Nuevos SG1
    NUEVOS_SG1_FILE = f"{WORK_DIR}/Nuevos SG1.xlsx"
    nuevos_sg1_data = []
    if os.path.exists(NUEVOS_SG1_FILE):
        print(f"Loading new SG SKUs from {NUEVOS_SG1_FILE}...")
        try:
            df_n = pd.read_excel(NUEVOS_SG1_FILE)
            if "item_sku" not in df_n.columns and not df_n.empty:
                df_n.columns = df_n.iloc[0]
                df_n = df_n[1:]

            # Filter for SGRI
            sg_new = df_n[df_n["item_sku"].astype(str).str.endswith("SGRI")]
            for _, row in sg_new.iterrows():
                sku = str(row["item_sku"]).strip()
                price = to_num(row["standard_price"])
                # Product number is the SKU without 'SGRI'
                prod_num = sku[:-4] if sku.endswith("SGRI") else sku
                nuevos_sg1_data.append(
                    {"sku": sku, "price": price, "prod_num": prod_num}
                )
            print(f"Loaded {len(nuevos_sg1_data)} new SG SKUs.")
        except Exception as e:
            print(f"Error loading Nuevos SG1: {e}")

    output_rows = [list(h) for h in AMAZON_NEW_TEMPLATE_HEADERS]

    print("Calculating updates...")
    seen_skus = set()
    for entry in catalog:
        sku = entry["sku"]
        if sku in seen_skus:
            continue
        seen_skus.add(sku)

        if sku in SKUS_ELIMINADOS:
            continue

        provider = entry["provider"]
        l_id = entry["id"]

        if sku in SKUS_MANUALES and provider != "MD":
            continue

        sheet_map = {
            "VC": "INICIOVC",
            "DC": "DcasaWeb",
            "SG": "Signes",
            "MD": "Madelcar",
        }
        sheet_name = sheet_map.get(provider)

        # Prepare IDs and secondary info
        lookup_id = l_id.replace(".0", "")
        provider_info = excel_prices.get(sheet_name, {}).get(l_id, {})
        divisor = (
            provider_info.get("divisor", 1.0)
            if isinstance(provider_info, dict)
            else 1.0
        )

        # FINAL PRICE: Always from precios ES.xlsx, ensuring it's a number
        final_price = to_num(prices_map.get(sku, 0))

        # Fixed Price Override
        if sku in PRECIOS_FIJOS:
            final_price = PRECIOS_FIJOS[sku]

        # Apply 5% increase + rounding to .95 for VC SKUs with price < 30
        if (
            sku not in PRECIOS_FIJOS
            and sku.endswith("VC")
            and final_price > 0
            and final_price < VC_PRICE_THRESHOLD
        ):
            final_price = apply_vc_price_increase(final_price)

        # Apply 5% increase + round so 2nd decimal is 5 for DC providers
        if sku not in PRECIOS_FIJOS and provider == "DC" and final_price > 0:
            final_price = final_price * 1.05
            final_price = round(int(final_price * 10) / 10.0 + 0.05, 2)

        # If not found in precios ES, use a default or 0 (User said prices ARE in that file)
        final_stock = 0

        try:
            if provider == "VC" and "VC" in data and lookup_id in data["VC"].index:
                row = data["VC"].loc[lookup_id]
                if isinstance(row, pd.DataFrame):
                    row = row.iloc[0]
                raw_stock = to_num(row[7])
                if raw_stock >= 5:
                    final_stock = raw_stock
                else:
                    final_stock = 0

            elif provider == "DC" and "DC" in data and lookup_id in data["DC"].index:
                row = data["DC"].loc[lookup_id]
                if isinstance(row, pd.DataFrame):
                    row = row.iloc[0]
                s_raw, p_cost = to_num(row["STOCK_DISPONIBLE"]), to_num(row["Tarifa A"])
                # Stock filter for DC (using divisor from excel_prices_final if needed for the stock threshold)
                if s_raw > 3 and (s_raw >= 20 or (p_cost * divisor) >= 20):
                    final_stock = 99 if s_raw > 1 else (1 if s_raw == 1 else 0)
                else:
                    final_stock = 0

            elif provider == "SG" and "SG" in data and lookup_id in data["SG"].index:
                row = data["SG"].loc[lookup_id]
                if isinstance(row, pd.DataFrame):
                    row = row.iloc[0]
                s_raw = to_num(row[2])
                # Stock filter for SG
                if s_raw > 3:
                    # Filter '83627' (User request)
                    if "83627" in sku:
                        final_stock = 0
                    else:
                        p_cost = to_num(row[3])
                        if s_raw >= 20 or p_cost >= 20:
                            final_stock = s_raw
                        else:
                            final_stock = 0
                else:
                    final_stock = 0

            elif provider == "MD":
                # Priority 1: STOCK TREDISER.xls (using numeric ID)
                # Priority 2: base_stocks.json (using full SKU)
                if l_id in trediser_stocks:
                    final_stock = trediser_stocks[l_id]
                else:
                    final_stock = 0

        except:
            pass

        # Force stock=0 for SKUs in SKUS_FORZAR_CERO (any provider)
        clean_sku_upper = str(sku).upper().strip()
        forzar_set = {str(s).upper().strip() for s in SKUS_FORZAR_CERO}
        if (
            clean_sku_upper in forzar_set
            or any(clean_sku_upper.startswith(s) for s in forzar_set)
            or (provider == "VC" and lookup_id in {"4085", "5324", "2598"})
        ):
            final_stock = 0

        # Final Safety Check
        if final_price <= 0:
            final_price = 0

        # Min/Max (Col C, D)
        min_p_val = round(final_price / 2, 2) if final_price > 0 else 0
        max_p_val = round(final_price * 2, 2) if final_price > 0 else 0

        # Format prices with COMMA as decimal separator for Amazon ES
        final_price_str = (
            f"{final_price:g}".replace(".", ",") if final_price > 0 else "0"
        )
        min_p_str = f"{min_p_val:g}".replace(".", ",") if min_p_val > 0 else ""
        max_p_str = f"{max_p_val:g}".replace(".", ",") if max_p_val > 0 else ""

        row_32 = [""] * 32
        row_32[0] = sku
        row_32[1] = "Logística por parte del vendedor (predeterminado)"
        row_32[2] = str(int(final_stock))
        row_32[6] = final_price_str
        row_32[8] = min_p_str
        row_32[9] = max_p_str
        output_rows.append(row_32)

    # Process Nuevos SG1 items that were not in catalog
    print("Processing items from Nuevos SG1...")
    for item in nuevos_sg1_data:
        sku = item["sku"]
        if sku in seen_skus:
            continue
        seen_skus.add(sku)

        if sku in SKUS_ELIMINADOS:
            continue

        if sku in SKUS_MANUALES:
            continue

        final_price = item["price"]
        prod_num = item["prod_num"]
        final_stock = 0

        try:
            if "SG" in data and prod_num in data["SG"].index:
                row = data["SG"].loc[prod_num]
                if isinstance(row, pd.DataFrame):
                    row = row.iloc[0]
                s_raw = to_num(row[2])
                p_cost = to_num(row[3])
                # Stock filter for SG
                if s_raw > 3:
                    if s_raw >= 20 or p_cost >= 20:
                        final_stock = s_raw
                    else:
                        final_stock = 0
                else:
                    final_stock = 0
        except:
            pass

        # Min/Max (Col C, D)
        min_p_val = round(final_price / 2, 2) if final_price > 0 else 0
        max_p_val = round(final_price * 2, 2) if final_price > 0 else 0

        # Format prices with COMMA
        final_price_str = (
            f"{final_price:g}".replace(".", ",") if final_price > 0 else "0"
        )
        min_p_str = f"{min_p_val:g}".replace(".", ",") if min_p_val > 0 else ""
        max_p_str = f"{max_p_val:g}".replace(".", ",") if max_p_val > 0 else ""

        row_32 = [""] * 32
        row_32[0] = sku
        row_32[1] = "Logística por parte del vendedor (predeterminado)"
        row_32[2] = str(int(final_stock))
        row_32[6] = final_price_str
        row_32[8] = min_p_str
        row_32[9] = max_p_str
        output_rows.append(row_32)

    print(f"Writing {len(output_rows)} rows to {OUTPUT_FILE}...")
    with open(OUTPUT_FILE, "w", encoding="utf-8", newline="") as f:
        writer = csv.writer(f, delimiter="\t")
        writer.writerows(output_rows)

    # Also save as CSV for compatibility
    for csv_name in ("STOCK AMZ.csv", "stock amz.csv"):
        shutil.copy2(OUTPUT_FILE, f"{WORK_DIR}/{csv_name}")

    print("Export finished.")


if __name__ == "__main__":
    run_fast_update()
