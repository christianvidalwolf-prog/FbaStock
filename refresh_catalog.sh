#!/bin/bash
cd /Users/christianvidalwolf/Stock
echo "Actualizando definiciones del catálogo y precios base desde el archivo Excel..."
echo "Este proceso puede tardar un par de minutos debido al gran tamaño del archivo."
/Library/Developer/CommandLineTools/usr/bin/python3 extract_full.py
echo "Definiciones y precios base actualizados correctamente."
sleep 3
