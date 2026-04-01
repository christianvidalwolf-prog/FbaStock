#!/bin/bash
cd /Users/christianvidalwolf/Stock
echo "Actualizando valores de STOCK AMZ.txt..."
python3 -c "from auto_update_stock import export_amazon; export_amazon()"
echo "Finalizado. Ya puedes cerrar esta ventana."
sleep 5
