# Auditoría de stock Amazon — 5 de septiembre de 2026

## Conclusión

La discrepancia principal está en DCASA y en la versión de fichero utilizada. El Excel local `STOCK+AMZ-Nueva_Versión.xlsm` (nombre con acento descompuesto en disco, modificado el 4 de septiembre a las 14:02) contiene 23.139 SKU y 3.274 con stock. El TXT local del 5 de septiembre contiene los mismos 23.139 SKU y 6.819 con stock. Su copia de Dropbox es idéntica byte por byte. La diferencia neta de 3.545 SKU con stock corresponde íntegramente a DCASA.

No se ha confirmado cuál se subió a Amazon ni se ha consultado su informe de procesamiento. La cifra comunicada de 3.276 es próxima a la del Excel, pero no coincide exactamente: no debe confundirse con un recuento verificado en Seller Central.

| Proveedor | Excel 4 septiembre | TXT 5 septiembre | Diferencia neta |
|---|---:|---:|---:|
| Signes Grimal | 1.418 | 1.418 | 0 |
| DCASA | 659 | 4.204 | +3.545 |
| Minerales / Vives Cortada | 921 | 921 | 0 |
| Trediser / MD | 276 | 276 | 0 |
| **Total SKU con stock** | **3.274** | **6.819** | **+3.545** |

Hay 3.547 cantidades diferentes entre Excel y TXT, todas DCASA: 3.546 pasan de cero a positivo y una pasa de positivo a cero. El TXT contiene una fila de ejemplo `ABC123`, cantidad 152, anterior al inicio de datos declarado; se excluye del recuento. Contarla daría 6.820, cifra preliminar corregida durante la auditoría.

## Descargas y causa técnica

### DCASA: caída anómala aceptada como fichero válido

`DataWeb20260831.csv` contiene solo 1.058 registros válidos de proveedor. `DataWeb20260904.csv` contiene 6.964. Los registros de las ejecuciones muestran el uso del fichero reducido durante varios días. El Excel examinado conserva el resultado reducido: 659 SKU de DCASA con stock.

El proceso acepta ficheros con al menos 1.000 filas (`auto_update_stock_fast.py:1256`), por lo que 1.058 supera el control pese a la pérdida de aproximadamente el 85 % del catálogo. No compara cobertura de referencias con el último fichero completo. Las referencias que no se encuentran reciben stock cero. Esto convierte una descarga anormalmente pequeña en miles de ceros en Amazon.

La evidencia prueba la reducción del fichero recibido y el fallo de validación local. No permite distinguir si el proveedor publicó un catálogo parcial o si hubo una transferencia incompleta: para eso se necesitaría una copia íntegra del archivo remoto de aquella fecha o registros de transferencia más detallados.

El generador actual escribe TXT y copias CSV; no actualiza el XLSM examinado. Por tanto, volver a subir ese XLSM reutilizaría sus datos antiguos aunque el TXT diario sea más reciente.

### Minerales / Vives Cortada: URL rota, stock sin refrescar

La URL configurada del feed devuelve HTTP 404 y una página HTML; comprobado mediante una descarga aislada durante esta auditoría. El log de hoy confirma que se conserva `minerales_feed.xml` anterior. Este fichero local, modificado el 4 de septiembre a las 12:35, tiene 5.676 referencias y solo 960 con stock positivo; produce 921 SKU de Amazon con stock tras cruces y exclusiones.

En las copias de los días 3 y 4, Minerales aportaba cero SKU con stock; el día 2 aportaba 1.927. Hoy se recuperan parcialmente con el fichero local. No está verificada la vigencia comercial de sus cantidades. No se ha encontrado una fuente nueva confirmada que sustituya al endpoint roto.

### Signes Grimal: descarga disponible, filtros y cobertura del catálogo

La URL responde HTTP 200 en la comprobación aislada. El fichero utilizado por la generación contiene 5.371 referencias, 5.288 con stock positivo. En Amazon llegan 1.418 SKU con stock, incluyendo los añadidos de `Nuevos SG1.xlsx`.

De los SKU exportados de este proveedor, 3.294 no encuentran su referencia en el fichero actual; 28 encuentran stock cero; 93 se anulan por tener como máximo 3 unidades; otros 306 por tener menos de 20 unidades y precio de origen inferior a 20 euros. Son reglas existentes, no fallos nuevos de descarga.

Hay 3.497 referencias positivas del proveedor sin una referencia asociada entre los SKU exportados. Esto señala una diferencia de cobertura con el catálogo de Amazon; no demuestra que deban subirse automáticamente, ni que exista un listing para cada una.

### Trediser / MD: fichero local, no descarga automática

El proceso lee `STOCK TREDISER.xls`, modificado el 4 de septiembre a las 12:31. No existe un paso de descarga de Trediser en el flujo diario examinado. Contiene 913 referencias con stock; de los 639 SKU MD exportados, 276 encuentran stock y 363 no encuentran la referencia. Quedan 637 referencias positivas del proveedor sin correspondencia entre los SKU exportados.

## Reconciliación de los ficheros actuales

| Proveedor | Referencias en origen | Referencias positivas | SKU exportados | SKU positivos |
|---|---:|---:|---:|---:|
| DCASA | 6.964 | 6.964 | 9.396 | 4.204 |
| Signes Grimal | 5.371 | 5.288 | 5.139 | 1.418 |
| Minerales | 5.676 | 960 | 7.912 | 921 |
| Trediser | 913 | 913 | 639 | 276 |

Además se exportan 53 SKU de proveedor desconocido con cantidad cero. Referencias de proveedor y SKU de Amazon son medidas diferentes: hay varias ofertas por referencia, exclusiones y referencias sin oferta asociada. No deben sumarse ni compararse como si fueran productos únicos.

En DCASA, 5.060 SKU no encuentran referencia y 132 se anulan por el filtro de 20 unidades o 20 euros (con divisor del catálogo). En Minerales, 2.528 no encuentran referencia, 4.361 tienen cero en origen y 102 con stock se anulan por exclusiones explícitas. El CSV de detalle permite revisar cada SKU.

## Historial de copias de Amazon

| Fecha | Signes | DCASA | Minerales | Trediser | Total positivo |
|---|---:|---:|---:|---:|---:|
| 1 septiembre | 1.428 | 4.223 | 1.930 | 274 | 7.855 |
| 2 septiembre | 1.425 | 659 | 1.927 | 274 | 4.285 |
| 3 septiembre | 1.419 | 659 | 0 | 274 | 2.352 |
| 4 septiembre, copia diaria | 1.418 | 659 | 0 | 274 | 2.351 |
| 4 septiembre, Excel posterior | 1.418 | 659 | 921 | 276 | 3.274 |
| 5 septiembre | 1.418 | 4.204 | 921 | 276 | 6.819 |

## Otras deficiencias comprobadas

- Dos LaunchAgents, `com.christian.stockupdate` y `com.christianvidalwolf.stockupdate`, programan el mismo flujo a las 07:00. El log del 5 de septiembre demuestra dos ejecuciones simultáneas. El flujo general carece de bloqueo y comparte nombres temporales de descargas y ficheros de salida. Es una condición de carrera; no se atribuye a ella la caída concreta de DCASA sin evidencia adicional.
- Los lectores de feeds y el cálculo de cantidades contienen `except: pass` y lectura CSV con `on_bad_lines="skip"`: pueden ocultar errores y convertirlos en cantidades cero.
- El catálogo de asignación es de junio y los divisores de marzo. Esto merece una reconciliación periódica con listings activos; no se deben crear ofertas a partir de referencias sin correspondencia de forma automática.

## Correcciones propuestas

1. Validar estructura, número de referencias y cobertura por proveedor antes de aceptar una descarga; detectar caídas bruscas como 6.964 a 1.058. Conservar el último fichero válido y señalar su antigüedad. Bloquear una exportación masiva de ceros por ausencia o fallo del feed.
2. Sustituir el endpoint de Minerales por uno vigente confirmado con el proveedor y verificar cantidades antes de refrescar Amazon.
3. Unificar la tarea programada y añadir bloqueo del proceso completo; descargas temporales únicas y reemplazos atómicos.
4. Establecer una salida diaria inequívoca, con fecha y resumen por proveedor. Si se usa XLSM, actualizarlo explícitamente desde la misma salida validada.
5. Automatizar la obtención de Trediser mediante una fuente confirmada y registrar su fecha efectiva.
6. Revisar cobertura de catálogo y filtros de stock por separado, conservando las exclusiones comerciales existentes.

## Verificación y alcance

Se ha reproducido el generador con los mismos ficheros locales, anulando descargas y escrituras en producción. Las 23.139 cantidades resultantes coinciden con el TXT actual. Se han contrastado cinco copias de Dropbox y los dos libros XLSM locales. Las comprobaciones nuevas de HTTP se guardaron separadas y no sustituyeron fuentes de producción.

No se han modificado scripts de producción, tareas programadas, stock comercial ni ficheros preparados para subir; tampoco se ha publicado ni enviado nada a Amazon.

Artefactos: `summary.json` (recuentos y hashes), `detalle_skus.csv` (causa por SKU), `diferencias_excel_txt.csv` (3.547 cambios DCASA), `download_checks.json` (HTTP), `audit.py` (reproducción local).
