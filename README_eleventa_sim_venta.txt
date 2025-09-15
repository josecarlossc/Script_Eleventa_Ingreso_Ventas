ELEVENTA — Simulación de venta desde Excel
========================================

Este paquete incluye:
- eleventa_sim_venta.py  -> script que lee un Excel (.xlsx) con Código y Cantidad y simula teclearlo en Eleventa.
- requirements_eleventa_sim_venta.txt -> dependencias de Python.

Requisitos:
- Windows
- Python 3.9 o superior
- Teclado configurado para usar la tecla ENTER como “Agregar producto” en Eleventa (comportamiento por defecto).
- Impresora por defecto o configurada en Eleventa (para exportar PDF usa "Microsoft Print to PDF").

Instalación:
1) Instala Python si no lo tienes: https://www.python.org/downloads/
2) Abre un terminal/PowerShell en esta carpeta y ejecuta:
   pip install -r requirements_eleventa_sim_venta.txt

Tu Excel (.xlsx):
- Debe tener una columna de CÓDIGO (acepta nombres como: Código, SKU, Barcode, Producto, etc.).
- Opcionalmente una columna de CANTIDAD (Cantidad, Qty, Cant, Unidades).
- Si no hay encabezados, tomará la 1ª columna como Código y la 2ª (si existe) como Cantidad.

Uso típico (prueba):
   python eleventa_sim_venta.py r1.xlsx --window "eleventa" --dry-run

Ejecución real (añade artículos repitiendo ENTER por cantidad):
   python eleventa_sim_venta.py r1.xlsx --window "eleventa" --mode repeat

Ejecución alternativa (si Eleventa acepta 'codigo*cantidad' + Enter):
   python eleventa_sim_venta.py r1.xlsx --window "eleventa" --mode star

Parámetros útiles:
- --start-delay  Segundos de cuenta atrás antes de empezar (por defecto 5).
- --between-items Pausa entre artículos (por defecto 0.2s).
- --finalize     Al final intenta presionar F12 (Cobrar).
- --fast-enter   Presiona Enter inmediatamente tras cada escritura.
- --window       Texto del título de la ventana (por defecto "eleventa").

Seguridad:
- PyAutoGUI tiene un “failsafe”: mueve el mouse a la esquina superior izquierda para abortar.

Exportar PDF (manual):
1) Genera la venta en Eleventa.
2) Presiona “Reimprimir Último Ticket” o entra a Reportes (Ventas del día) y elige Imprimir.
3) En el diálogo de impresión, selecciona “Microsoft Print to PDF” y guarda el archivo.
