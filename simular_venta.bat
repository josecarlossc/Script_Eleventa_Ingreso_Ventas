@echo off
setlocal
REM Ejecuta minimizado; cambia pythonw.exe si quieres sin consola.
start "" /min cmd /c python "%~dp0\eleventa_sim_venta.py" "%~dp0\r1.xlsx" --window "eleventa MultiCaja" --mode repeat --pre-offset 120 95 --click-each --enter-key numenter --interval 0.02 --between-items 0.3 --start-delay 7 --finalize
