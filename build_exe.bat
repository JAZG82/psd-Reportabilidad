@echo off
REM Genera dist\PSD_Reportabilidad.exe (ejecutar en Windows, con Python 3.10+ instalado)
python -m pip install -r requirements.txt
python -m PyInstaller --onefile --windowed --name PSD_Reportabilidad psd_app.py
echo.
echo Listo: dist\PSD_Reportabilidad.exe
pause
