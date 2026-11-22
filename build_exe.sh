#!/bin/sh
# Genera dist/PSD_Reportabilidad (ejecutar en el mismo sistema operativo donde se usará)
python3 -m pip install -r requirements.txt
python3 -m PyInstaller --onefile --windowed --name PSD_Reportabilidad psd_app.py
