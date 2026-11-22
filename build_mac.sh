#!/bin/bash
# Genera dist/PSD_Reportabilidad.app (ejecutar SOLO después de que ejecutar_mac.command funcione bien)
cd "$(dirname "$0")"
[ -d .venv ] || ./ejecutar_mac.command
.venv/bin/pip install pyinstaller
.venv/bin/python -m PyInstaller --windowed --noconfirm --name PSD_Reportabilidad psd_app.py
echo "Listo: dist/PSD_Reportabilidad.app"
