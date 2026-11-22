#!/bin/bash
# Doble clic para abrir el programa en Mac (crea un entorno la primera vez; requiere internet y Python 3.10+)
cd "$(dirname "$0")"
if [ ! -d .venv ]; then
  python3 -m venv .venv || { echo "Falta Python 3. Instálelo desde python.org"; read -p "Enter"; exit 1; }
  .venv/bin/pip install --upgrade pip
  .venv/bin/pip install PySide6-Essentials matplotlib openpyxl
fi
.venv/bin/python psd_app.py
