# PSD Reportabilidad – Planta Santo Domingo

Programa de escritorio (sin licencias de pago) que lee las planillas Excel de la planta y muestra
alimentación, leyes, recuperaciones, filtrado, P80 y reactivos en tablas y gráficos, contra el plan del mes.

## Uso en Mac (primera etapa de pruebas)
1. Instalar Python 3.10 o superior desde python.org (si no está).
2. Descomprimir el zip y, en Terminal, dentro de la carpeta: `chmod +x ejecutar_mac.command build_mac.sh`
3. Doble clic en `ejecutar_mac.command` (la primera vez descarga las librerías; si macOS lo bloquea: clic derecho > Abrir).
4. Cuando funcione sin problemas, `./build_mac.sh` genera `dist/PSD_Reportabilidad.app`.

## Uso
1. Ejecutar `python psd_app.py` (o el .exe generado con `build_exe.bat`).
2. Cargar archivos: botón **Abrir archivos…**, **Abrir carpeta…** o arrastrar y soltar. El tipo de archivo se detecta por sus hojas.
3. Ajustar el plan del mes arriba (molienda, ley cabeza, recuperación, ley concentrado, días). Todo se recalcula.
4. **Exportar a Excel…** guarda las tablas de cada pestaña.

## Archivos que lee (los de origen; el Análisis Corrientes NO se carga)
| Archivo | Aporta |
|---|---|
| Log de Muestras Lux Mining (hoja Log Muestras) | CuT, CuS, Ag por N° Met. La repetición "Rep Met-N" tiene prioridad sobre la muestra simple |
| Control Muestras Planta (hoja Ingresar Datos) | Qué Met corresponde a cada turno y corriente (cabeza M04, concentrado re-cleaner, relave final) |
| Reporte Jefe de Turno | Molienda TA/TB, horas, chancado, agua de mar, reactivos y -#100 por turno (hojas diarias). Su filtrado solo se compara; el tonelaje oficial es el de Filtrado V-S |
| P80 (Concentrado, Rouger, Relave) | Granulometría |
| Filtrado V-S humedad | Toneladas oficiales por lote; la ley de cada lote sale del Log por su N° Met |

El mes analizado es el del Reporte Jefe de Turno cargado.

## Reglas
- Las leyes salen solo del Log; nunca se estiman. Si hay más de un "Rep" para la misma muestra, no se elige: se usa la simple y se avisa en Conciliación.
- Recuperación por leyes (dos productos) solo es válida si conc > cabeza > relave; si no, el turno se marca en Conciliación.
- Recuperación por balance de masa se muestra acumulada y solo desde 10 t de Cu alimentado (por turno oscila por desfase de filtrado).
- Fechas dd-mm-aa; números con coma decimal.

## Limitaciones v1
- Reactivos: promedio por turno de las lecturas de las hojas diarias del Jefe de Turno (alias de rótulos incluidos). El pH no se registra en esas hojas.
- Sin corrección por inventario en espesadores.
- El ejecutable se genera en el mismo sistema operativo donde se usará.
