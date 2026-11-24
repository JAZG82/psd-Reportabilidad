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
| Recepción y Lotes Chancado (hojas Registro-Lotes, RESUMEN, Mineral Interno) | Camiones recibidos con fecha (bruto, tara, neto), N° Met por lote, humedad y lotes especiales. La ley de cada lote sale del Log por su N° Met |
| Reporte Jefe de Turno | Molienda TA/TB, horas, chancado, agua de mar, reactivos y -#100 por turno (hojas diarias). Su filtrado solo se compara; el tonelaje oficial es el de Filtrado V-S |
| P80 (Concentrado, Rouger, Relave) | Granulometría |
| Filtrado V-S humedad | Toneladas oficiales por lote; la ley de cada lote sale del Log por su N° Met |

El mes analizado es el del Reporte Jefe de Turno cargado.

## Pestañas de recepción de mineral
- **Recepción:** toneladas recibidas por día (mes en curso), semana o mes, con camiones, secas, ley CuT ponderada y Cu fino. En vista diaria compara recibido y molido acumulados.
- **Lotes:** ley CuT de cada lote recibido y tabla con N° Met, humedad, CuT, CuS, Ag y estado. Los lotes sin ley en el Log se marcan.
- **Acopio:** stock reportado por el Jefe de Turno en cada turno (cancha de chancado y pila de molienda, mina y especial) y su variación comparada con recibido menos molido (orientativa).

## Reglas
- Las leyes salen solo del Log; nunca se estiman. Si hay más de un "Rep" para la misma muestra, no se elige: se usa la simple y se avisa en Conciliación.
- Recuperación por leyes (dos productos) solo es válida si conc > cabeza > relave; si no, el turno se marca en Conciliación.
- Recuperación por balance de masa se muestra acumulada y solo desde 10 t de Cu alimentado (por turno oscila por desfase de filtrado).
- Las toneladas de recepción salen de los camiones de Registro-Lotes (neto); si bruto − tara no coincide con el neto, se avisa en Conciliación.
- Fechas dd-mm-aa; números con coma decimal.

## Limitaciones v1
- Reactivos: promedio por turno de las lecturas de las hojas diarias del Jefe de Turno (alias de rótulos incluidos). El pH no se registra en esas hojas.
- Sin corrección por inventario en espesadores.
- El ejecutable se genera en el mismo sistema operativo donde se usará.

## Listado semanal de pendientes (miércoles)

Botón **Pendientes para metalurgistas…** (barra superior): genera, a partir de las alertas detectadas con los archivos cargados, un texto agrupado por prioridad (alta / media / baja), numerado, con fecha de emisión y corte de datos, y la solicitud de responsable y fecha de cierre. Indica si hoy es miércoles o cuál es el próximo envío. El texto es editable; se puede **Copiar** o **Guardar .txt**.

El programa no envía correos: el listado se prepara y se pega en el mail/Teams. Cargue los archivos actualizados antes de generarlo.

**Enviar por correo…**: abre su programa de correo predeterminado (Outlook, Mail, etc.) con asunto y destinatarios ya puestos (campo «Para», se recuerda entre sesiones). Si el listado es largo, el cuerpo queda con un aviso y el texto completo ya está copiado: péguelo con Ctrl+V / Cmd+V. Requiere un cliente de correo configurado en el equipo; el envío final lo hace usted en ese programa.

## Seguimiento de pendientes (compartido)

En **Pendientes para metalurgistas…**, la pestaña *Seguimiento* permite indicar por cada pendiente: **estado** (Pendiente / En proceso / Cerrado), **responsable**, **fecha de cierre** y comentario. Cada fila muestra quién y cuándo la modificó por última vez, y arriba se ve la última actualización de cualquier usuario.

**Compartir entre varias personas (SharePoint):**
1. En SharePoint, cree una carpeta (p. ej. `PSD_Seguimiento`) en la biblioteca del equipo y use **Sincronizar** para que aparezca en el computador (OneDrive).
2. En el programa: *Elegir carpeta compartida…* y seleccione esa carpeta. Cada persona lo hace una vez en su equipo.
3. Todos escriben en el mismo archivo `seguimiento_pendientes.json`. Al guardar, el programa vuelve a leer el archivo y fusiona: **por cada pendiente prevalece la modificación más reciente**. También fusiona las copias de conflicto que OneDrive genere (`seguimiento_pendientes*.json`).
4. *Recargar* trae lo que otros hayan cambiado mientras la ventana estaba abierta.

Sin carpeta compartida, el seguimiento se guarda solo en el computador (`~/.psd_reportabilidad/`). Los pendientes vencidos salen en rojo y como «VENCIDO» en el mensaje. Si un pendiente deja de detectarse en los datos, el mensaje lo lista como «confirmar cierre».

Limitación: la sincronización depende de OneDrive; si dos personas editan el mismo pendiente a la vez, gana la edición más reciente.

## Datos compartidos (últimos datos para todos)

Con la carpeta compartida configurada (botón **Carpeta compartida…**), cada vez que alguien carga planillas el programa **publica una copia** en la subcarpeta `datos/` (una por tipo: log, control, jefe, p80, oficial, recepcion) y registra quién y cuándo. Quien abra el programa después **carga automáticamente los últimos datos** (también con **Cargar datos compartidos**), sin necesidad de tener los Excel. Todos ven los mismos pendientes.

Si el archivo que alguien abre es **más antiguo** que lo ya publicado, el programa avisa y pregunta antes de reemplazar. La comparación usa la fecha de modificación del archivo. Las planillas se leen y publican tal como están: no se modifican.

## Tema oscuro y colores

Toda la ventana usa tema oscuro (letras blancas sobre fondo oscuro). Semáforo único en tablas y alertas: **alta** = fondo rojo con letras amarillas; **media** = fondo amarillo con letras negras; **cerrado** = fondo blanco con letras negras. Los colores de los gráficos están validados para fondo oscuro y para daltonismo.
