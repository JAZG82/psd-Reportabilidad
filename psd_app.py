"""PSD Reportabilidad — programa de escritorio.

Carga las planillas de la planta y muestra parámetros de operación, leyes,
recuperaciones, filtrado, P80 y reactivos como tablas y gráficos.
Uso:  python psd_app.py [archivo.xlsx ...]
"""
import os
import sys
import datetime as dt

from PySide6.QtCore import Qt, QTimer
from PySide6.QtGui import QColor, QFont
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QTabWidget,
    QPushButton, QLabel, QFileDialog, QTableWidget, QTableWidgetItem,
    QDoubleSpinBox, QSpinBox, QMessageBox, QSplitter, QAbstractItemView,
    QHeaderView, QFrame)

import matplotlib
matplotlib.use("QtAgg")
from matplotlib.figure import Figure
from matplotlib.backends.backend_qtagg import (
    FigureCanvasQTAgg as Canvas, NavigationToolbar2QT as Toolbar)

import psd_core as core

# ------------------------------------------------------------------ estilo
BLUE, ORANGE, AQUA, YELLOW = "#2a78d6", "#eb6834", "#1baf7a", "#eda100"
MAGENTA, GREEN, VIOLET, RED = "#e87ba4", "#008300", "#6250d6", "#e34948"
INK, INK2, GRID, SURFACE, PLAN = "#0b0b0b", "#52514e", "#e6e5e1", "#fcfcfb", "#8a8984"
BLUE_LIGHT = "#b7d3f6"
# entidad -> color fijo en toda la aplicación
C_CAB, C_CONC, C_REL, C_REC = BLUE, ORANGE, AQUA, VIOLET
C_TA, C_TB = BLUE, ORANGE

matplotlib.rcParams.update({
    "figure.dpi": 100, "figure.facecolor": SURFACE, "axes.facecolor": SURFACE,
    "axes.edgecolor": GRID, "axes.labelcolor": INK2, "axes.titlecolor": INK,
    "axes.titlesize": 10, "axes.titleweight": "semibold", "axes.titlelocation": "left",
    "axes.labelsize": 8.5, "xtick.color": INK2, "ytick.color": INK2,
    "xtick.labelsize": 8, "ytick.labelsize": 8, "text.color": INK,
    "axes.spines.top": False, "axes.spines.right": False,
    "axes.grid": True, "grid.color": GRID, "grid.linewidth": 0.8,
    "axes.axisbelow": True, "legend.frameon": False, "legend.fontsize": 8,
    "lines.linewidth": 1.8,
})


def fn(v, n=1, suf=""):
    """Número con formato chileno (miles con punto, decimales con coma)."""
    if v is None:
        return "—"
    s = f"{v:,.{n}f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return s + suf


def fd(d):
    return d.strftime("%d-%m-%y") if d else "—"


# ----------------------------------------------------------- hover en gráficos
class Hover:
    """Etiqueta con el valor del punto más cercano al cursor."""

    def __init__(self, canvas, ax, pts):
        self.canvas, self.ax = canvas, ax
        self.pts = [p for p in pts if p[1] is not None]
        self.ann = ax.annotate(
            "", xy=(0, 0), xytext=(12, 12), textcoords="offset points", fontsize=8,
            bbox=dict(boxstyle="round,pad=0.35", fc="white", ec="#b8b7b0"), zorder=20)
        self.ann.set_visible(False)
        canvas.mpl_connect("motion_notify_event", self.move)

    def move(self, ev):
        vis = self.ann.get_visible()
        if ev.inaxes is not self.ax or not self.pts:
            if vis:
                self.ann.set_visible(False)
                self.canvas.draw_idle()
            return
        tr = self.ax.transData
        best, dmin = None, 1e12
        for x, y, txt in self.pts:
            px, py = tr.transform((x, y))
            dd = (px - ev.x) ** 2 + (py - ev.y) ** 2
            if dd < dmin:
                best, dmin = (x, y, txt), dd
        if best and dmin < 28 ** 2:
            self.ann.xy = (best[0], best[1])
            self.ann.set_text(best[2])
            self.ann.set_visible(True)
            self.canvas.draw_idle()
        elif vis:
            self.ann.set_visible(False)
            self.canvas.draw_idle()


def eje_turnos(ax, filas, cada=2):
    ticks = [i for i, r in enumerate(filas) if r["turno"] == "TA" and r["fecha"].day % cada == 1]
    ax.set_xticks(ticks)
    ax.set_xticklabels([filas[i]["fecha"].strftime("%d-%m") for i in ticks])
    ax.set_xlim(-1, len(filas))


def eje_dias(ax, dias, cada=2):
    ticks = [i for i, r in enumerate(dias) if r["fecha"].day % cada == 1]
    ax.set_xticks(ticks)
    ax.set_xticklabels([dias[i]["fecha"].strftime("%d-%m") for i in ticks])
    ax.set_xlim(-0.8, len(dias) - 0.2)


def linea_plan(ax, y, texto, abajo=False):
    ax.axhline(y, color=PLAN, lw=1.2, ls=(0, (5, 3)), zorder=1)
    ax.annotate(texto, xy=(1, y), xycoords=("axes fraction", "data"), xytext=(-3, -3 if abajo else 3),
                textcoords="offset points", ha="right", va="top" if abajo else "bottom", fontsize=8, color=INK2)


NAN = float("nan")


def con_huecos(tv, clave, ok=None):
    """Serie por turno con NaN donde no hay dato, para que la línea no una turnos no consecutivos."""
    out = []
    for r in tv:
        v = r[clave]
        out.append(v if (v is not None and v > 0 and (ok is None or ok(r))) else NAN)
    return out


def filas_con_datos(tv, claves):
    return [r for r in tv if r["molido"] > 0 or any(r[k] for k in claves)]


# ------------------------------------------------------------- página genérica
class Pagina(QWidget):
    """Gráfico arriba (con barra de herramientas) y tabla abajo."""

    def __init__(self, alto_fig=5.2, con_fig=True, reparto=(3, 2), estirar_ultima=False):
        super().__init__()
        self.reparto = reparto
        self.estirar_ultima = estirar_ultima
        self._ajustada = False
        lay = QVBoxLayout(self)
        lay.setContentsMargins(6, 6, 6, 6)
        self.split = QSplitter(Qt.Vertical)
        lay.addWidget(self.split)
        self.fig = None
        if con_fig:
            self.fig = Figure(figsize=(10, alto_fig), constrained_layout=True)
            self.canvas = Canvas(self.fig)
            cont = QWidget()
            v = QVBoxLayout(cont)
            v.setContentsMargins(0, 0, 0, 0)
            v.addWidget(Toolbar(self.canvas, self))
            v.addWidget(self.canvas, 1)
            self.split.addWidget(cont)
        self.tabla = QTableWidget()
        self.tabla.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.tabla.setAlternatingRowColors(True)
        self.tabla.verticalHeader().setVisible(False)
        self.split.addWidget(self.tabla)
        self.split.setStretchFactor(0, reparto[0])
        self.split.setStretchFactor(1, reparto[1])

    def ajustar(self):
        if self.fig is None or self._ajustada or self.height() < 200:
            return
        a, b = self.reparto
        h = self.height()
        self.split.setSizes([h * a // (a + b), h * b // (a + b)])
        self._ajustada = True

    def llenar(self, cabeceras, filas, colores=None, ancho_col=None):
        t = self.tabla
        t.setSortingEnabled(False)
        t.clear()
        t.setColumnCount(len(cabeceras))
        t.setRowCount(len(filas))
        t.setHorizontalHeaderLabels(cabeceras)
        for i, fila in enumerate(filas):
            for j, v in enumerate(fila):
                it = QTableWidgetItem(v if isinstance(v, str) else str(v))
                it.setTextAlignment((Qt.AlignRight if j > 0 and _es_num(v) else Qt.AlignLeft) | Qt.AlignVCenter)
                if colores and colores[i]:
                    it.setBackground(QColor(colores[i]))
                t.setItem(i, j, it)
        t.resizeColumnsToContents()
        t.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        t.horizontalHeader().setStretchLastSection(self.estirar_ultima)

    def limpiar(self):
        if self.fig is not None:
            self.fig.clear()

    def dibujar(self):
        if self.fig is not None:
            self.canvas.draw_idle()


def _es_num(v):
    if not isinstance(v, str):
        return True
    s = v.replace(".", "").replace(",", "").replace("-", "").replace("+", "").replace("%", "").strip()
    return s.isdigit()


def vacio(ax, msg="Sin datos cargados para este gráfico"):
    ax.text(0.5, 0.5, msg, transform=ax.transAxes, ha="center", va="center", color=INK2)
    ax.set_xticks([])
    ax.set_yticks([])
    ax.grid(False)


# ---------------------------------------------------------------- ventana
class Ventana(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PSD · Reportabilidad — Planta Santo Domingo")
        self.resize(1360, 860)
        self.setAcceptDrops(True)
        self.datos = core.Datos()
        self.R = None

        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)

        # barra superior: archivos
        top = QHBoxLayout()
        b = QPushButton("Abrir archivos…")
        b.clicked.connect(self.abrir)
        top.addWidget(b)
        b = QPushButton("Abrir carpeta…")
        b.clicked.connect(self.abrir_carpeta)
        top.addWidget(b)
        b = QPushButton("Limpiar")
        b.clicked.connect(self.limpiar)
        top.addWidget(b)
        b = QPushButton("Exportar a Excel…")
        b.clicked.connect(self.exportar)
        top.addWidget(b)
        top.addSpacing(14)
        self.lbl_fuentes = QLabel()
        top.addWidget(self.lbl_fuentes, 1)
        root.addLayout(top)

        # barra de plan
        pl = QHBoxLayout()
        pl.addWidget(QLabel("<b>Plan del mes</b>"))
        self.sp = {}
        for clave, rot, vmin, vmax, dec, val, suf in (
                ("molienda", "Molienda", 0, 200000, 0, 21300, " t"),
                ("ley", "Ley cabeza", 0, 10, 2, 1.14, " %"),
                ("rec", "Recuperación", 0, 100, 1, 77, " %"),
                ("ley_conc", "Ley concentrado", 0, 40, 1, 23, " %")):
            pl.addWidget(QLabel(rot))
            s = QDoubleSpinBox()
            s.setRange(vmin, vmax)
            s.setDecimals(dec)
            s.setValue(val)
            s.setSuffix(suf)
            s.setGroupSeparatorShown(True)
            s.valueChanged.connect(self.refrescar)
            pl.addWidget(s)
            self.sp[clave] = s
        pl.addWidget(QLabel("Días del mes"))
        s = QSpinBox()
        s.setRange(28, 31)
        s.setValue(30)
        s.valueChanged.connect(self.refrescar)
        pl.addWidget(s)
        self.sp["dias"] = s
        pl.addStretch(1)
        root.addLayout(pl)

        # pestañas
        self.tabs = QTabWidget()
        root.addWidget(self.tabs, 1)
        self.p = {}
        for clave, titulo, kw in (
                ("resumen", "Resumen", {"alto_fig": 3.4, "reparto": (2, 3), "estirar_ultima": True}), ("oper", "Operación", {"alto_fig": 6.2}),
                ("leyes", "Leyes", {"alto_fig": 6.2}), ("rec", "Recuperación", {}),
                ("filt", "Filtrado", {}), ("p80", "P80", {"alto_fig": 6.2}),
                ("reac", "Reactivos", {"alto_fig": 8.0}), ("conc", "Conciliación", {"con_fig": False, "estirar_ultima": True})):
            pg = Pagina(**kw)
            self.p[clave] = pg
            self.tabs.addTab(pg, titulo)
        self.tabs.currentChanged.connect(lambda _: self._ajustar_actual())
        self.lbl_cob = QLabel()
        self.lbl_cob.setWordWrap(True)
        self.lbl_cob.setContentsMargins(6, 6, 6, 6)
        self.p["conc"].layout().insertWidget(0, self.lbl_cob)
        self.refrescar()

    def _ajustar_actual(self):
        pg = self.tabs.currentWidget()
        if isinstance(pg, Pagina):
            pg.ajustar()

    def showEvent(self, e):
        super().showEvent(e)
        QTimer.singleShot(0, self._ajustar_actual)

    # -- archivos -------------------------------------------------------------
    def dragEnterEvent(self, e):
        if e.mimeData().hasUrls():
            e.acceptProposedAction()

    def dropEvent(self, e):
        self.cargar([u.toLocalFile() for u in e.mimeData().urls()])

    def abrir(self):
        rutas, _ = QFileDialog.getOpenFileNames(self, "Seleccione las planillas", "", "Excel (*.xlsx *.xlsm)")
        if rutas:
            self.cargar(rutas)

    def abrir_carpeta(self):
        c = QFileDialog.getExistingDirectory(self, "Seleccione la carpeta con las planillas")
        if c:
            self.cargar([os.path.join(c, f) for f in sorted(os.listdir(c))
                         if f.lower().endswith((".xlsx", ".xlsm")) and not f.startswith("~$")])

    def cargar(self, rutas):
        errores = []
        QApplication.setOverrideCursor(Qt.WaitCursor)
        try:
            for r in rutas:
                if not r.lower().endswith((".xlsx", ".xlsm")):
                    continue
                try:
                    self.datos.cargar(r)
                except Exception as ex:  # archivo no reconocido o dañado
                    errores.append(f"{os.path.basename(r)}: {ex}")
            self.refrescar()
        finally:
            QApplication.restoreOverrideCursor()
        if errores:
            QMessageBox.warning(self, "Archivos no cargados", "\n".join(errores))

    def limpiar(self):
        self.datos = core.Datos()
        self.refrescar()

    # -- refresco general -------------------------------------------------------
    def plan(self):
        return dict(molienda=self.sp["molienda"].value(), ley=self.sp["ley"].value(),
                    rec=self.sp["rec"].value(), ley_conc=self.sp["ley_conc"].value(),
                    dias=self.sp["dias"].value())

    def refrescar(self, *_):
        if not hasattr(self, "p"):
            return
        d = self.datos
        partes = []
        for k, nombre in core.NOMBRES.items():
            partes.append(f"<span style='color:{'#0a7a3d' if k in d.fuentes else '#a0a09a'}'>"
                          f"{'●' if k in d.fuentes else '○'} {nombre}</span>")
        self.lbl_fuentes.setText("&nbsp;&nbsp;".join(partes))
        self.R = core.calcular(d, self.plan())
        R = self.R
        for fn_ in (self.pg_resumen, self.pg_oper, self.pg_leyes, self.pg_rec,
                    self.pg_filt, self.pg_p80, self.pg_reac, self.pg_conc):
            fn_(R)

    # turnos con datos hasta el último día registrado
    def turnos_vis(self, R):
        u = R.K["ultimo_dia"]
        return [r for r in R.turnos if u and r["fecha"] <= u]

    def dias_vis(self, R):
        u = R.K["ultimo_dia"]
        return [r for r in R.dias if u and r["fecha"] <= u]

    # -- Resumen --------------------------------------------------------------
    def pg_resumen(self, R):
        pg, K, pl = self.p["resumen"], R.K, R.plan
        pg.limpiar()
        t_trans = K["dias_transc"] / pl["dias"] * 100 if pl["dias"] else None

        def av(real, plan_):
            return real / plan_ * 100 if (plan_ and real is not None) else None
        cu_alim_plan = pl["molienda"] * pl["ley"] / 100
        filas = [
            ("Tiempo transcurrido del mes", f"{K['dias_transc']} de {pl['dias']} días", "", fn(t_trans, 1, " %"),
             "hasta el último día con registro"),
            ("Molienda (t)", fn(K["molido"], 0), fn(pl["molienda"], 0), fn(av(K["molido"], pl["molienda"]), 1, " %"),
             f"{K['n_turnos_molino']} turnos con molino operando"),
            ("Horas molino (h)", fn(K["horas"], 1), "", "", "Reporte Jefe de Turno"),
            ("Rate molino (t/h)", fn(K["rate"], 1), "", "", "molido / horas operadas"),
            ("Ley cabeza CuT (%)", fn(K["cab"], 3), fn(pl["ley"], 2), fn(av(K["cab"], pl["ley"]), 0, " %"),
             f"ponderada por molido, {K['n_con_ley']} turnos con ley"),
            ("Ley cabeza CuS (%)", fn(K["cab_cus"], 3), "", "", "cobre soluble en alimentación"),
            ("Cu fino alimentado (t)", fn(K["cu_alim"], 1), fn(cu_alim_plan, 1), "", "solo turnos con ley de cabeza"),
            ("Concentrado filtrado seco (DMT)", fn(K["dry"], 1), fn(K["dmt_plan"], 1), fn(av(K["dry"], K["dmt_plan"]), 1, " %"),
             "plan = Cu fino plan / ley de concentrado"),
            ("Concentrado filtrado húmedo (t)", fn(K["wet"], 1), "", "", f"humedad ponderada {fn(K['hum'], 2, ' %')}"),
            ("Ley concentrado CuT (%)", fn(K["conc"], 2), fn(pl["ley_conc"], 1), fn(av(K["conc"], pl["ley_conc"]), 0, " %"),
             f"{K['lotes_bajo_obj']} de {K['lotes_con_ley']} lotes bajo el objetivo"),
            ("Cu fino filtrado (t)", fn(K["cu_filt"], 1), fn(K["cu_plan"], 1), fn(av(K["cu_filt"], K["cu_plan"]), 1, " %"),
             f"solo lotes con ley; {fn(K['dry_sin_ley'], 1)} t secas sin ley"),
            ("Recuperación balance de masa (%)", fn(K["rec_bm"], 1), fn(pl["rec"], 1),
             (("+" if K["rec_bm"] - pl["rec"] >= 0 else "") + fn(K["rec_bm"] - pl["rec"], 1, " pp")) if K["rec_bm"] else "",
             "acumulado del periodo; sesgo a la baja por lotes sin ley"),
            ("Recuperación por leyes (%)", fn(K["rec_leyes"], 1), fn(pl["rec"], 1),
             (("+" if K["rec_leyes"] - pl["rec"] >= 0 else "") + fn(K["rec_leyes"] - pl["rec"], 1, " pp")) if K["rec_leyes"] else "",
             f"dos productos, ponderada por molido, {K['n_rec']} turnos"),
        ]
        pg.llenar(["Indicador", "Real", "Plan", "Avance / diferencia", "Nota"], filas)
        pg.tabla.setColumnWidth(0, 270)
        ax = pg.fig.add_subplot(111)
        if K["molido"] == 0 and K["dry"] == 0:
            vacio(ax, "Cargue Log de Muestras, Control Muestras y Reporte de Jefe de Turno")
            pg.dibujar()
            return
        nombres = ["Molienda", "Concentrado seco (DMT)", "Cu fino filtrado"]
        vals = [av(K["molido"], pl["molienda"]), av(K["dry"], K["dmt_plan"]), av(K["cu_filt"], K["cu_plan"])]
        colores = [BLUE, ORANGE, AQUA]
        y = range(len(nombres))
        ax.barh(list(y), [v or 0 for v in vals], color=colores, height=0.55)
        for i, v in enumerate(vals):
            ax.text((v or 0) + 1.2, i, fn(v, 1, " %"), va="center", fontsize=9, color=INK)
        ax.set_yticks(list(y))
        ax.set_yticklabels(nombres)
        ax.invert_yaxis()
        ax.set_xlim(0, max(110, max(v or 0 for v in vals) + 15))
        if t_trans:
            ax.axvline(t_trans, color=PLAN, lw=1.3, ls=(0, (5, 3)))
            ax.annotate(f"Tiempo transcurrido {fn(t_trans, 0, ' %')}", xy=(t_trans, 1), xycoords=("data", "axes fraction"),
                        xytext=(4, -2), textcoords="offset points", va="top", fontsize=8, color=INK2)
        ax.set_title("Avance del mes contra plan (%)")
        ax.grid(axis="y", visible=False)
        pg.dibujar()

    # -- Operación ------------------------------------------------------------
    def pg_oper(self, R):
        pg = self.p["oper"]
        pg.limpiar()
        dias = self.dias_vis(R)
        axs = pg.fig.subplots(2, 2)
        if not dias:
            for a in axs.flat:
                vacio(a)
            pg.llenar([], [])
            pg.dibujar()
            return
        x = list(range(len(dias)))
        plan_dia = R.plan["molienda"] / R.plan["dias"]
        a = axs[0][0]
        a.bar(x, [r["mol_ta"] for r in dias], color=C_TA, width=0.72, label="TA")
        a.bar(x, [r["mol_tb"] for r in dias], bottom=[r["mol_ta"] for r in dias], color=C_TB, width=0.72,
              edgecolor=SURFACE, linewidth=1, label="TB")
        linea_plan(a, plan_dia, f"Plan {fn(plan_dia, 0)} t/día")
        a.set_title("Molienda diaria (t)")
        a.legend(loc="upper left", ncols=2)
        eje_dias(a, dias)
        Hover(pg.canvas, a, [(i, r["mol_ta"] + r["mol_tb"],
                              f"{fd(r['fecha'])}\nTA {fn(r['mol_ta'], 0)} t\nTB {fn(r['mol_tb'], 0)} t")
                             for i, r in enumerate(dias)])
        a = axs[0][1]
        a.bar(x, [r["h_ta"] for r in dias], color=C_TA, width=0.72, label="TA")
        a.bar(x, [r["h_tb"] for r in dias], bottom=[r["h_ta"] for r in dias], color=C_TB, width=0.72,
              edgecolor=SURFACE, linewidth=1, label="TB")
        linea_plan(a, 24, "24 h")
        a.set_title("Horas de molino operando (h)")
        a.legend(loc="upper left", ncols=2)
        eje_dias(a, dias)
        Hover(pg.canvas, a, [(i, r["h_ta"] + r["h_tb"], f"{fd(r['fecha'])}\n{fn(r['h_ta'] + r['h_tb'], 1)} h")
                             for i, r in enumerate(dias)])
        a = axs[1][0]
        rate = [((r["mol_ta"] + r["mol_tb"]) / (r["h_ta"] + r["h_tb"])) if (r["h_ta"] + r["h_tb"]) > 0 else None
                for r in dias]
        a.plot([i for i, v in enumerate(rate) if v], [v for v in rate if v], color=BLUE, marker="o", ms=4)
        a.set_title("Rate de molienda (t/h)")
        eje_dias(a, dias)
        Hover(pg.canvas, a, [(i, v, f"{fd(dias[i]['fecha'])}\n{fn(v, 1)} t/h") for i, v in enumerate(rate)])
        a = axs[1][1]
        a.bar(x, [r["chanc"] or 0 for r in dias], color=BLUE, width=0.72)
        a.set_title("Alimentación a chancado (t)")
        eje_dias(a, dias)
        Hover(pg.canvas, a, [(i, r["chanc"], f"{fd(r['fecha'])}\n{fn(r['chanc'], 0)} t")
                             for i, r in enumerate(dias) if r["chanc"]])
        filas = []
        for r in dias:
            h = r["h_ta"] + r["h_tb"]
            filas.append((fd(r["fecha"]), fn(r["mol_ta"], 0), fn(r["mol_tb"], 0), fn(r["mol_ta"] + r["mol_tb"], 0),
                          fn(r["h_ta"], 1), fn(r["h_tb"], 1), fn((r["mol_ta"] + r["mol_tb"]) / h if h else None, 1),
                          fn(r["chanc"], 0) if r["chanc"] is not None else "—"))
        pg.llenar(["Fecha", "Molido TA (t)", "Molido TB (t)", "Molido día (t)", "Horas TA", "Horas TB",
                   "Rate (t/h)", "Chancado (t)"], filas)
        pg.dibujar()

    # -- Leyes ----------------------------------------------------------------
    def pg_leyes(self, R):
        pg = self.p["leyes"]
        pg.limpiar()
        tv = self.turnos_vis(R)
        axs = pg.fig.subplots(3, 1, sharex=True)
        if not tv:
            for a in axs:
                vacio(a)
            pg.llenar([], [])
            pg.dibujar()
            return
        cfg = (("cab", "Ley de cabeza CuT (%)", C_CAB, R.plan["ley"], f"Plan {fn(R.plan['ley'], 2)} %", 2),
               ("conc", "Ley de concentrado final CuT (%)", C_CONC, R.plan["ley_conc"], f"Plan {fn(R.plan['ley_conc'], 1)} %", 1),
               ("rel", "Ley de relave CuT (%)", C_REL, None, "", 2))
        for a, (k, titulo, col, plan_, txt, dec) in zip(axs, cfg):
            ok = (lambda r: not (r["cab"] and r["rel"] is not None and r["rel"] >= r["cab"])) if k == "rel" else None
            ys = con_huecos(tv, k, ok)
            a.plot(range(len(tv)), ys, color=col, marker="o", ms=4.5, mfc=col, mec=SURFACE, mew=1)
            if plan_:
                linea_plan(a, plan_, txt)
            a.set_title(titulo)
            Hover(pg.canvas, a, [(i, v, f"{fd(tv[i]['fecha'])} {tv[i]['turno']}\n{fn(v, 2)} %")
                                 for i, v in enumerate(ys) if v == v])
        eje_turnos(axs[2], tv)
        anom = [r for r in tv if r["cab"] and r["rel"] is not None and r["rel"] >= r["cab"]]
        if anom:
            axs[2].annotate(f"{len(anom)} relave(s) anómalo(s) no graficado(s)", xy=(0.01, 0.92),
                            xycoords="axes fraction", fontsize=8, color=INK2)
        tt = filas_con_datos(tv, ("cab", "conc", "rel"))
        filas = [(fd(r["fecha"]), r["turno"], fn(r["molido"], 0), fn(r["cab"], 2), fn(r["cab_cus"], 2),
                  fn(r["conc"], 2), fn(r["rel"], 2)) for r in tt]
        colores = ["#fde8d8" if (r["cab"] and r["rel"] is not None and r["rel"] >= r["cab"]) else None for r in tt]
        pg.llenar(["Fecha", "Turno", "Molido (t)", "Cabeza CuT %", "Cabeza CuS %", "Conc. final CuT %", "Relave CuT %"],
                  filas, colores)
        pg.dibujar()

    # -- Recuperación -----------------------------------------------------------
    def pg_rec(self, R):
        pg = self.p["rec"]
        pg.limpiar()
        tv, dias = self.turnos_vis(R), self.dias_vis(R)
        a1, a2 = pg.fig.subplots(2, 1)
        ys1 = con_huecos(tv, "rec")
        if any(v == v for v in ys1):
            a1.plot(range(len(tv)), ys1, color=C_REC, marker="o", ms=4.5, mec=SURFACE, mew=1, lw=1.4)
            linea_plan(a1, R.plan["rec"], f"Plan {fn(R.plan['rec'], 0)} %", abajo=True)
            a1.set_ylim(60, 100)
            Hover(pg.canvas, a1, [(i, v, f"{fd(tv[i]['fecha'])} {tv[i]['turno']}\n{fn(v, 1)} %")
                                  for i, v in enumerate(ys1) if v == v])
            eje_turnos(a1, tv)
        else:
            vacio(a1)
        a1.set_title("Recuperación por leyes, por turno (%)")
        ys = [(i, r["rec_acum"]) for i, r in enumerate(dias) if r["rec_acum"]]
        if ys:
            a2.plot([i for i, _ in ys], [v for _, v in ys], color=C_REC, marker="o", ms=4, mec=SURFACE, mew=1)
            linea_plan(a2, R.plan["rec"], f"Plan {fn(R.plan['rec'], 0)} %", abajo=True)
            Hover(pg.canvas, a2, [(i, v, f"{fd(dias[i]['fecha'])}\nacumulada {fn(v, 1)} %") for i, v in ys])
            eje_dias(a2, dias)
        else:
            vacio(a2)
        a2.set_title("Recuperación acumulada, balance de masa (%)  ·  Cu fino filtrado / alimentado, desde 10 t de Cu alimentado")
        filas = []
        for r in filas_con_datos(tv, ("cab", "conc", "rel")):
            rr = r["rec"]
            nota = "—"
            if rr is None and r["cab"] and r["rel"] is not None and r["rel"] >= r["cab"]:
                nota = "relave anómalo"
            filas.append((fd(r["fecha"]), r["turno"], fn(r["molido"], 0), fn(r["cab"], 2), fn(r["conc"], 2),
                          fn(r["rel"], 2), fn(rr, 1) if rr else nota))
        pg.llenar(["Fecha", "Turno", "Molido (t)", "Cabeza %", "Conc. final %", "Relave %", "Rec. leyes %"], filas)
        pg.dibujar()

    # -- Filtrado ---------------------------------------------------------------
    def pg_filt(self, R):
        pg = self.p["filt"]
        pg.limpiar()
        dias = [r for r in self.dias_vis(R)]
        a1, a2 = pg.fig.subplots(2, 1)
        if not R.lotes:
            vacio(a1)
            vacio(a2)
            pg.llenar([], [])
            pg.dibujar()
            return
        x = list(range(len(dias)))
        sec = [r["dry"] for r in dias]
        agua = [max(r["wet"] - r["dry"], 0) for r in dias]
        a1.bar(x, sec, color=BLUE, width=0.72, label="Seco")
        a1.bar(x, agua, bottom=sec, color=BLUE_LIGHT, width=0.72, edgecolor=SURFACE, linewidth=1, label="Humedad")
        a1.set_title("Concentrado filtrado por día (t)")
        a1.legend(loc="upper left", ncols=2)
        eje_dias(a1, dias)
        Hover(pg.canvas, a1, [(i, r["wet"], f"{fd(r['fecha'])}\nhúmedo {fn(r['wet'], 2)} t\nseco {fn(r['dry'], 2)} t")
                              for i, r in enumerate(dias)])
        a2.bar(x, [r["cu_filt"] for r in dias], color=C_CONC, width=0.72)
        a2.set_title("Cu fino filtrado por día (t, solo lotes con ley)")
        eje_dias(a2, dias)
        Hover(pg.canvas, a2, [(i, r["cu_filt"], f"{fd(r['fecha'])}\n{fn(r['cu_filt'], 2)} t Cu") for i, r in enumerate(dias)])
        filas, colores = [], []
        for l in sorted(R.lotes, key=lambda z: (z["fecha"], z["turno"])):
            filas.append((l["met"], fd(l["fecha"]), l["turno"], fn(l["wet"] / 1000, 2), fn(l["hum"], 2),
                          fn(l["dry"] / 1000 if l["dry"] is not None else None, 2), fn(l["cut"], 2),
                          fn(l["cus"], 2), fn(l["ag"], 1)))
            colores.append("#fde8d8" if not l["cut"] else (None if l["cut"] >= R.plan["ley_conc"] else "#fff3cf"))
        pg.llenar(["Lote", "Fecha", "Turno", "Húmedo (t)", "Humedad %", "Seco (t)", "CuT %", "CuS %", "Ag g/t"],
                  filas, colores)
        pg.dibujar()

    # -- P80 --------------------------------------------------------------------
    def pg_p80(self, R):
        pg, d = self.p["p80"], self.datos
        pg.limpiar()
        axs = pg.fig.subplots(3, 1)
        filas = []
        ro = [p for p in d.p80_ro]
        cn = [p for p in d.p80_cn]

        def serie(a, lista, clave, color, titulo, objetivo=None, txt=""):
            pts = [(i, p[clave]) for i, p in enumerate(lista) if p[clave] is not None]
            if pts:
                a.plot([i for i, _ in pts], [v for _, v in pts], color=color, marker="o", ms=4.5, mec=SURFACE, mew=1)
                Hover(pg.canvas, a, [(i, v, f"{fd(lista[i]['fecha'])} {lista[i]['turno']}\n{fn(v, 1)}") for i, v in pts])
                if objetivo:
                    linea_plan(a, objetivo, txt)
                ticks = [i for i, p in enumerate(lista) if p["turno"] == "TA" and p["fecha"].day % 3 == 1]
                a.set_xticks(ticks)
                a.set_xticklabels([lista[i]["fecha"].strftime("%d-%m") for i in ticks])
            else:
                vacio(a)
            a.set_title(titulo)
        obj_ro = next((p["obj"] for p in ro if p["obj"]), 130)
        serie(axs[0], ro, "p80", BLUE, "P80 rougher (µm)", obj_ro, f"Objetivo {fn(obj_ro, 0)} µm")
        serie(axs[1], cn, "p80", ORANGE, "P80 concentrado (µm)")
        serie(axs[2], cn, "m100", ORANGE, "% pasante malla 100 en concentrado", 95, "Objetivo 95 %")
        for p in ro:
            filas.append(("Rougher", fd(p["fecha"]), p["turno"], p["p80_txt"] if p["p80"] is None else fn(p["p80"], 1),
                          fn(p["p50"], 1) if p["p50"] else "< 44", fn(p["m100"], 1), "—"))
        for p in cn:
            filas.append(("Concentrado", fd(p["fecha"]), p["turno"], p["p80_txt"] if p["p80"] is None else fn(p["p80"], 1),
                          fn(p["p50"], 1) if p["p50"] else "< 44", fn(p["m100"], 1), str(p["id"] or "—")))
        pg.llenar(["Muestra", "Fecha", "Turno", "P80 (µm)", "P50 (µm)", "% pas. M100", "ID"], filas)
        pg.dibujar()

    # -- Reactivos --------------------------------------------------------------
    def pg_reac(self, R):
        pg = self.p["reac"]
        pg.limpiar()
        tv = self.turnos_vis(R)
        axs = pg.fig.subplots(4, 2)
        cfg = (("mibc", "MIBC (mm/min)"), ("pac", "Xantato / PAC (mm/min)"), ("pq733", "PQ 733 (mm/min)"),
               ("pq510", "PQ 510 (mm/min)"), ("dowf", "Downfroath 250 (mm/min)"), ("nash", "NaSH (valor reportado)"),
               ("colchon", "Colchón de espuma re-cleaner (cm)"), ("m100", "% bajo malla 100 en molino 04"))
        for a, (k, titulo) in zip(axs.flat, cfg):
            ys = con_huecos(tv, k)
            if any(v == v for v in ys):
                a.plot(range(len(tv)), ys, color=BLUE, marker="o", ms=4, mec=SURFACE, mew=1)
                Hover(pg.canvas, a, [(i, v, f"{fd(tv[i]['fecha'])} {tv[i]['turno']}\n{fn(v, 1)}")
                                     for i, v in enumerate(ys) if v == v])
                eje_turnos(a, tv, 3)
            else:
                vacio(a, "Sin datos en los archivos cargados")
            a.set_title(titulo)
        claves = ("mibc", "pac", "pq733", "pq510", "dowf", "nash", "colchon", "m100")
        filas = [(fd(r["fecha"]), r["turno"]) + tuple(fn(r[k], 1) for k in claves)
                 for r in filas_con_datos(tv, claves)]
        pg.llenar(["Fecha", "Turno", "MIBC", "Xantato", "PQ 733", "PQ 510", "Downfroath", "NaSH", "Colchón", "% < M100"], filas)
        pg.dibujar()

    # -- Conciliación -----------------------------------------------------------
    def pg_conc(self, R):
        pg = self.p["conc"]
        c = R.cobertura
        self.lbl_cob.setText(
            "<b>Cobertura de datos</b><br>"
            f"Turnos con molino operando: {c['turnos_molino']} · con ley de cabeza: {c['con_ley']} · "
            f"con recuperación por leyes: {c['rec']}<br>"
            f"Lotes de concentrado: {c['lotes']} · con ley: {c['lotes_con_ley']}<br>"
            "<span style='color:#52514e'>El programa no estima leyes: lo que falta aparece como pendiente.</span>")
        col = {"alta": "#fde8d8", "media": "#fff3cf", "baja": None}
        filas = [({"alta": "Alta", "media": "Media", "baja": "Baja"}[g], tema, det) for g, tema, det in R.issues]
        pg.llenar(["Gravedad", "Tema", "Detalle"], filas, [col[g] for g, _, _ in R.issues])
        pg.tabla.setColumnWidth(0, 80)
        pg.tabla.setColumnWidth(1, 260)

    # -- Exportar ---------------------------------------------------------------
    def exportar(self):
        if not self.R:
            return
        ruta, _ = QFileDialog.getSaveFileName(self, "Guardar informe", "PSD_Reportabilidad.xlsx", "Excel (*.xlsx)")
        if not ruta:
            return
        from openpyxl import Workbook
        R, K = self.R, self.R.K
        wb = Workbook()
        ws = wb.active
        ws.title = "Resumen"
        pg = self.p["resumen"].tabla
        ws.append([pg.horizontalHeaderItem(j).text() for j in range(pg.columnCount())])
        for i in range(pg.rowCount()):
            ws.append([pg.item(i, j).text() for j in range(pg.columnCount())])
        for nombre, pagina in (("Operación", "oper"), ("Leyes", "leyes"), ("Recuperación", "rec"),
                               ("Filtrado lotes", "filt"), ("P80", "p80"), ("Reactivos", "reac"), ("Conciliación", "conc")):
            t = self.p[pagina].tabla
            if t.columnCount() == 0:
                continue
            w = wb.create_sheet(nombre)
            w.append([t.horizontalHeaderItem(j).text() for j in range(t.columnCount())])
            for i in range(t.rowCount()):
                w.append([t.item(i, j).text() for j in range(t.columnCount())])
        wb.save(ruta)
        QMessageBox.information(self, "Exportado", f"Informe guardado en:\n{ruta}")


def main(argv):
    app = QApplication(argv)
    app.setStyle("Fusion")
    f = app.font()
    f.setPointSize(10)
    app.setFont(f)
    v = Ventana()
    v.show()
    archivos = [a for a in argv[1:] if os.path.isfile(a)]
    if archivos:
        v.cargar(archivos)
    return app.exec()


if __name__ == "__main__":
    sys.exit(main(sys.argv))
