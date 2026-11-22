"""Núcleo de lectura y cálculo de PSD Analisis Corrientes.

Lee las planillas de la planta (Log de Muestras, Control Muestras, Reporte Jefe de Turno,
P80, Filtrado oficial) y calcula balance, leyes, recuperaciones y conciliación.
No inventa datos: lo que no está en las planillas queda vacío (None).
"""
import datetime as dt
import re
from collections import defaultdict
from openpyxl import load_workbook

TURNOS = ("TA", "TB")


# ----------------------------------------------------------------- utilidades
def num(v):
    if isinstance(v, bool) or v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, str):
        s = v.strip().replace(",", ".")
        try:
            return float(s)
        except ValueError:
            return None
    return None


def fecha(v):
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    if isinstance(v, str):
        try:
            return dt.date.fromisoformat(v.strip()[:10])
        except ValueError:
            return None
    return None


def turno(v):
    if v is None:
        return None
    s = str(v).strip().upper()
    if s in ("A", "TA"):
        return "TA"
    if s in ("B", "TB"):
        return "TB"
    return None


def fx(v, n=1):
    """Número con coma decimal y punto de miles."""
    return f"{v:,.{n}f}".replace(",", "X").replace(".", ",").replace("X", ".")


def fmt_fecha(d):
    return d.strftime("%d-%m-%y")


def tipo_archivo(wb):
    n = set(wb.sheetnames)
    if "Log Muestras" in n:
        return "log"
    if "Ingresar Datos" in n:
        return "control"
    if {"R(1)", "RESUMEN"} <= n:
        return "jefe"
    if {"Concentrado", "Rouger"} <= n:
        return "p80"
    if "CONC FILTRADO" in n:
        return "oficial"
    return None


NOMBRES = {
    "log": "Log de Muestras (leyes)",
    "control": "Control Muestras (asignación)",
    "jefe": "Reporte Jefe de Turno",
    "p80": "P80 granulometría",
    "oficial": "Filtrado oficial (V-S humedad)",
}


# ------------------------------------------------------------------- datos
RE_MET = re.compile(r"MET[-\s]*(\d+)")


class Datos:
    def __init__(self):
        self.turnos = defaultdict(dict)   # (fecha, 'TA'|'TB') -> campos
        self.dias = defaultdict(dict)     # fecha -> campos diarios (jefe de turno)
        self.lotes = []                   # lotes de concentrado (filtrado oficial + ley del Log)
        self.log = {}                     # N° Met -> {cut, cus, ag}
        self.rep_multiples = []
        self.asignacion = []              # Control Muestras: turno -> N° Met por corriente
        self.codigos_raros = []
        self.avisos_cons = []
        self.mes = None
        self.oficial = []                 # lotes oficiales (filtrado V-S humedad)
        self.p80_ro = []                  # P80 rougher
        self.p80_cn = []                  # P80 concentrado
        self.fuentes = {}                 # tipo -> nombre de archivo
        self.avisos = []

    # -- carga ---------------------------------------------------------------
    def cargar(self, ruta):
        wb = load_workbook(ruta, data_only=True, read_only=False)
        tipo = tipo_archivo(wb)
        if tipo is None:
            raise ValueError("No reconozco este archivo: hojas = " + ", ".join(wb.sheetnames[:6]))
        {"log": self._log, "control": self._control, "jefe": self._jefe,
         "p80": self._p80, "oficial": self._oficial}[tipo](wb)
        self.fuentes[tipo] = ruta.replace("\\", "/").split("/")[-1]
        return tipo

    # -- Reporte Jefe de Turno ---------------------------------------------
    def _jefe(self, wb):
        ws = wb["R(1)"]
        fechas_dia = {}
        for r in range(3, ws.max_row + 1):
            d = fecha(ws.cell(r, 1).value)
            if d is None:
                if r > 5:
                    break
                continue
            fechas_dia[d.day] = d
            dia = self.dias[d]
            dia["filt_ta"] = num(ws.cell(r, 2).value) or 0.0
            dia["filt_tb"] = num(ws.cell(r, 3).value) or 0.0
            dia["agua_ta"] = num(ws.cell(r, 27).value) or 0.0
            dia["agua_tb"] = num(ws.cell(r, 28).value) or 0.0
            for t, cm, ch in (("TA", 12, 20), ("TB", 13, 21)):
                self.turnos[(d, t)]["molido_jt"] = num(ws.cell(r, cm).value) or 0.0
                self.turnos[(d, t)]["horas_molino"] = num(ws.cell(r, ch).value) or 0.0
        self._jefe_hojas_diarias(wb)
        ws = wb["RESUMEN"]
        for r in range(5, 40):
            n = num(ws.cell(r, 1).value)
            if n is None or int(n) not in fechas_dia:
                continue
            dia = self.dias[fechas_dia[int(n)]]
            dia["chanc_ta"] = num(ws.cell(r, 2).value) or 0.0
            dia["chanc_tb"] = num(ws.cell(r, 3).value) or 0.0
            dia["chanc_h_ta"] = num(ws.cell(r, 5).value) or 0.0
            dia["chanc_h_tb"] = num(ws.cell(r, 6).value) or 0.0

    # -- Hojas diarias (1)…(30): reactivos y -#100 por turno (promedio de las lecturas) --
    ALIAS_REAC = (("dowf", ("downfroath", "dowfroath", "downfroth", "dowfroth")),
                  ("mibc", ("mibc",)),
                  ("pac", ("xantato", "pax", "pac ", "socol", "solcol")),
                  ("pq510", ("510",)),
                  ("pq733", ("733",)),
                  ("nash", ("nash", "nahs")),
                  ("colchon", ("colch",)),
                  ("m100_molino", ("-#100 molino 04",)))

    def _jefe_hojas_diarias(self, wb):
        for nombre in wb.sheetnames:
            if not re.fullmatch(r"\(\d+\)", nombre):
                continue
            ws = wb[nombre]
            d = None
            for r in range(1, 15):
                if str(ws.cell(r, 3).value or "").strip().lower() == "fecha":
                    d = fecha(ws.cell(r, 4).value)
                    break
            if d is None:
                continue
            turno_act, col_act = None, None
            for r in range(15, min(ws.max_row, 160) + 1):
                for c in (5, 7):
                    v = ws.cell(r, c).value
                    if isinstance(v, str) and v.strip() in ("TA", "TB"):
                        turno_act, col_act = v.strip(), c
                if str(ws.cell(r, 3).value or "").strip() == "Inicio":
                    turno_act = None
                if turno_act is None or col_act is None:
                    continue
                etiqueta = ws.cell(r, col_act).value
                if not isinstance(etiqueta, str):
                    continue
                e = etiqueta.strip().lower() + " "
                clave = next((k for k, als in self.ALIAS_REAC if any(a in e for a in als)), None)
                if clave is None:
                    continue
                vals = [num(ws.cell(r, c).value) for c in range(col_act + 2, col_act + 9)]
                vals = [v for v in vals if v is not None]
                if vals:
                    self.turnos[(d, turno_act)][clave] = sum(vals) / len(vals)

    # -- Log de Muestras: leyes por N° Met (la repetición "Rep" tiene prioridad) ----
    def _log(self, wb):
        ws = wb["Log Muestras"]
        simples, reps = {}, defaultdict(list)
        for row in ws.iter_rows(min_row=4, values_only=True):
            if len(row) < 42 or row[1] is None:
                continue
            m = RE_MET.search(str(row[1]).upper())
            if not m:
                continue
            n = int(m.group(1))
            v = dict(cut=num(row[39]), cus=num(row[40]), ag=num(row[41]))
            if "REP" in str(row[1]).upper():
                reps[n].append(v)
            else:
                simples[n] = v
        self.log = {}
        self.rep_multiples = []
        for n in set(simples) | set(reps):
            v = dict(simples.get(n, {}))
            if len(reps.get(n, [])) == 1:
                for k, val in reps[n][0].items():
                    if val is not None:
                        v[k] = val
            elif len(reps.get(n, [])) > 1:
                self.rep_multiples.append(n)      # regla: no elegir sola, se informa
            self.log[n] = v

    # -- Control de Muestras: qué Met corresponde a cada turno y corriente --------
    def _control(self, wb):
        ws = wb["Ingresar Datos"]
        self.asignacion = []
        self.codigos_raros = []
        for r in range(5, ws.max_row + 1):
            d, t = fecha(ws.cell(r, 2).value), turno(ws.cell(r, 3).value)
            if not d or not t:
                continue
            reg = dict(fecha=d, turno=t)
            for clave, col in (("cab", 4), ("conc", 9), ("rel", 13)):
                v = ws.cell(r, col).value
                if v in (None, "", "\xa0"):
                    reg[clave] = None
                    continue
                m = RE_MET.search(str(v).upper())
                if m:
                    reg[clave] = int(m.group(1))
                else:
                    reg[clave] = None
                    self.codigos_raros.append((d, t, clave, str(v)))
            self.asignacion.append(reg)

    def consolidar(self):
        """Junta Log + Control (+ filtrado oficial) en turnos y lotes. Idempotente."""
        self.avisos_cons = []
        fechas = [f for (f, t) in self.turnos if self.turnos[(f, t)].get("molido_jt")]
        base = max(fechas) if fechas else None
        if base is None and self.asignacion:
            base = max(a["fecha"] for a in self.asignacion if any(a[k] for k in ("cab", "conc", "rel")))
        self.mes = (base.year, base.month) if base else None
        enmes = lambda f: self.mes is None or (f.year, f.month) == self.mes
        for x in self.turnos.values():
            for k in ("cab_cut", "cab_cus", "conc_cut", "conc_cus", "rel_cut", "rel_cus",
                      "met_cab", "met_conc", "met_rel"):
                x.pop(k, None)
        for a in self.asignacion:
            if not enmes(a["fecha"]):
                continue
            x = self.turnos[(a["fecha"], a["turno"])]
            for clave, pre in (("cab", "cab"), ("conc", "conc"), ("rel", "rel")):
                n = a[clave]
                if n is None:
                    continue
                x["met_" + clave] = n
                v = self.log.get(n)
                if v is None or v.get("cut") is None:
                    self.avisos_cons.append((a["fecha"], a["turno"], clave, n))
                    continue
                x[pre + "_cut"] = v["cut"]
                x[pre + "_cus"] = v.get("cus")
        # lotes de concentrado: tonelaje oficial + ley del Log por N° Met
        self.usados = {a[k] for a in self.asignacion if enmes(a["fecha"]) for k in ("cab", "conc", "rel") if a[k]}
        self.raros_mes = [c for c in self.codigos_raros if enmes(c[0])]
        self.lotes = []
        for o in self.oficial:
            if not enmes(o["fecha"]):
                continue
            m = RE_MET.search(str(o["met"]).upper())
            n = int(m.group(1)) if m else None
            if n is not None:
                self.usados.add(n)
            v = self.log.get(n, {}) if n is not None else {}
            hum = o["hum"] * 100 if (o["hum"] is not None and o["hum"] <= 1) else o["hum"]
            dry = o["dry"] if o["dry"] else (o["wet"] * (1 - hum / 100) if hum is not None else None)
            self.lotes.append(dict(met=str(o["met"]).strip(), fecha=o["fecha"], turno=o["turno"],
                                   wet=o["wet"], hum=hum, dry=dry,
                                   cut=v.get("cut"), cus=v.get("cus"), ag=v.get("ag")))

    # -- P80 -----------------------------------------------------------------
    def _p80(self, wb):
        ws = wb["Concentrado"]
        for r in range(4, ws.max_row + 1):
            d, t = fecha(ws.cell(r, 2).value), turno(ws.cell(r, 3).value)
            p80 = ws.cell(r, 41).value
            if d and t and p80 not in (None, ""):
                self.p80_cn.append(dict(fecha=d, turno=t, id=ws.cell(r, 1).value,
                                        p80=num(p80), p80_txt=str(p80),
                                        p50=num(ws.cell(r, 42).value),
                                        m100=num(ws.cell(r, 43).value)))
        ws = wb["Rouger"]
        for r in range(4, ws.max_row + 1):
            d, t = fecha(ws.cell(r, 1).value), turno(ws.cell(r, 2).value)
            p80 = ws.cell(r, 40).value
            if d and t and p80 not in (None, ""):
                self.p80_ro.append(dict(fecha=d, turno=t,
                                        p80=num(p80), p80_txt=str(p80),
                                        p50=num(ws.cell(r, 41).value),
                                        m100=num(ws.cell(r, 42).value),
                                        obj=num(ws.cell(r, 43).value)))

    # -- Filtrado oficial ------------------------------------------------------
    def _oficial(self, wb):
        ws = wb["CONC FILTRADO"]
        for r in range(4, ws.max_row + 1):
            d, t = fecha(ws.cell(r, 2).value), turno(ws.cell(r, 3).value)
            wet = num(ws.cell(r, 6).value)
            if d and t and wet and wet > 0:
                self.oficial.append(dict(
                    fecha=d, turno=t, met=str(ws.cell(r, 5).value).strip(),
                    wet=wet, hum=num(ws.cell(r, 7).value), dry=num(ws.cell(r, 8).value)))


# ---------------------------------------------------------------- cálculo
PLAN_DEFECTO = dict(molienda=21300.0, ley=1.14, rec=77.0, ley_conc=23.0, dias=30)


def rec_dos_productos(cab, conc, rel):
    """Recuperación (%) por leyes, fórmula de dos productos. None si no es válida."""
    if None in (cab, conc, rel) or cab <= 0 or conc <= 0 or rel < 0:
        return None
    if not (conc > cab > rel):
        return None
    r = conc * (cab - rel) / (cab * (conc - rel)) * 100.0
    return r if 0 < r <= 100 else None


class Resultado:
    pass


def calcular(d, plan=None):
    plan = dict(PLAN_DEFECTO, **(plan or {}))
    R = Resultado()
    R.plan = plan
    R.issues = []  # (gravedad, tema, detalle)
    usa_jefe = "jefe" in d.fuentes
    d.consolidar()

    # --- turnos
    filas = []
    for (f, t), x in sorted(d.turnos.items()):
        molido = x.get("molido_jt")
        if d.mes and (f.year, f.month) != d.mes:
            continue
        cab, conc, rel = x.get("cab_cut"), x.get("conc_cut"), x.get("rel_cut")
        rec = rec_dos_productos(cab, conc, rel)
        if cab and rel is not None and rel >= cab:
            R.issues.append(("alta", "Relave anómalo",
                             f"{fmt_fecha(f)} {t}: relave {fx(rel, 2)} % ≥ cabeza {fx(cab, 2)} % (revisar códigos Met imputados)"))
        if (molido or 0) > 0 and not cab and "log" in d.fuentes:
            R.issues.append(("media", "Molienda sin ley de cabeza",
                             f"{fmt_fecha(f)} {t}: {fx(molido, 0)} t molidas sin ley de cabeza"))
        filas.append(dict(fecha=f, turno=t, molido=molido or 0.0,
                          horas=x.get("horas_molino"), cab=cab, cab_cus=x.get("cab_cus"),
                          conc=conc, rel=rel, rec=rec,
                          m100=x.get("m100_molino"), mibc=x.get("mibc"), pac=x.get("pac"),
                          pq733=x.get("pq733"), pq510=x.get("pq510"), dowf=x.get("dowf"),
                          nash=x.get("nash"), colchon=x.get("colchon")))
    R.turnos = filas
    for f, t, clave, n in d.avisos_cons:
        nom = {"cab": "cabeza", "conc": "concentrado", "rel": "relave"}[clave]
        R.issues.append(("alta", "Muestra sin ley en el Log",
                         f"{fmt_fecha(f)} {t}: Met-{n} ({nom}) asignada en Control Muestras pero sin CuT en el Log"))
    for f, t, clave, v in d.raros_mes:
        R.issues.append(("media", "Código no reconocido en Control Muestras",
                         f"{fmt_fecha(f)} {t} ({clave}): «{v}»"))
    rep_m = sorted(n for n in d.rep_multiples if n in d.usados)
    if rep_m:
        R.issues.append(("alta", "Más de una repetición (Rep) para la misma muestra",
                         "Se usó la ley simple; defina cuál Rep corresponde: "
                         + ", ".join(f"Met-{n}" for n in rep_m)))

    # --- lotes de concentrado
    lotes = d.lotes
    R.lotes = lotes
    wet = sum(l["wet"] for l in lotes) / 1000.0
    dry = sum(l["dry"] for l in lotes if l["dry"] is not None) / 1000.0
    con_ley = [l for l in lotes if l["cut"] and l["dry"] is not None]
    sin_ley = [l for l in lotes if not l["cut"]]
    cu_filt = sum(l["dry"] * l["cut"] / 100.0 for l in con_ley) / 1000.0
    dry_con_ley = sum(l["dry"] for l in con_ley) / 1000.0
    R.sin_ley = sin_ley
    if sin_ley:
        t_sin = sum(l["dry"] or l["wet"] for l in sin_ley) / 1000.0
        R.issues.append(("alta", "Lotes sin ley",
                         f"{len(sin_ley)} lotes ({fx(t_sin)} t secas, {fx(t_sin / dry * 100)} % del filtrado) sin ley en el Log: "
                         + ", ".join(l['met'] for l in sin_ley)))

    # --- alimentación
    con_cab = [r for r in filas if r["molido"] > 0 and r["cab"]]
    molido_tot = sum(r["molido"] for r in filas)
    molido_cab = sum(r["molido"] for r in con_cab)
    cu_alim = sum(r["molido"] * r["cab"] / 100.0 for r in con_cab)
    con_cus = [r for r in con_cab if r["cab_cus"] is not None]
    rec_ok = [r for r in filas if r["rec"] and r["molido"] > 0]

    K = {}
    K["molido"] = molido_tot
    K["molido_con_ley"] = molido_cab
    K["horas"] = sum(r["horas"] or 0 for r in filas)
    K["rate"] = molido_tot / K["horas"] if K["horas"] else None
    K["cab"] = cu_alim / molido_cab * 100 if molido_cab else None
    K["cab_cus"] = (sum(r["molido"] * r["cab_cus"] for r in con_cus) / sum(r["molido"] for r in con_cus)
                    if con_cus else None)
    K["cu_alim"] = cu_alim
    K["wet"], K["dry"] = wet, dry
    K["hum"] = (1 - dry / wet) * 100 if wet else None
    K["conc"] = (sum(l["dry"] * l["cut"] for l in con_ley) / 1000.0 / dry_con_ley) if dry_con_ley else None
    K["cu_filt"] = cu_filt
    K["dry_sin_ley"] = dry - dry_con_ley
    K["rec_bm"] = cu_filt / cu_alim * 100 if cu_alim else None
    K["rec_leyes"] = (sum(r["rec"] * r["molido"] for r in rec_ok) / sum(r["molido"] for r in rec_ok)
                      if rec_ok else None)
    K["n_rec"] = len(rec_ok)
    K["n_turnos_molino"] = sum(1 for r in filas if r["molido"] > 0)
    K["n_con_ley"] = len(con_cab)
    K["lotes_bajo_obj"] = sum(1 for l in con_ley if l["cut"] < plan["ley_conc"])
    K["lotes_con_ley"] = len(con_ley)

    # --- plan
    cu_plan = plan["molienda"] * plan["ley"] / 100 * plan["rec"] / 100
    K["cu_plan"] = cu_plan
    K["dmt_plan"] = cu_plan / (plan["ley_conc"] / 100)
    fechas = [f for f in {r["fecha"] for r in filas if r["molido"] > 0}] + \
             [l["fecha"] for l in lotes]
    K["ultimo_dia"] = max(fechas) if fechas else None
    K["dias_transc"] = K["ultimo_dia"].day if K["ultimo_dia"] else 0
    R.K = K

    # --- series diarias
    dias = sorted({r["fecha"] for r in filas} | {l["fecha"] for l in lotes} | set(d.dias))
    D = []
    cum_cu_alim = cum_cu_filt = 0.0
    for f in dias:
        ta = next((r for r in filas if r["fecha"] == f and r["turno"] == "TA"), None)
        tb = next((r for r in filas if r["fecha"] == f and r["turno"] == "TB"), None)
        cu_a = sum(r["molido"] * r["cab"] / 100 for r in (ta, tb) if r and r["molido"] > 0 and r["cab"])
        ls = [l for l in lotes if l["fecha"] == f]
        cu_f = sum(l["dry"] * l["cut"] / 100000.0 for l in ls if l["cut"] and l["dry"] is not None)
        cum_cu_alim += cu_a
        cum_cu_filt += cu_f
        dj = d.dias.get(f, {})
        D.append(dict(
            fecha=f,
            mol_ta=ta["molido"] if ta else 0.0, mol_tb=tb["molido"] if tb else 0.0,
            h_ta=(ta or {}).get("horas") or 0.0, h_tb=(tb or {}).get("horas") or 0.0,
            chanc=(dj.get("chanc_ta", 0) or 0) + (dj.get("chanc_tb", 0) or 0) if dj else None,
            wet=sum(l["wet"] for l in ls) / 1000.0,
            dry=sum(l["dry"] for l in ls if l["dry"] is not None) / 1000.0,
            wet_jt=(dj.get("filt_ta", 0) + dj.get("filt_tb", 0)) if "filt_ta" in dj else None,
            cu_filt=cu_f, cu_alim=cu_a,
            rec_acum=(cum_cu_filt / cum_cu_alim * 100) if cum_cu_alim >= 10 else None))
    R.dias = D

    # --- conciliación Jefe de Turno vs lotes (húmedo por día)
    if usa_jefe and lotes:
        for r in D:
            if r["wet_jt"] is not None and abs(r["wet_jt"] - r["wet"]) > 0.5:
                R.issues.append(("baja", "Filtrado: Jefe de Turno ≠ oficial (se usa el oficial)",
                                 f"{fmt_fecha(r['fecha'])}: Jefe de Turno {fx(r['wet_jt'], 2)} t vs Filtrado V-S humedad {fx(r['wet'], 2)} t húmedas"))
    # --- P80
    for nombre, lista in (("rougher", d.p80_ro), ("concentrado", d.p80_cn)):
        fuera = [p for p in lista if p["p80"] is None]
        if fuera:
            R.issues.append(("baja", f"P80 {nombre} fuera de rango",
                             f"{len(fuera)} de {len(lista)} muestras con P80 fuera de mallas (< 44 o > 278 µm): no se grafican"))
    # --- cobertura
    R.cobertura = dict(
        turnos_molino=K["n_turnos_molino"], con_ley=K["n_con_ley"],
        rec=K["n_rec"], lotes=len(lotes), lotes_con_ley=len(con_ley))
    orden = {"alta": 0, "media": 1, "baja": 2}
    R.issues.sort(key=lambda x: orden[x[0]])
    return R
