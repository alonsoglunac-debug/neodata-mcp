# -*- coding: utf-8 -*-
"""
Neodata MCP — servidor MCP para Neodata Precios Unitarios (PU Win+).

Neodata guarda cada obra como una base SQL Server (.mdf) servida por
SQL Server LocalDB, instancia (localdb)\\MSSQLLocalDB. Este servidor se
conecta con autenticación de Windows (no usa la contraseña 'sa' de Neodata).

Lectura: presupuesto, partidas, conceptos, matrices (APU), insumos,
explosión de insumos, SQL de solo lectura.
Archivos: exporta a Excel; plantilla e importación desde Excel.
Escritura: usa los procedimientos almacenados de Neodata
(NeoPartidas_GrabarActualizar, NeoConceptos_GrabarActualizar,
NeoExpIns_GrabarActualizar, NeoMatrices_GrabarSoloMatriz), con:
  * dry_run=True por omisión (se ejecuta en una transacción y se revierte),
  * respaldo .bak automático antes de cada escritura real,
  * bloqueo si Neodata está abierto,
  * marca RequiereCalcular/RequiereTotalizar para que Neodata recalcule precios.
"""
from __future__ import annotations

import datetime as dt
import glob
import os
import re
import subprocess
from collections import defaultdict
from decimal import Decimal
from typing import Any, Optional

import pyodbc
from mcp.server.fastmcp import FastMCP

# ----------------------------------------------------------------------------
# Configuración
# ----------------------------------------------------------------------------
SERVIDOR = os.environ.get("NEODATA_SERVIDOR", r"(localdb)\MSSQLLocalDB")
DRIVER = os.environ.get("NEODATA_DRIVER", "ODBC Driver 17 for SQL Server")
CARPETAS_OBRAS = [p for p in os.environ.get(
    "NEODATA_CARPETAS",
    r"C:\NeodataWinPlusObras;C:\Neodata2021Obras;C:\Neodata2018Obras").split(";") if p]
CARPETA_RESPALDOS = os.environ.get("NEODATA_RESPALDOS", r"C:\NeodataWinPlusObras\respaldos_mcp")
CARPETA_SALIDA = os.environ.get(
    "NEODATA_SALIDA", os.path.join(os.path.expanduser("~"), "Documents", "Neodata_MCP"))
PROCESOS_NEODATA = ("puwinplus.exe", "puwinplusa.exe")
MONEDA = 1          # PuMonedas: 1 = PESOS
ID_USUARIO = 1      # SegUsuarios 'neodata'
MAX_RESPALDOS_POR_OBRA = 30

TIPOS = {1: "MATERIALES", 2: "MANO DE OBRA", 3: "EQUIPO Y HERRAMIENTA", 4: "BASICOS", 5: "SUBCONTRATOS"}
ALIAS_TIPO = {
    "material": 1, "materiales": 1, "mat": 1,
    "mano de obra": 2, "mo": 2, "mano_obra": 2,
    "equipo": 3, "herramienta": 3, "equipo y herramienta": 3, "eq": 3,
    "basico": 4, "basicos": 4, "básico": 4, "básicos": 4, "concepto": 4, "matriz": 4,
    "subcontrato": 5, "subcontratos": 5, "sub": 5,
}
ALIAS_UNIDAD = {"m2": "m²", "m^2": "m²", "m3": "m³", "m^3": "m³", "pza": "PZA", "pieza": "PZA",
                "lt": "LT", "litro": "LT", "hr": "HR", "hora": "HR", "jgo": "JGO", "lote": "LOTE"}

mcp = FastMCP(
    "neodata",
    instructions=(
        "Neodata Precios Unitarios (PU Win+) vía SQL Server LocalDB. "
        "Empieza con neodata_estado o neodata_listar_obras. Construbase (catálogo de precios con 4 maestros: "
        "CONSTRUCTOR, CONCURSOS, REMODELADOR, DESARROLLADOR) se consulta como obra='Construbase_Octubre_2026' "
        "con presupuesto=<maestro>; la primera consulta a una base grande puede tardar 3-4 min (apertura en frío), "
        "las siguientes ~2 s. Cada herramienta acepta 'obra' "
        "(nombre o ruta del .mdf; si se omite se usa la obra única/más reciente). "
        "Costos = costo directo calculado de las matrices; 'precio' solo existe después de "
        "Calcular en Neodata. Las escrituras son dry_run=True por omisión: muestra el resultado "
        "al usuario y repite con dry_run=False solo con su aprobación. Neodata debe estar cerrado "
        "para escribir. Unidades métricas (m, m², m³)."),
)


# ----------------------------------------------------------------------------
# Utilidades de conexión
# ----------------------------------------------------------------------------
def _num(v: Any) -> Any:
    if isinstance(v, Decimal):
        return float(v)
    if isinstance(v, (dt.datetime, dt.date)):
        return v.isoformat(sep=" ") if isinstance(v, dt.datetime) else v.isoformat()
    return v


def _conectar(base: str = "master", autocommit: bool = True) -> pyodbc.Connection:
    cs = f"DRIVER={{{DRIVER}}};SERVER={SERVIDOR};Trusted_Connection=yes;DATABASE={{{base}}}"
    return pyodbc.connect(cs, timeout=60, autocommit=autocommit)


def _filas(cur: pyodbc.Cursor) -> list[dict]:
    if not cur.description:
        return []
    cols = [c[0] for c in cur.description]
    return [{c: _num(v) for c, v in zip(cols, r)} for r in cur.fetchall()]


def _ultimo_resultado(cur: pyodbc.Cursor) -> list[dict]:
    """Recorre todos los result sets de un lote y devuelve el último no vacío."""
    ultimo: list[dict] = []
    while True:
        if cur.description:
            ultimo = _filas(cur)
        if not cur.nextset():
            break
    return ultimo


def neodata_abierto() -> bool:
    try:
        out = subprocess.run(["tasklist", "/FO", "CSV", "/NH"], capture_output=True, text=True,
                             timeout=20, creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0)).stdout.lower()
    except Exception:
        return False
    return any(p in out for p in PROCESOS_NEODATA)


def _bases_adjuntas() -> list[dict]:
    with _conectar() as cn:
        cur = cn.execute(
            "SELECT d.name AS base, mf.physical_name AS mdf, d.state_desc AS estado "
            "FROM sys.databases d JOIN sys.master_files mf ON mf.database_id=d.database_id AND mf.type=0 "
            "WHERE d.database_id>4")
        return _filas(cur)


def _mdfs_en_disco() -> list[str]:
    res = []
    for c in CARPETAS_OBRAS:
        res += glob.glob(os.path.join(c, "*.mdf"))
    return sorted(set(os.path.normcase(os.path.abspath(p)) for p in res))


def _adjuntar(mdf: str) -> str:
    mdf = os.path.abspath(mdf)
    if not os.path.isfile(mdf):
        raise ValueError(f"No existe el archivo {mdf}")
    nombre = mdf.upper()  # Neodata usa la ruta en mayúsculas como nombre de la base
    base, _ = os.path.splitext(mdf)
    ldf = next((p for p in (base + "_log.ldf", base + "_LOG.ldf", base + ".ldf") if os.path.isfile(p)), None)
    q = nombre.replace("]", "]]")
    lit = lambda s: "N'" + s.replace("'", "''") + "'"  # CREATE DATABASE no admite parámetros en FILENAME
    with _conectar() as cn:
        if ldf:
            cur = cn.execute(f"CREATE DATABASE [{q}] ON (FILENAME={lit(mdf)}), (FILENAME={lit(ldf)}) FOR ATTACH")
        else:
            cur = cn.execute(f"CREATE DATABASE [{q}] ON (FILENAME={lit(mdf)}) FOR ATTACH_REBUILD_LOG")
        while cur.nextset():
            pass
    return nombre


_ANCLAS: dict[str, pyodbc.Connection] = {}


def _anclar(base: str) -> str:
    """Mantiene una conexión abierta a la base mientras viva el servidor.
    Neodata adjunta sus bases con AUTO_CLOSE=ON: sin conexiones, LocalDB la cierra y la siguiente
    consulta la reabre en frío (en Construbase, ~850 MB, eso tarda 3-4 minutos)."""
    cn = _ANCLAS.get(base)
    if cn is not None:
        try:
            cn.execute("SELECT 1").fetchone()
            return base
        except Exception:
            _ANCLAS.pop(base, None)
    try:
        _ANCLAS[base] = _conectar(base)
    except Exception:
        pass
    return base


def _soltar_ancla(base: str) -> None:
    cn = _ANCLAS.pop(base, None)
    if cn is not None:
        try:
            cn.close()
        except Exception:
            pass


def _resolver_obra(obra: Optional[str]) -> str:
    """Devuelve el nombre de la base en LocalDB (adjuntándola si hace falta) y la mantiene abierta."""
    return _anclar(_resolver_obra_sin_ancla(obra))


def _resolver_obra_sin_ancla(obra: Optional[str]) -> str:
    adj = _bases_adjuntas()
    if not obra:
        # Construbase es catálogo de precios, nunca la obra por omisión
        adj = [a for a in adj if "construbase" not in a["mdf"].lower()]
        if len(adj) == 1:
            return adj[0]["base"]
        disco = [p for p in _mdfs_en_disco() if "construbase" not in p.lower() and "prueba mcp" not in p.lower()]
        if not disco and not adj:
            raise ValueError("No encontré obras de Neodata. Indica 'obra' con la ruta del .mdf.")
        if disco:
            reciente = max(disco, key=os.path.getmtime)
            obra = reciente
        else:
            raise ValueError("Hay varias obras adjuntas; indica 'obra'. Usa neodata_listar_obras.")
    o = obra.strip().strip('"')
    on = os.path.normcase(os.path.abspath(o)) if (":" in o or "\\" in o or "/" in o) else None
    # 1) ya adjunta
    for a in adj:
        mdfn = os.path.normcase(a["mdf"])
        stem = os.path.splitext(os.path.basename(a["mdf"]))[0].lower()
        if (on and mdfn == on) or a["base"].lower() == o.lower() or stem == o.lower():
            return a["base"]
    cand = [a for a in adj if o.lower() in os.path.basename(a["mdf"]).lower()]
    if len(cand) == 1:
        return cand[0]["base"]
    # 2) en disco
    if on and os.path.isfile(on):
        return _adjuntar(on)
    disco = [p for p in _mdfs_en_disco() if o.lower() in os.path.basename(p).lower()]
    if len(disco) == 1:
        return _adjuntar(disco[0])
    if len(disco) > 1 or len(cand) > 1:
        raise ValueError(f"'{obra}' coincide con varias obras: {[os.path.basename(p) for p in disco] or [c['mdf'] for c in cand]}")
    raise ValueError(f"No encontré la obra '{obra}'. Usa neodata_listar_obras.")


def _es_obra_pu(cn: pyodbc.Connection) -> bool:
    return cn.execute("SELECT OBJECT_ID('dbo.PuPresupuestos')").fetchone()[0] is not None


def _presupuesto(cn: pyodbc.Connection, presupuesto: Optional[str]) -> dict:
    if not _es_obra_pu(cn):
        raise ValueError("Esta base no tiene el esquema de presupuestos de Neodata PU "
                         "(¿es un catálogo tipo Construbase?). Usa neodata_consulta_sql.")
    rows = _filas(cn.execute(
        "SELECT IdPresupuesto, Presupuesto, Nombre, PorcentajeIVA, Bloqueado, Autorizado, Cerrado, "
        "RequiereCalcular, RequiereTotalizar, EsMaestro, FechaUltimoUso FROM PuPresupuestos ORDER BY IdPresupuesto"))
    if not rows:
        raise ValueError("La obra no tiene presupuestos.")
    if presupuesto is None:
        no_maestro = [r for r in rows if not r["EsMaestro"]]
        return (no_maestro or rows)[0]
    p = str(presupuesto).strip().lower()
    for r in rows:
        if str(r["IdPresupuesto"]) == p or (r["Presupuesto"] or "").lower() == p:
            return r
    raise ValueError(f"No existe el presupuesto '{presupuesto}'. Disponibles: {[r['Presupuesto'] for r in rows]}")


def _decimales(cn: pyodbc.Connection, idp: int) -> int:
    r = cn.execute("SELECT Valor FROM PuPresupuestosParametros WHERE IdPresupuesto=? AND Parametro='DecimalesRedondeo'",
                   idp).fetchone() or cn.execute("SELECT Valor FROM PuParametros WHERE Parametro='DecimalesRedondeo'").fetchone()
    try:
        return int(r[0])
    except Exception:
        return 2


# ----------------------------------------------------------------------------
# Modelo de costos (cálculo de costo directo desde las matrices)
# ----------------------------------------------------------------------------
class Modelo:
    def __init__(self, cn: pyodbc.Connection, idp: int):
        self.idp = idp
        self.dec = _decimales(cn, idp)
        self.ei: dict[int, dict] = {}
        q = ("SELECT EI.IdExpIns, EI.IdCodigo, C.Codigo, C.Descripcion, C.DescripcionLarga, U.Unidad, C.IdTipo, "
             "C.EsPorcentaje, COALESCE(EIC.Costo,0) AS CostoBD, EI.IdProveedor, EI.InsumoIntegrado, EI.EsCostoHorario, EI.IdIntelimat "
             "FROM PuExpIns EI JOIN PuCatalogo C ON C.IdCodigo=EI.IdCodigo JOIN PuUnidades U ON U.IdUnidad=C.IdUnidad "
             "LEFT JOIN PuExpInsCostos EIC ON EIC.IdExpIns=EI.IdExpIns AND EIC.IdMoneda=? WHERE EI.IdPresupuesto=?")
        for r in _filas(cn.execute(q, MONEDA, idp)):
            self.ei[r["IdExpIns"]] = r
        self.mat: dict[int, list[dict]] = defaultdict(list)
        q = ("SELECT M.IdMatriz, M.IdCodigoMatriz, M.Renglon, M.IdCodigoInsumo, M.Dividir, M.Volumen, M.Expresion "
             "FROM PuMatrices M JOIN PuExpIns EI ON EI.IdExpIns=M.IdCodigoMatriz WHERE EI.IdPresupuesto=? ORDER BY M.IdCodigoMatriz, M.Renglon")
        for r in _filas(cn.execute(q, idp)):
            self.mat[r["IdCodigoMatriz"]].append(r)
        self.por_codigo = {v["Codigo"].upper(): k for k, v in self.ei.items()}
        self._memo: dict[int, float] = {}

    def tiene_matriz(self, ide: int) -> bool:
        return bool(self.mat.get(ide))

    def lineas(self, ide: int, _pila: tuple = ()) -> list[dict]:
        out = []
        for l in self.mat.get(ide, []):
            ins = self.ei.get(l["IdCodigoInsumo"], {})
            vol = l["Volumen"] or 0.0
            d = {"renglon": l["Renglon"], "id_matriz": l["IdMatriz"], "id_expins": l["IdCodigoInsumo"],
                 "codigo": ins.get("Codigo"), "descripcion": ins.get("DescripcionLarga") or ins.get("Descripcion"),
                 "unidad": ins.get("Unidad"), "tipo": TIPOS.get(ins.get("IdTipo"), ins.get("IdTipo")),
                 "id_tipo": ins.get("IdTipo"), "cantidad": vol, "dividir": bool(l["Dividir"]),
                 "es_porcentaje": int(ins.get("EsPorcentaje") or 0), "tiene_matriz": self.tiene_matriz(l["IdCodigoInsumo"])}
            if d["es_porcentaje"]:
                d["costo"] = None
            else:
                c = self.costo(l["IdCodigoInsumo"], _pila + (ide,))
                d["costo"] = c
                imp = (c / vol if vol else 0.0) if d["dividir"] else c * vol
                d["importe"] = round(imp, self.dec)
            out.append(d)
        # Insumos porcentuales (p. ej. herramienta menor = 3% de la mano de obra):
        # su costo es la suma de importes del tipo al que hacen referencia.
        for d in out:
            if d["es_porcentaje"]:
                base = sum(x.get("importe", 0) for x in out if not x["es_porcentaje"] and x["id_tipo"] == d["es_porcentaje"])
                d["costo"] = round(base, self.dec)
                d["importe"] = round(base * d["cantidad"], self.dec)
                d["porcentaje_de"] = TIPOS.get(d["es_porcentaje"], d["es_porcentaje"])
        return out

    def costo(self, ide: int, _pila: tuple = ()) -> float:
        if ide in self._memo:
            return self._memo[ide]
        e = self.ei.get(ide)
        if e is None:
            return 0.0
        if not self.mat.get(ide) or ide in _pila:
            c = float(e["CostoBD"] or 0)
        else:
            c = round(sum(l["importe"] for l in self.lineas(ide, _pila)), self.dec)
        self._memo[ide] = c
        return c

    def explotar(self, ide: int, cant: float, acum: dict, _pila: tuple = ()):
        """Explosión de insumos: acumula cantidad e importe de insumos hoja."""
        if not self.mat.get(ide) or ide in _pila:
            e = self.ei.get(ide, {})
            a = acum.setdefault(ide, {"codigo": e.get("Codigo"), "descripcion": e.get("DescripcionLarga") or e.get("Descripcion"),
                                      "unidad": e.get("Unidad"), "tipo": TIPOS.get(e.get("IdTipo"), e.get("IdTipo")),
                                      "costo": float(e.get("CostoBD") or 0), "cantidad": 0.0, "importe": 0.0})
            a["cantidad"] += cant
            a["importe"] += cant * a["costo"]
            return
        for l in self.lineas(ide, _pila):
            if l["es_porcentaje"]:
                e = self.ei.get(l["id_expins"], {})
                a = acum.setdefault(l["id_expins"], {"codigo": l["codigo"], "descripcion": l["descripcion"], "unidad": l["unidad"],
                                                     "tipo": l["tipo"], "costo": None, "cantidad": 0.0, "importe": 0.0,
                                                     "nota": f"porcentaje de {l.get('porcentaje_de')}"})
                a["cantidad"] += cant * l["cantidad"]
                a["importe"] += cant * l["importe"]
                continue
            vol = l["cantidad"]
            sub = (cant / vol if vol else 0.0) if l["dividir"] else cant * vol
            self.explotar(l["id_expins"], sub, acum, _pila + (ide,))


def _partidas(cn: pyodbc.Connection, idp: int) -> list[dict]:
    rows = _filas(cn.execute(
        "SELECT P.IdPresupuestoPartida, P.IdPartidaPadre, P.Partida, P.PartidaWBS, P.Renglon, P.Cantidad, "
        "P.DescripcionPartida, P.DescripcionPartidaLarga, PC.Costo, PC.Precio, PC.CostoTotal, PC.PrecioTotal "
        "FROM PuPresupuestosPartidas P LEFT JOIN PuPresupuestosPartidasCostos PC "
        "ON PC.IdPresupuestoPartida=P.IdPresupuestoPartida AND PC.IdMoneda=? WHERE P.IdPresupuesto=?", MONEDA, idp))
    por_id = {r["IdPresupuestoPartida"]: r for r in rows}
    hijos = defaultdict(list)
    for r in rows:
        hijos[r["IdPartidaPadre"]].append(r)
    orden = []

    def rec(padre, nivel, factor, ruta):
        for r in sorted(hijos.get(padre, []), key=lambda x: (x["Renglon"] or 0, x["Partida"] or "")):
            r["nivel"] = nivel
            r["factor"] = factor * float(r["Cantidad"] or 1)
            r["ruta"] = ruta + [r["Partida"]]
            orden.append(r)
            rec(r["IdPresupuestoPartida"], nivel + 1, r["factor"], r["ruta"])
    rec(None, 0, 1.0, [])
    for r in rows:  # huérfanas (por si acaso)
        if r not in orden:
            r.update(nivel=0, factor=float(r["Cantidad"] or 1), ruta=[r["Partida"]])
            orden.append(r)
    return orden


def _conceptos(cn: pyodbc.Connection, idp: int) -> list[dict]:
    return _filas(cn.execute(
        "SELECT PC.IdPresupuestoConcepto, PC.IdPresupuestoPartida, PC.IdExpIns, PC.Renglon, PC.Cantidad, PC.CantidadTotal, "
        "C.Codigo, C.Descripcion, C.DescripcionLarga, U.Unidad, C.IdTipo, PCP.Precio "
        "FROM PuPresupuestosConceptos PC JOIN PuExpIns EI ON EI.IdExpIns=PC.IdExpIns "
        "JOIN PuCatalogo C ON C.IdCodigo=EI.IdCodigo JOIN PuUnidades U ON U.IdUnidad=C.IdUnidad "
        "LEFT JOIN PuPresupuestosConceptosPrecios PCP ON PCP.IdPresupuestoConcepto=PC.IdPresupuestoConcepto AND PCP.IdMoneda=? "
        "WHERE PC.IdPresupuesto=? ORDER BY PC.IdPresupuestoPartida, PC.Renglon", MONEDA, idp))


def _arbol(cn, idp, factor_sobrecosto: Optional[float] = None):
    """Partidas + conceptos con costo directo calculado e importes acumulados."""
    m = Modelo(cn, idp)
    parts = _partidas(cn, idp)
    por_id = {p["IdPresupuestoPartida"]: p for p in parts}
    conc = _conceptos(cn, idp)
    for p in parts:
        p["conceptos"] = []
        p["importe_cd"] = 0.0
        p["importe_precio_neodata"] = 0.0
    for c in conc:
        p = por_id.get(c["IdPresupuestoPartida"])
        factor = p["factor"] if p else 1.0
        c["cantidad_total"] = round(float(c["Cantidad"] or 0) * factor, 6)
        c["costo_directo"] = m.costo(c["IdExpIns"])
        c["tiene_matriz"] = m.tiene_matriz(c["IdExpIns"])
        c["importe_cd"] = round(c["cantidad_total"] * c["costo_directo"], m.dec)
        if factor_sobrecosto:
            c["precio_estimado"] = round(c["costo_directo"] * factor_sobrecosto, m.dec)
            c["importe_estimado"] = round(c["cantidad_total"] * c["precio_estimado"], m.dec)
        if p:
            p["conceptos"].append(c)
    # acumula importes hacia arriba
    for p in reversed(parts):
        p["importe_cd"] += sum(c["importe_cd"] for c in p["conceptos"])
        p["importe_precio_neodata"] += sum((c["Precio"] or 0) * c["cantidad_total"] for c in p["conceptos"])
        padre = por_id.get(p["IdPartidaPadre"])
        if padre:
            padre["importe_cd"] += p["importe_cd"]
            padre["importe_precio_neodata"] += p["importe_precio_neodata"]
    return m, parts, conc


# ----------------------------------------------------------------------------
# Herramientas de lectura
# ----------------------------------------------------------------------------
@mcp.tool()
def neodata_estado() -> dict:
    """Estado del entorno: LocalDB, si Neodata está abierto, obras adjuntas y archivos .mdf encontrados."""
    info: dict[str, Any] = {"servidor": SERVIDOR, "neodata_abierto": neodata_abierto(),
                            "carpetas_obras": CARPETAS_OBRAS, "carpeta_respaldos": CARPETA_RESPALDOS,
                            "carpeta_salida": CARPETA_SALIDA}
    try:
        with _conectar() as cn:
            info["sql_version"] = cn.execute("SELECT @@VERSION").fetchone()[0].splitlines()[0]
        info["obras_adjuntas"] = _bases_adjuntas()
    except Exception as e:
        info["error_localdb"] = str(e)
    info["mdf_en_disco"] = _mdfs_en_disco()
    return info


@mcp.tool()
def neodata_listar_obras() -> list[dict]:
    """Lista las obras (.mdf) de Neodata: archivo, tamaño, fecha, si está adjunta y sus presupuestos."""
    adj = {os.path.normcase(a["mdf"]): a for a in _bases_adjuntas()}
    rutas = set(_mdfs_en_disco()) | set(adj)
    res = []
    for r in sorted(rutas):
        d = {"archivo": r, "nombre": os.path.splitext(os.path.basename(r))[0],
             "mb": round(os.path.getsize(r) / 1048576, 1) if os.path.isfile(r) else None,
             "modificado": dt.datetime.fromtimestamp(os.path.getmtime(r)).isoformat(sep=" ", timespec="minutes") if os.path.isfile(r) else None,
             "adjunta": r in adj}
        if r in adj:
            try:
                with _conectar(adj[r]["base"]) as cn:
                    if _es_obra_pu(cn):
                        d["presupuestos"] = _filas(cn.execute(
                            "SELECT IdPresupuesto, Presupuesto, Nombre, (SELECT COUNT(*) FROM PuPresupuestosConceptos c "
                            "WHERE c.IdPresupuesto=p.IdPresupuesto) AS conceptos FROM PuPresupuestos p"))
                    else:
                        d["nota"] = "No es una obra PU (posible catálogo de precios)."
            except Exception as e:
                d["error"] = str(e)
        res.append(d)
    return res


@mcp.tool()
def neodata_resumen(obra: Optional[str] = None, presupuesto: Optional[str] = None) -> dict:
    """Resumen del presupuesto: datos generales, árbol de partidas con importes de costo directo,
    totales, IVA y alertas (conceptos sin matriz, insumos sin costo, recálculo pendiente)."""
    base = _resolver_obra(obra)
    with _conectar(base) as cn:
        p = _presupuesto(cn, presupuesto)
        m, parts, conc = _arbol(cn, p["IdPresupuesto"])
        total = sum(x["importe_cd"] for x in parts if x["nivel"] == 0)
        precio_nd = sum(x["importe_precio_neodata"] for x in parts if x["nivel"] == 0)
        hojas = [e for k, e in m.ei.items() if not m.tiene_matriz(k)]
        usados = {l["IdCodigoInsumo"] for ls in m.mat.values() for l in ls}
        sin_costo = [e["Codigo"] for e in hojas if e["IdExpIns"] in usados and not e["CostoBD"] and not e["EsPorcentaje"]]
        return {
            "obra": base, "presupuesto": {k: p[k] for k in ("IdPresupuesto", "Presupuesto", "Nombre", "PorcentajeIVA",
                                                            "Bloqueado", "Cerrado", "RequiereCalcular", "RequiereTotalizar")},
            "partidas": [{"partida": x["Partida"], "nivel": x["nivel"], "descripcion": x["DescripcionPartidaLarga"] or x["DescripcionPartida"],
                          "cantidad": x["Cantidad"], "conceptos": len(x["conceptos"]), "importe_costo_directo": round(x["importe_cd"], 2),
                          "importe_precio_neodata": round(x["importe_precio_neodata"], 2) or None} for x in parts],
            "total_costo_directo": round(total, 2),
            "total_precio_neodata": round(precio_nd, 2) or None,
            "iva": round(precio_nd * float(p["PorcentajeIVA"] or 0), 2) if precio_nd else None,
            "conceptos": len(conc),
            "conceptos_sin_matriz": (lambda l: {"cantidad": len(l), "codigos": l[:25]})(
                [c["Codigo"] for c in conc if not c["tiene_matriz"] and not c["costo_directo"]]),
            "insumos_sin_costo": {"cantidad": len(sin_costo), "codigos": sin_costo[:25]},
            "nota": ("El precio unitario (con indirectos, financiamiento, utilidad y cargos) lo calcula Neodata con su "
                     "libro de sobrecosto; 'importe_precio_neodata' viene vacío hasta que se ejecute Calcular en Neodata."),
        }


@mcp.tool()
def neodata_conceptos(obra: Optional[str] = None, partida: Optional[str] = None, buscar: Optional[str] = None,
                      presupuesto: Optional[str] = None, limite: int = 300) -> dict:
    """Catálogo de conceptos del presupuesto con unidad, cantidad, costo directo unitario e importe.
    Filtra por partida (clave/nombre) y/o texto en código o descripción."""
    base = _resolver_obra(obra)
    with _conectar(base) as cn:
        p = _presupuesto(cn, presupuesto)
        m, parts, conc = _arbol(cn, p["IdPresupuesto"])
        sel = parts
        if partida:
            pl = partida.lower()
            sel = [x for x in parts if (x["Partida"] or "").lower() == pl or pl in (x["DescripcionPartidaLarga"] or x["DescripcionPartida"] or "").lower()
                   or (x["PartidaWBS"] or "") == partida]
            if not sel:
                raise ValueError(f"No existe la partida '{partida}'.")
        out = []
        for x in sel:
            for c in x["conceptos"]:
                if buscar and buscar.lower() not in f"{c['Codigo']} {c['DescripcionLarga']}".lower():
                    continue
                out.append({"partida": x["Partida"], "codigo": c["Codigo"], "descripcion": c["DescripcionLarga"] or c["Descripcion"],
                            "unidad": c["Unidad"], "cantidad": c["Cantidad"], "cantidad_total": c["cantidad_total"],
                            "costo_directo": c["costo_directo"], "importe_costo_directo": c["importe_cd"],
                            "precio_neodata": c["Precio"], "tiene_matriz": c["tiene_matriz"]})
        return {"obra": base, "total": len(out), "conceptos": out[:limite], "truncado": len(out) > limite}


@mcp.tool()
def neodata_matriz(codigo: str, obra: Optional[str] = None, presupuesto: Optional[str] = None, niveles: int = 1) -> dict:
    """Análisis de precio unitario (matriz/APU) de un concepto o básico: insumos, cantidades, costos,
    importes y subtotales por tipo. niveles>1 despliega también las matrices de los básicos."""
    base = _resolver_obra(obra)
    with _conectar(base) as cn:
        p = _presupuesto(cn, presupuesto)
        m = Modelo(cn, p["IdPresupuesto"])
        ide = m.por_codigo.get(codigo.upper())
        if ide is None:
            raise ValueError(f"El código '{codigo}' no está en el presupuesto.")

        def armar(i, nivel):
            e = m.ei[i]
            ls = m.lineas(i)
            sub = defaultdict(float)
            for l in ls:
                sub[l["tipo"]] += l["importe"]
                if nivel < niveles and l["tiene_matriz"]:
                    l["matriz"] = armar(l["id_expins"], nivel + 1)
            return {"codigo": e["Codigo"], "descripcion": e["DescripcionLarga"] or e["Descripcion"], "unidad": e["Unidad"],
                    "costo_directo": m.costo(i), "lineas": ls, "subtotales_por_tipo": {k: round(v, m.dec) for k, v in sub.items()}}
        r = armar(ide, 1)
        if not r["lineas"]:
            r["nota"] = "Sin matriz: el costo es el capturado directamente en el insumo/concepto."
        return r


@mcp.tool()
def neodata_insumos(obra: Optional[str] = None, tipo: Optional[str] = None, buscar: Optional[str] = None,
                    solo_sin_costo: bool = False, presupuesto: Optional[str] = None, limite: int = 500) -> dict:
    """Insumos del presupuesto (materiales, mano de obra, equipo, subcontratos y básicos) con costo,
    número de matrices donde se usan y si tienen matriz propia."""
    base = _resolver_obra(obra)
    with _conectar(base) as cn:
        p = _presupuesto(cn, presupuesto)
        m = Modelo(cn, p["IdPresupuesto"])
        conceptos = {r[0] for r in cn.execute("SELECT IdExpIns FROM PuPresupuestosConceptos WHERE IdPresupuesto=?", p["IdPresupuesto"])}
        usos = defaultdict(int)
        for ls in m.mat.values():
            for l in ls:
                usos[l["IdCodigoInsumo"]] += 1
        idt = _tipo_id(tipo) if tipo else None
        out = []
        for k, e in sorted(m.ei.items(), key=lambda kv: (kv[1]["IdTipo"], kv[1]["Codigo"])):
            if k in conceptos and not usos[k]:
                continue  # es un concepto del catálogo, no un insumo
            if idt and e["IdTipo"] != idt:
                continue
            if buscar and buscar.lower() not in f"{e['Codigo']} {e['DescripcionLarga']}".lower():
                continue
            c = m.costo(k)
            if solo_sin_costo and (c or e["EsPorcentaje"]):
                continue
            out.append({"codigo": e["Codigo"], "descripcion": e["DescripcionLarga"] or e["Descripcion"], "unidad": e["Unidad"],
                        "tipo": TIPOS.get(e["IdTipo"], e["IdTipo"]), "costo": c, "es_porcentaje": bool(e["EsPorcentaje"]),
                        "tiene_matriz": m.tiene_matriz(k), "usado_en_matrices": usos[k]})
        return {"obra": base, "total": len(out), "insumos": out[:limite], "truncado": len(out) > limite}


@mcp.tool()
def neodata_explosion_insumos(obra: Optional[str] = None, tipo: Optional[str] = None,
                              presupuesto: Optional[str] = None) -> dict:
    """Explosión de insumos de todo el presupuesto: cantidad total e importe por insumo (despliega básicos),
    con totales por tipo (materiales, mano de obra, equipo, subcontratos)."""
    base = _resolver_obra(obra)
    with _conectar(base) as cn:
        p = _presupuesto(cn, presupuesto)
        m, parts, conc = _arbol(cn, p["IdPresupuesto"])
        acum: dict = {}
        sin_matriz = []
        for c in conc:
            if not c["tiene_matriz"]:
                sin_matriz.append(c["Codigo"])
                continue
            m.explotar(c["IdExpIns"], c["cantidad_total"], acum)
        filas = sorted(acum.values(), key=lambda a: (str(a["tipo"]), -a["importe"]))
        if tipo:
            nt = TIPOS.get(_tipo_id(tipo))
            filas = [f for f in filas if f["tipo"] == nt]
        tot = defaultdict(float)
        for f in filas:
            f["cantidad"] = round(f["cantidad"], 6)
            f["importe"] = round(f["importe"], 2)
            tot[f["tipo"]] += f["importe"]
        return {"obra": base, "insumos": filas, "totales_por_tipo": {k: round(v, 2) for k, v in tot.items()},
                "total": round(sum(tot.values()), 2),
                "conceptos_sin_matriz_excluidos": {"cantidad": len(sin_matriz), "codigos": sin_matriz[:30]}}


_PROHIBIDO = re.compile(r"\b(insert|update|delete|merge|drop|alter|create|exec|execute|truncate|grant|revoke|deny|"
                        r"backup|restore|dbcc|into|openrowset|opendatasource|bulk|shutdown|kill|reconfigure|sp_\w+|xp_\w+)\b", re.I)


@mcp.tool()
def neodata_consulta_sql(sql: str, obra: Optional[str] = None, limite: int = 500) -> dict:
    """Consulta SQL de SOLO LECTURA (SELECT/WITH) sobre la base de la obra, para lo que no cubren las
    otras herramientas. Tablas clave: PuPresupuestos, PuPresupuestosPartidas, PuPresupuestosConceptos,
    PuCatalogo, PuExpIns, PuExpInsCostos, PuMatrices, PuUnidades, PuTipoInsumos."""
    s = sql.strip().rstrip(";")
    if not re.match(r"^\s*(select|with)\b", s, re.I) or _PROHIBIDO.search(s) or ";" in s:
        raise ValueError("Solo se permite una sentencia SELECT/WITH sin modificaciones.")
    base = _resolver_obra(obra)
    with _conectar(base, autocommit=False) as cn:
        try:
            cur = cn.execute(s)
            cols = [c[0] for c in cur.description] if cur.description else []
            rows = cur.fetchmany(limite + 1)
        finally:
            cn.rollback()
    return {"columnas": cols, "filas": [[_num(v) for v in r] for r in rows[:limite]], "truncado": len(rows) > limite}


# ----------------------------------------------------------------------------
# Archivos: exportar / plantilla / importar
# ----------------------------------------------------------------------------
def _ruta_salida(ruta: Optional[str], nombre: str) -> str:
    if ruta:
        ruta = os.path.abspath(os.path.expanduser(ruta))
        if os.path.isdir(ruta):
            ruta = os.path.join(ruta, nombre)
    else:
        os.makedirs(CARPETA_SALIDA, exist_ok=True)
        ruta = os.path.join(CARPETA_SALIDA, nombre)
    os.makedirs(os.path.dirname(ruta), exist_ok=True)
    return ruta


def _hoja(wb, titulo, encabezados, filas, anchos=None, formatos=None):
    from openpyxl.styles import Alignment, Font, PatternFill
    ws = wb.create_sheet(titulo)
    ws.append(encabezados)
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="1F4E78")
        c.alignment = Alignment(vertical="center", wrap_text=True)
    for f in filas:
        ws.append(f)
    ws.freeze_panes = "A2"
    for i, w in enumerate(anchos or [], start=1):
        ws.column_dimensions[ws.cell(1, i).column_letter].width = w
    for col, fmt in (formatos or {}).items():
        for cell in ws[col][1:]:
            cell.number_format = fmt
    return ws


@mcp.tool()
def neodata_exportar_excel(obra: Optional[str] = None, ruta: Optional[str] = None, presupuesto: Optional[str] = None,
                           factor_sobrecosto: Optional[float] = None) -> dict:
    """Exporta el presupuesto a Excel (.xlsx): Resumen, Presupuesto (catálogo de conceptos con importes),
    Matrices (APU de cada concepto), Insumos y Explosión de insumos. factor_sobrecosto (p. ej. 1.25)
    agrega columnas de precio estimado = costo directo × factor."""
    from openpyxl import Workbook
    base = _resolver_obra(obra)
    with _conectar(base) as cn:
        p = _presupuesto(cn, presupuesto)
        m, parts, conc = _arbol(cn, p["IdPresupuesto"], factor_sobrecosto)
    nombre_obra = os.path.splitext(os.path.basename(base))[0]
    ruta = _ruta_salida(ruta, f"{nombre_obra}_{p['Presupuesto']}_{dt.datetime.now():%Y%m%d_%H%M}.xlsx")
    wb = Workbook()
    wb.remove(wb.active)
    mon = '#,##0.00'
    total = sum(x["importe_cd"] for x in parts if x["nivel"] == 0)
    # Resumen
    filas = [["  " * x["nivel"] + (x["Partida"] or ""), x["DescripcionPartidaLarga"] or x["DescripcionPartida"], len(x["conceptos"]),
              round(x["importe_cd"], 2), (x["importe_cd"] / total) if total else 0] for x in parts]
    filas.append(["TOTAL COSTO DIRECTO", "", len(conc), round(total, 2), 1 if total else 0])
    if factor_sobrecosto:
        filas.append([f"PRECIO ESTIMADO (× {factor_sobrecosto})", "", "", round(total * factor_sobrecosto, 2), ""])
        filas.append([f"IVA {float(p['PorcentajeIVA'] or 0):.0%}", "", "", round(total * factor_sobrecosto * float(p["PorcentajeIVA"] or 0), 2), ""])
    ws = _hoja(wb, "Resumen", ["Partida", "Descripción", "Conceptos", "Importe C.D.", "%"], filas, [28, 50, 11, 18, 9],
               {"D": mon, "E": "0.00%"})
    ws.insert_rows(1, 3)
    ws["A1"] = f"Obra: {nombre_obra}   Presupuesto: {p['Presupuesto']} — {p['Nombre']}"
    ws["A2"] = f"Generado {dt.datetime.now():%Y-%m-%d %H:%M} con Neodata MCP. Importes en pesos (M.N.), costo directo calculado de las matrices."
    # Presupuesto
    enc = ["Partida", "Código", "Descripción", "Unidad", "Cantidad", "C.D. unitario", "Importe C.D.", "Precio Neodata"]
    fmt = {"E": "#,##0.0000", "F": mon, "G": mon, "H": mon}
    if factor_sobrecosto:
        enc += ["P.U. estimado", "Importe estimado"]
        fmt.update({"I": mon, "J": mon})
    filas = []
    for x in parts:
        filas.append([" / ".join(x["ruta"]), "", (x["DescripcionPartidaLarga"] or x["DescripcionPartida"] or "").upper(), "", None, None, round(x["importe_cd"], 2), None])
        for c in x["conceptos"]:
            f = [x["Partida"], c["Codigo"], c["DescripcionLarga"] or c["Descripcion"], c["Unidad"], c["cantidad_total"],
                 c["costo_directo"], c["importe_cd"], c["Precio"]]
            if factor_sobrecosto:
                f += [c["precio_estimado"], c["importe_estimado"]]
            filas.append(f)
    ws = _hoja(wb, "Presupuesto", enc, filas, [22, 14, 70, 8, 12, 14, 16, 14, 14, 16], fmt)
    from openpyxl.styles import Alignment, Font
    for row in ws.iter_rows(min_row=2):
        row[2].alignment = Alignment(wrap_text=True, vertical="top")
        if not row[1].value:
            for cell in row:
                cell.font = Font(bold=True)
    # Matrices
    filas = []
    for c in conc:
        filas.append([c["Codigo"], "", (c["DescripcionLarga"] or c["Descripcion"]), c["Unidad"], None, None, c["costo_directo"]])
        for l in m.lineas(c["IdExpIns"]):
            filas.append(["", l["codigo"], l["descripcion"], l["unidad"], l["cantidad"], l["costo"], l["importe"], l["tipo"]])
        filas.append([])
    ws = _hoja(wb, "Matrices", ["Concepto", "Insumo", "Descripción", "Unidad", "Cantidad", "Costo", "Importe", "Tipo"],
               filas, [14, 14, 70, 8, 12, 14, 14, 20], {"E": "#,##0.000000", "F": mon, "G": mon})
    for row in ws.iter_rows(min_row=2):
        if row[0].value:
            for cell in row:
                cell.font = Font(bold=True)
    # Insumos
    filas = [[e["Codigo"], e["DescripcionLarga"] or e["Descripcion"], e["Unidad"], TIPOS.get(e["IdTipo"]), m.costo(k)]
             for k, e in sorted(m.ei.items(), key=lambda kv: (kv[1]["IdTipo"], kv[1]["Codigo"])) if not m.tiene_matriz(k)
             and k not in {c["IdExpIns"] for c in conc}]
    _hoja(wb, "Insumos", ["Código", "Descripción", "Unidad", "Tipo", "Costo"], filas, [16, 70, 8, 22, 14], {"E": mon})
    # Explosión
    acum: dict = {}
    for c in conc:
        if c["tiene_matriz"]:
            m.explotar(c["IdExpIns"], c["cantidad_total"], acum)
    filas = [[a["tipo"], a["codigo"], a["descripcion"], a["unidad"], round(a["cantidad"], 6), a["costo"], round(a["importe"], 2)]
             for a in sorted(acum.values(), key=lambda a: (str(a["tipo"]), -a["importe"]))]
    _hoja(wb, "Explosión", ["Tipo", "Código", "Descripción", "Unidad", "Cantidad", "Costo", "Importe"], filas,
          [22, 16, 70, 8, 14, 14, 16], {"E": "#,##0.0000", "F": mon, "G": mon})
    wb.save(ruta)
    return {"archivo": ruta, "conceptos": len(conc), "total_costo_directo": round(total, 2)}


@mcp.tool()
def neodata_plantilla_importacion(ruta: Optional[str] = None) -> dict:
    """Crea una plantilla Excel para cargar datos a Neodata con neodata_importar_excel:
    hojas Insumos, Conceptos y Matrices con ejemplos (unidades métricas)."""
    from openpyxl import Workbook
    ruta = _ruta_salida(ruta, "plantilla_importacion_neodata.xlsx")
    wb = Workbook()
    wb.remove(wb.active)
    _hoja(wb, "Insumos", ["Codigo", "Descripcion", "Unidad", "Tipo", "Costo"], [
        ["MAT-CEM-01", "Cemento gris CPC 30R, bulto 50 kg", "PZA", "MATERIALES", 245.00],
        ["MAT-ARE-01", "Arena de río", "m³", "MATERIALES", 520.00],
        ["MO-ALB-01", "Cuadrilla albañil + peón", "JOR", "MANO DE OBRA", 1450.00],
        ["EQ-HM-01", "Herramienta menor (3% de mano de obra)", "%", "EQUIPO Y HERRAMIENTA", None],
    ], [16, 60, 10, 24, 12])
    _hoja(wb, "Conceptos", ["Partida", "Codigo", "Descripcion", "Unidad", "Cantidad"], [
        ["CIMENTACION", "CIM-001", "Plantilla de concreto f'c=100 kg/cm² de 5 cm de espesor", "m²", 45.5],
    ], [20, 14, 70, 10, 12])
    _hoja(wb, "Matrices", ["Concepto", "Insumo", "Cantidad", "Dividir"], [
        ["CIM-001", "MAT-CEM-01", 0.35, ""],
        ["CIM-001", "MAT-ARE-01", 0.03, ""],
        ["CIM-001", "MO-ALB-01", 25, "SI"],
        ["CIM-001", "EQ-HM-01", 0.03, ""],
    ], [14, 14, 12, 10])
    ws = wb.create_sheet("LEEME")
    for t in [
        "Hojas: Insumos, Conceptos y Matrices (cualquiera puede ir vacía o faltar).",
        "Insumos: Tipo = MATERIALES | MANO DE OBRA | EQUIPO Y HERRAMIENTA | BASICOS | SUBCONTRATOS. Costo en pesos.",
        "  Unidad '%' marca un insumo porcentual; en Matrices su Cantidad es la fracción (0.03 = 3%) del tipo indicado",
        "  en la columna opcional 'PorcentajeDe' de Insumos (por omisión MANO DE OBRA).",
        "Conceptos: si la partida no existe se crea. Si el código ya existe se actualiza descripción/unidad/cantidad.",
        "Matrices: reemplaza la matriz completa de cada concepto listado. Dividir=SI significa costo/cantidad",
        "  (rendimiento, p. ej. 25 m²/jornada); vacío = costo × cantidad.",
        "Unidades métricas: m, m², m³, kg, t, PZA, JOR, HR, LT, %.",
        "Primero corre neodata_importar_excel con dry_run=True y revisa el reporte.",
    ]:
        ws.append([t])
    ws.column_dimensions["A"].width = 120
    wb.save(ruta)
    return {"archivo": ruta}


# ----------------------------------------------------------------------------
# Escritura
# ----------------------------------------------------------------------------
def _tipo_id(tipo: Any) -> int:
    if isinstance(tipo, int) or (isinstance(tipo, str) and tipo.strip().isdigit()):
        t = int(tipo)
        if t in TIPOS:
            return t
    s = str(tipo or "").strip().lower()
    for k, v in TIPOS.items():
        if v.lower() == s:
            return k
    if s in ALIAS_TIPO:
        return ALIAS_TIPO[s]
    raise ValueError(f"Tipo de insumo no válido: '{tipo}'. Usa {list(TIPOS.values())}.")


def _unidad_id(cur, unidad: str) -> int:
    u = (unidad or "").strip()
    if not u:
        raise ValueError("Falta la unidad.")
    u = ALIAS_UNIDAD.get(u.lower(), u)
    r = cur.execute("SELECT IdUnidad FROM PuUnidades WHERE Unidad=?", u).fetchone() or \
        cur.execute("SELECT IdUnidad FROM PuUnidades WHERE UPPER(Unidad)=UPPER(?)", u).fetchone()
    if r:
        return r[0]
    res = _ultimo_resultado(cur.execute(
        "SET NOCOUNT ON; INSERT INTO PuUnidades (Unidad, Descripcion) VALUES (?, ?); SELECT CAST(SCOPE_IDENTITY() AS int) AS id;",
        u[:20], u[:40]))
    return res[-1]["id"]


def _catalogo(cur, codigo, descripcion=None, unidad=None, tipo=None, porcentaje_de=None, crear=True) -> int:
    codigo = str(codigo).strip()
    if not codigo or len(codigo) > 20:
        raise ValueError(f"Código inválido '{codigo}' (1 a 20 caracteres).")
    r = cur.execute("SELECT IdCodigo FROM PuCatalogo WHERE Codigo=?", codigo).fetchone()
    es_porc = None
    if unidad and str(unidad).strip() == "%":
        es_porc = _tipo_id(porcentaje_de or "MANO DE OBRA")
    if r:
        sets, vals = [], []
        if descripcion:
            sets += ["Descripcion=?", "DescripcionLarga=?"]
            vals += [str(descripcion)[:40], str(descripcion)]
        if unidad:
            sets.append("IdUnidad=?")
            vals.append(_unidad_id(cur, unidad))
        if tipo is not None:
            sets.append("IdTipo=?")
            vals.append(_tipo_id(tipo))
        if es_porc is not None:
            sets.append("EsPorcentaje=?")
            vals.append(es_porc)
        if sets:
            cur.execute(f"UPDATE PuCatalogo SET {', '.join(sets)} WHERE IdCodigo=?", *vals, r[0])
        return r[0]
    if not crear:
        raise ValueError(f"El código '{codigo}' no existe en el catálogo.")
    if not descripcion or not unidad:
        raise ValueError(f"Para crear '{codigo}' se necesitan descripción y unidad.")
    idu = _unidad_id(cur, unidad)
    res = _ultimo_resultado(cur.execute(
        "SET NOCOUNT ON; INSERT INTO PuCatalogo (Codigo, IdUnidad, Descripcion, DescripcionLarga, IdTipo, EsPorcentaje) "
        "VALUES (?,?,?,?,?,?); SELECT CAST(SCOPE_IDENTITY() AS int) AS id;",
        codigo, idu, str(descripcion)[:40], str(descripcion), _tipo_id(tipo if tipo is not None else 4), es_porc or 0))
    return res[-1]["id"]


def _expins(cur, idp: int, idcodigo: int, costo: Optional[float] = None) -> int:
    r = cur.execute("SELECT IdExpIns, IdProveedor, InsumoIntegrado, EsCostoHorario, IdIntelimat FROM PuExpIns "
                    "WHERE IdPresupuesto=? AND IdCodigo=?", idp, idcodigo).fetchone()
    if r and costo is None:
        return r[0]
    if r:
        sql = ("SET NOCOUNT ON; DECLARE @id int = ?; EXEC NeoExpIns_GrabarActualizar @IdExpIns=@id OUTPUT, @IdPresupuesto=?, "
               "@IdCodigo=?, @IdUsuario=?, @Costo=?, @IdMoneda=?, @IdProveedor=?, @CostoFecha=NULL, @InsumoIntegrado=?, "
               "@EsCostoHorario=?, @IdIntelimat=?; SELECT @id AS id;")
        res = _ultimo_resultado(cur.execute(sql, r[0], idp, idcodigo, ID_USUARIO, costo, MONEDA, r[1], r[2], r[3], r[4]))
    else:
        sql = ("SET NOCOUNT ON; DECLARE @id int = 0; EXEC NeoExpIns_GrabarActualizar @IdExpIns=@id OUTPUT, @IdPresupuesto=?, "
               "@IdCodigo=?, @IdUsuario=?, @Costo=?, @IdMoneda=?; SELECT @id AS id;")
        res = _ultimo_resultado(cur.execute(sql, idp, idcodigo, ID_USUARIO, costo or 0, MONEDA))
    return res[-1]["id"]


def _partida_id(cur, idp: int, partida: Any, crear: bool = False, padre: Any = None) -> int:
    if isinstance(partida, int) or (isinstance(partida, str) and partida.isdigit()):
        r = cur.execute("SELECT IdPresupuestoPartida FROM PuPresupuestosPartidas WHERE IdPresupuesto=? AND IdPresupuestoPartida=?",
                        idp, int(partida)).fetchone()
        if r:
            return r[0]
    s = str(partida).strip()
    r = cur.execute("SELECT IdPresupuestoPartida FROM PuPresupuestosPartidas WHERE IdPresupuesto=? AND (UPPER(Partida)=UPPER(?) "
                    "OR PartidaWBS=? OR UPPER(DescripcionPartida)=UPPER(?))", idp, s, s, s).fetchall()
    if len(r) == 1:
        return r[0][0]
    if len(r) > 1:
        raise ValueError(f"La partida '{s}' es ambigua ({len(r)} coincidencias); usa su IdPresupuestoPartida.")
    if not crear:
        raise ValueError(f"No existe la partida '{s}'.")
    id_padre = _partida_id(cur, idp, padre) if padre else None
    sql = ("SET NOCOUNT ON; DECLARE @id int = 0, @ren int = NULL, @folio int; "
           "EXEC @folio = NeoPresupuestoConsecutivo_DameActual 'FolioEstructuraPartidas', ?; "
           "EXEC NeoPartidas_GrabarActualizar @IdPresupuestoPartida=@id OUTPUT, @IdPresupuesto=?, @IdPartidaPadre=?, @Partida=?, "
           "@IdImagen=NULL, @Renglon=@ren OUTPUT, @Cantidad=1, @Costo=0, @Precio=0, @MetrosCuadrados=1, @DescripcionPartida=?, "
           "@DescripcionPartidaLarga=?, @FolioEstructuraPartidas=@folio OUTPUT, @IdMoneda=?; SELECT @id AS id;")
    res = _ultimo_resultado(cur.execute(sql, idp, idp, id_padre, s.upper()[:20], s[:40], s, MONEDA))
    return res[-1]["id"]


def _guardar_concepto(cur, idp, partida, codigo, descripcion=None, unidad=None, cantidad=None, crear_partida=True) -> dict:
    idpar = _partida_id(cur, idp, partida, crear=crear_partida)
    idcod = _catalogo(cur, codigo, descripcion, unidad, None if cur.execute(
        "SELECT 1 FROM PuCatalogo WHERE Codigo=?", str(codigo).strip()).fetchone() else 4)
    ide = _expins(cur, idp, idcod)
    r = cur.execute("SELECT IdPresupuestoConcepto, IdPresupuestoPartida FROM PuPresupuestosConceptos WHERE IdPresupuesto=? AND IdExpIns=?",
                    idp, ide).fetchone()
    if r:
        if r[1] != idpar:
            raise ValueError(f"El concepto '{codigo}' ya está en otra partida (Id {r[1]}). Muévelo en Neodata.")
        if cantidad is not None:
            cur.execute("UPDATE PuPresupuestosConceptos SET Cantidad=? WHERE IdPresupuestoConcepto=?", cantidad, r[0])
        return {"accion": "actualizado", "codigo": codigo, "id_concepto": r[0]}
    sql = ("SET NOCOUNT ON; DECLARE @idc int = 0, @ren int = NULL, @ide int = ?, @ctl int = NULL; "
           "EXEC NeoConceptos_GrabarActualizar @IdPresupuesto=?, @IdPresupuestoPartida=?, @IdPresupuestoConcepto=@idc OUTPUT, "
           "@IdCodigo=?, @IdAgrupador=NULL, @Renglon=@ren OUTPUT, @IdProveedor=NULL, @Costo=0, @IdExpIns=@ide OUTPUT, "
           "@Cantidad=?, @CodigoAuxiliar=NULL, @DescripcionAuxiliar=NULL, @IndirectoEspecial=0, @PorcentajeIndirectoEspecial=0, "
           "@Precio=0, @RutaGenerador=NULL, @IdEstimacionTipo=1, @Control=@ctl OUTPUT, @IdMoneda=?, @IdInteliMat=?, "
           "@GrabarCosto=0, @IdUsuario=?; SELECT @idc AS id;")
    intelimat = cur.execute("SELECT IdIntelimat FROM PuExpIns WHERE IdExpIns=?", ide).fetchone()[0]
    res = _ultimo_resultado(cur.execute(sql, ide, idp, idpar, idcod, cantidad or 0, MONEDA, intelimat, ID_USUARIO))
    return {"accion": "creado", "codigo": codigo, "id_concepto": res[-1]["id"]}


def _definir_matriz(cur, idp, concepto: str, lineas: list[dict], reemplazar: bool = True) -> dict:
    r = cur.execute("SELECT EI.IdExpIns FROM PuExpIns EI JOIN PuCatalogo C ON C.IdCodigo=EI.IdCodigo "
                    "WHERE EI.IdPresupuesto=? AND C.Codigo=?", idp, str(concepto).strip()).fetchone()
    if not r:
        raise ValueError(f"El concepto/básico '{concepto}' no está en el presupuesto; créalo primero.")
    idm = r[0]
    borradas = 0
    if reemplazar:
        for (idmat,) in cur.execute("SELECT IdMatriz FROM PuMatrices WHERE IdCodigoMatriz=?", idm).fetchall():
            cur.execute("EXEC NeoMatrices_Eliminar ?", idmat)
            borradas += 1
    n = 0
    for l in lineas:
        cod = l.get("insumo") or l.get("codigo")
        cant = float(l.get("cantidad") if l.get("cantidad") is not None else l.get("volumen"))
        div = str(l.get("dividir") or "").strip().lower() in ("1", "true", "si", "sí", "s", "yes", "/")
        idcod = _catalogo(cur, cod, l.get("descripcion"), l.get("unidad"), l.get("tipo"), l.get("porcentaje_de"),
                          crear=bool(l.get("descripcion") and l.get("unidad")))
        ide = _expins(cur, idp, idcod, l.get("costo"))
        if ide == idm:
            raise ValueError(f"'{cod}' no puede estar dentro de su propia matriz.")
        sql = ("SET NOCOUNT ON; DECLARE @idm int, @ren int = NULL; EXEC NeoMatrices_GrabarSoloMatriz @IdPresupuesto=?, "
               "@IdMatriz=@idm OUTPUT, @Renglon=@ren OUTPUT, @IdCodigoMatriz=?, @IdCodigoInsumo=?, @Volumen=?, @Dividir=?, "
               "@Expresion=NULL; SELECT @idm AS id;")
        _ultimo_resultado(cur.execute(sql, idp, idm, ide, cant, 1 if div else 0))
        n += 1
    return {"concepto": concepto, "lineas_borradas": borradas, "lineas_agregadas": n}


def _recalcular_costos(cur, idp: int) -> int:
    """Guarda en PuExpInsCostos el costo directo de cada matriz (como hace Neodata al calcular)."""
    m = Modelo(cur.connection, idp)
    cambios = 0
    for ide in list(m.mat):
        c = m.costo(ide)
        if abs(c - float(m.ei[ide]["CostoBD"] or 0)) > 0.0001:
            _expins(cur, idp, m.ei[ide]["IdCodigo"], c)
            cambios += 1
    return cambios


def _escribir(obra, presupuesto, dry_run: bool, accion, descripcion: str) -> dict:
    """Marco común de escritura: validaciones, respaldo, transacción, recálculo de costos y banderas."""
    base = _resolver_obra(obra)
    if not dry_run and neodata_abierto():
        raise RuntimeError("Neodata está abierto. Ciérralo para escribir (evita que sobrescriba los cambios). "
                           "Puedes usar dry_run=True mientras tanto.")
    respaldo = None
    with _conectar(base) as cn0:
        p = _presupuesto(cn0, presupuesto)
    if p["Bloqueado"] or p["Cerrado"]:
        raise RuntimeError(f"El presupuesto {p['Presupuesto']} está bloqueado/cerrado en Neodata.")
    if not dry_run:
        respaldo = _respaldar(base, motivo=descripcion)
    cn = _conectar(base, autocommit=False)
    try:
        cur = cn.cursor()
        resultado = accion(cur, p["IdPresupuesto"])
        costos = _recalcular_costos(cur, p["IdPresupuesto"])
        cur.execute("UPDATE PuPresupuestos SET RequiereCalcular=1, RequiereTotalizar=1 WHERE IdPresupuesto=?", p["IdPresupuesto"])
        if dry_run:
            cn.rollback()
        else:
            cn.commit()
    except Exception:
        try:
            cn.rollback()
        except Exception:
            pass
        raise
    finally:
        cn.close()
    out = {"dry_run": dry_run, "obra": base, "presupuesto": p["Presupuesto"], "resultado": resultado,
           "costos_de_matrices_actualizados": costos}
    if dry_run:
        out["nota"] = "Simulación: se ejecutó y se revirtió. Repite con dry_run=False para guardar (Neodata cerrado)."
    else:
        out["respaldo"] = respaldo
        out["nota"] = "Guardado. Al abrir la obra en Neodata ejecuta Calcular para actualizar precios unitarios y totales."
    return out


@mcp.tool()
def neodata_crear_partida(nombre: str, padre: Optional[str] = None, obra: Optional[str] = None,
                          presupuesto: Optional[str] = None, dry_run: bool = True) -> dict:
    """Crea una partida (o subpartida si se indica 'padre'). dry_run=True por omisión."""
    def accion(cur, idp):
        try:
            existente = _partida_id(cur, idp, nombre)
            return {"accion": "ya_existia", "id_partida": existente}
        except ValueError:
            return {"accion": "creada", "id_partida": _partida_id(cur, idp, nombre, crear=True, padre=padre)}
    return _escribir(obra, presupuesto, dry_run, accion, f"partida {nombre}")


@mcp.tool()
def neodata_guardar_concepto(partida: str, codigo: str, descripcion: Optional[str] = None, unidad: Optional[str] = None,
                             cantidad: Optional[float] = None, obra: Optional[str] = None, presupuesto: Optional[str] = None,
                             dry_run: bool = True) -> dict:
    """Crea o actualiza un concepto del presupuesto en la partida indicada (la partida se crea si no existe).
    Si el código ya existe actualiza descripción, unidad y/o cantidad. dry_run=True por omisión."""
    return _escribir(obra, presupuesto, dry_run,
                     lambda cur, idp: _guardar_concepto(cur, idp, partida, codigo, descripcion, unidad, cantidad),
                     f"concepto {codigo}")


@mcp.tool()
def neodata_guardar_insumo(codigo: str, descripcion: Optional[str] = None, unidad: Optional[str] = None,
                           tipo: Optional[str] = None, costo: Optional[float] = None, porcentaje_de: Optional[str] = None,
                           obra: Optional[str] = None, presupuesto: Optional[str] = None, dry_run: bool = True) -> dict:
    """Crea o actualiza un insumo (MATERIALES, MANO DE OBRA, EQUIPO Y HERRAMIENTA, BASICOS, SUBCONTRATOS) y su costo
    en pesos. Unidad '%' = insumo porcentual del tipo 'porcentaje_de'. Recalcula el costo de las matrices que lo usan."""
    def accion(cur, idp):
        existe = cur.execute("SELECT 1 FROM PuCatalogo WHERE Codigo=?", codigo.strip()).fetchone()
        if not existe and tipo is None:
            raise ValueError("Para crear un insumo indica 'tipo'.")
        idc = _catalogo(cur, codigo, descripcion, unidad, tipo, porcentaje_de)
        _expins(cur, idp, idc, costo)
        return {"accion": "actualizado" if existe else "creado", "codigo": codigo, "costo": costo}
    return _escribir(obra, presupuesto, dry_run, accion, f"insumo {codigo}")


@mcp.tool()
def neodata_definir_matriz(concepto: str, lineas: list[dict], reemplazar: bool = True, obra: Optional[str] = None,
                           presupuesto: Optional[str] = None, dry_run: bool = True) -> dict:
    """Define la matriz (APU) de un concepto o básico. lineas = [{"insumo": "MAT-01", "cantidad": 0.35,
    "dividir": false}, ...]. Opcional por línea para crear el insumo al vuelo: descripcion, unidad, tipo, costo.
    dividir=true usa costo/cantidad (rendimiento). reemplazar=True borra la matriz anterior. dry_run=True por omisión;
    la simulación devuelve el costo directo resultante."""
    def accion(cur, idp):
        r = _definir_matriz(cur, idp, concepto, lineas, reemplazar)
        m = Modelo(cur.connection, idp)
        ide = m.por_codigo.get(concepto.upper())
        r["costo_directo_resultante"] = m.costo(ide) if ide else None
        r["lineas"] = [{k: l[k] for k in ("codigo", "unidad", "cantidad", "dividir", "costo", "importe")} for l in m.lineas(ide)] if ide else []
        return r
    return _escribir(obra, presupuesto, dry_run, accion, f"matriz {concepto}")


@mcp.tool()
def neodata_eliminar_concepto(codigo: str, obra: Optional[str] = None, presupuesto: Optional[str] = None,
                              dry_run: bool = True) -> dict:
    """Quita un concepto del presupuesto (usa NeoConceptos_Eliminar; el código queda en el catálogo).
    dry_run=True por omisión."""
    def accion(cur, idp):
        rows = cur.execute("SELECT PC.IdPresupuestoConcepto FROM PuPresupuestosConceptos PC JOIN PuExpIns EI ON EI.IdExpIns=PC.IdExpIns "
                           "JOIN PuCatalogo C ON C.IdCodigo=EI.IdCodigo WHERE PC.IdPresupuesto=? AND C.Codigo=?", idp, codigo.strip()).fetchall()
        if not rows:
            raise ValueError(f"El concepto '{codigo}' no está en el presupuesto.")
        for (idc,) in rows:
            cur.execute("EXEC NeoConceptos_Eliminar ?", idc)
        return {"accion": "eliminado", "codigo": codigo, "registros": len(rows)}
    return _escribir(obra, presupuesto, dry_run, accion, f"eliminar {codigo}")


def _leer_hoja(wb, nombre):
    if nombre not in wb.sheetnames:
        return []
    ws = wb[nombre]
    filas = list(ws.iter_rows(values_only=True))
    if not filas:
        return []
    enc = [str(h or "").strip().lower() for h in filas[0]]
    out = []
    for i, f in enumerate(filas[1:], start=2):
        if not any(v not in (None, "") for v in f):
            continue
        d = {enc[j]: (v.strip() if isinstance(v, str) else v) for j, v in enumerate(f) if j < len(enc) and enc[j]}
        d["_fila"] = i
        out.append(d)
    return out


@mcp.tool()
def neodata_importar_excel(ruta: str, obra: Optional[str] = None, presupuesto: Optional[str] = None,
                           dry_run: bool = True) -> dict:
    """Importa un Excel con hojas Insumos / Conceptos / Matrices (formato de neodata_plantilla_importacion).
    Todo o nada: si una fila falla no se guarda nada y se indica la fila. dry_run=True por omisión."""
    from openpyxl import load_workbook
    ruta = os.path.abspath(os.path.expanduser(ruta))
    wb = load_workbook(ruta, data_only=True, read_only=True)
    insumos, conceptos, matrices = _leer_hoja(wb, "Insumos"), _leer_hoja(wb, "Conceptos"), _leer_hoja(wb, "Matrices")
    wb.close()
    if not (insumos or conceptos or matrices):
        raise ValueError("El archivo no tiene filas en las hojas Insumos, Conceptos ni Matrices.")

    def accion(cur, idp):
        rep = {"insumos": [], "conceptos": [], "matrices": []}
        hoja = fila = None
        try:
            hoja = "Insumos"
            for d in insumos:
                fila = d["_fila"]
                existe = cur.execute("SELECT 1 FROM PuCatalogo WHERE Codigo=?", str(d.get("codigo")).strip()).fetchone()
                idc = _catalogo(cur, d.get("codigo"), d.get("descripcion"), d.get("unidad"), d.get("tipo") or (None if existe else 1),
                                d.get("porcentajede") or d.get("porcentaje_de"))
                costo = d.get("costo")
                _expins(cur, idp, idc, float(costo) if costo not in (None, "") else None)
                rep["insumos"].append(f"{d.get('codigo')}: {'actualizado' if existe else 'creado'}")
            hoja = "Conceptos"
            for d in conceptos:
                fila = d["_fila"]
                cant = d.get("cantidad")
                r = _guardar_concepto(cur, idp, d.get("partida"), d.get("codigo"), d.get("descripcion"), d.get("unidad"),
                                      float(cant) if cant not in (None, "") else None)
                rep["conceptos"].append(f"{d.get('codigo')}: {r['accion']}")
            hoja = "Matrices"
            grupos: dict[str, list] = defaultdict(list)
            for d in matrices:
                grupos[str(d.get("concepto")).strip()].append(d)
            for conc, ls in grupos.items():
                fila = ls[0]["_fila"]
                r = _definir_matriz(cur, idp, conc, ls, reemplazar=True)
                rep["matrices"].append(f"{conc}: {r['lineas_agregadas']} líneas")
        except Exception as e:
            raise ValueError(f"Hoja {hoja}, fila {fila}: {e}") from e
        m = Modelo(cur.connection, idp)
        rep["costos_directos"] = {c: m.costo(m.por_codigo[c.upper()]) for c in
                                  [str(d.get("codigo")).strip() for d in conceptos] + list(dict.fromkeys(str(d.get("concepto")).strip() for d in matrices))
                                  if c.upper() in m.por_codigo}
        return rep
    return _escribir(obra, presupuesto, dry_run, accion, "importar excel")


# ----------------------------------------------------------------------------
# Respaldos
# ----------------------------------------------------------------------------
def _respaldar(base: str, motivo: str = "manual") -> str:
    os.makedirs(CARPETA_RESPALDOS, exist_ok=True)
    stem = re.sub(r"[^\w\-]+", "_", os.path.splitext(os.path.basename(base))[0]).strip("_")
    archivo = os.path.join(CARPETA_RESPALDOS, f"{stem}_{dt.datetime.now():%Y%m%d_%H%M%S_%f}"[:-3] + ".bak")
    q = base.replace("]", "]]")
    with _conectar() as cn:
        cur = cn.execute(f"BACKUP DATABASE [{q}] TO DISK=? WITH COPY_ONLY, INIT, NAME=?", archivo, f"Neodata MCP: {motivo}"[:128])
        while cur.nextset():
            pass
    viejos = sorted(glob.glob(os.path.join(CARPETA_RESPALDOS, f"{stem}_*.bak")))[:-MAX_RESPALDOS_POR_OBRA]
    for v in viejos:
        try:
            os.remove(v)
        except OSError:
            pass
    return archivo


@mcp.tool()
def neodata_respaldar(obra: Optional[str] = None) -> dict:
    """Crea un respaldo .bak de la obra en la carpeta de respaldos (también se hace solo antes de cada escritura)."""
    base = _resolver_obra(obra)
    return {"respaldo": _respaldar(base)}


@mcp.tool()
def neodata_listar_respaldos(obra: Optional[str] = None) -> list[dict]:
    """Lista los respaldos .bak hechos por este MCP (todos o los de una obra)."""
    patron = "*.bak"
    if obra:
        stem = re.sub(r"[^\w\-]+", "_", os.path.splitext(os.path.basename(_resolver_obra(obra)))[0]).strip("_")
        patron = f"{stem}_*.bak"
    return [{"archivo": f, "mb": round(os.path.getsize(f) / 1048576, 1),
             "fecha": dt.datetime.fromtimestamp(os.path.getmtime(f)).isoformat(sep=" ", timespec="seconds")}
            for f in sorted(glob.glob(os.path.join(CARPETA_RESPALDOS, patron)), reverse=True)]


@mcp.tool()
def neodata_restaurar_respaldo(archivo: str, obra: Optional[str] = None, confirmar: bool = False) -> dict:
    """Restaura la obra desde un .bak (reemplaza la base actual). Requiere confirmar=True y Neodata cerrado.
    Antes de restaurar hace un respaldo del estado actual."""
    if not confirmar:
        return {"aviso": "Esto reemplaza la obra actual por el respaldo. Repite con confirmar=True tras la aprobación del usuario."}
    if neodata_abierto():
        raise RuntimeError("Cierra Neodata antes de restaurar.")
    base = _resolver_obra(obra)
    archivo = os.path.abspath(archivo)
    if not os.path.isfile(archivo):
        raise ValueError(f"No existe {archivo}")
    previo = _respaldar(base, "antes de restaurar")
    _soltar_ancla(base)
    q = base.replace("]", "]]")
    with _conectar() as cn:
        cn.execute(f"ALTER DATABASE [{q}] SET SINGLE_USER WITH ROLLBACK IMMEDIATE")
        try:
            cur = cn.execute(f"RESTORE DATABASE [{q}] FROM DISK=? WITH REPLACE", archivo)
            while cur.nextset():
                pass
        finally:
            cn.execute(f"ALTER DATABASE [{q}] SET MULTI_USER")
    return {"restaurado_desde": archivo, "respaldo_previo": previo}


if __name__ == "__main__":
    mcp.run()
