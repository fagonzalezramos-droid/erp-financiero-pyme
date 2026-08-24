# -*- coding: utf-8 -*-
"""
ERP Financiero por Proyectos para PYMEs (Chile)
================================================
Aplicación construida en Streamlit + SQLite + Pandas.

Módulos:
  1. Base de datos (SQLite) con carga masiva vía Excel/CSV.
  2. Motor financiero: IVA, retención de Boletas de Honorarios, costo empresa/hora.
  3. Panel de Gerencia: dashboard, alertas, carga masiva, OC, cuentas por cobrar/pagar.
  4. Vista de Terreno (sin login): reporte de gastos y horas vía URL directa
     (agregar ?vista=terreno a la URL de la app).
  5. Flujo de Caja Real consolidado + exportación a Excel.

Ejecutar con:  streamlit run app.py
"""

import os
import io
import sqlite3
from datetime import date, datetime, timedelta

import pandas as pd
import streamlit as st

# ============================================================================
# 0. CONFIGURACIÓN Y CONSTANTES FINANCIERAS
# ============================================================================

st.set_page_config(
    page_title="ERP Financiero PYME — Gestión por Proyectos",
    page_icon="📊",
    layout="wide",
)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, "erp_pyme.db")
UPLOADS_DIR = os.path.join(BASE_DIR, "uploads_terreno")
os.makedirs(UPLOADS_DIR, exist_ok=True)

HORAS_CONTRACTUALES_MES = 180          # Horas contractuales mensuales
IVA_RATE = 0.19                        # IVA 19% vigente en Chile
# Retención de Boletas de Honorarios: escala gradual Ley N°21.133.
# Tasa vigente desde el 01-01-2026: 15,25% (sube 0,75 pts/año hasta 17% en 2028).
RETENCION_HONORARIOS_RATE = 0.1525

TIPOS_DOCUMENTO_PAGAR = [
    "Factura de Compra",
    "Boleta de Terceros",
    "Gasto sin Factura",
    "Boleta de Honorarios",
]

COLS_TRABAJADORES = ["ID", "Nombre", "Rut", "Sueldo Base", "Costo Empresa Mensual"]
COLS_PROYECTOS = [
    "ID", "Nombre", "Presupuesto Total Asignado", "Presupuesto de Egresos",
    "Margen Esperado", "Gastos Acumulados", "Estado",
]
COLS_OC = ["ID", "Proyecto_ID", "Detalle", "Monto Presupuestado por Gerencia", "Estado"]


# ============================================================================
# 1. CAPA DE BASE DE DATOS
# ============================================================================

@st.cache_resource
def get_conn():
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def init_db():
    conn = get_conn()
    cur = conn.cursor()
    cur.executescript(
        """
        CREATE TABLE IF NOT EXISTS trabajadores (
            id INTEGER PRIMARY KEY,
            nombre TEXT NOT NULL,
            rut TEXT,
            sueldo_base REAL DEFAULT 0,
            costo_empresa_mensual REAL DEFAULT 0,
            valor_hora REAL DEFAULT 0
        );

        CREATE TABLE IF NOT EXISTS proyectos (
            id INTEGER PRIMARY KEY,
            nombre TEXT NOT NULL,
            presupuesto_total REAL DEFAULT 0,
            presupuesto_egresos REAL DEFAULT 0,
            margen_esperado REAL DEFAULT 0,
            gastos_acumulados REAL DEFAULT 0,
            estado TEXT DEFAULT 'Activo'
        );

        CREATE TABLE IF NOT EXISTS ordenes_compra (
            id INTEGER PRIMARY KEY,
            proyecto_id INTEGER,
            detalle TEXT,
            monto_presupuestado REAL DEFAULT 0,
            monto_ejecutado REAL DEFAULT 0,
            estado TEXT DEFAULT 'Pendiente'
        );

        CREATE TABLE IF NOT EXISTS cuentas_cobrar (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            proyecto_id INTEGER,
            folio_factura TEXT,
            monto_neto REAL,
            iva REAL,
            monto_total REAL,
            fecha_emision TEXT,
            fecha_vencimiento TEXT,
            estado TEXT DEFAULT 'Pendiente',
            fecha_pago TEXT
        );

        CREATE TABLE IF NOT EXISTS cuentas_pagar (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            proyecto_id INTEGER,
            oc_id INTEGER,
            tipo_documento TEXT,
            folio TEXT,
            monto_neto REAL,
            iva REAL,
            monto_total REAL,
            retencion_honorarios REAL DEFAULT 0,
            liquido_honorarios REAL DEFAULT 0,
            fecha_emision TEXT,
            fecha_vencimiento TEXT,
            estado TEXT DEFAULT 'Pendiente',
            fecha_pago TEXT,
            foto_path TEXT,
            cuadratura TEXT DEFAULT 'Cuadrado'
        );

        CREATE TABLE IF NOT EXISTS horas_trabajadas (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            trabajador_id INTEGER,
            proyecto_id INTEGER,
            fecha TEXT,
            horas REAL,
            costo_calculado REAL
        );
        """
    )
    conn.commit()


def seed_mock_data():
    """Precarga datos ficticios (3 proyectos, 10 trabajadores, 5 OC + movimientos)
    solo si la base de datos está vacía, para que la app funcione de inmediato."""
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT COUNT(*) FROM trabajadores")
    if cur.fetchone()[0] > 0:
        return  # Ya existen datos (propios o precargados en un run anterior)

    hoy = date.today()

    # --- Proyectos ---
    proyectos = [
        (1, "Edificio Los Aromos", 150_000_000, 120_000_000, 20.0, 18_500_000, "Activo"),
        (2, "Bodega Industrial San Bernardo", 80_000_000, 65_000_000, 18.75, 42_000_000, "Activo"),
        (3, "Remodelación Oficinas Las Condes", 45_000_000, 38_000_000, 15.5, 8_200_000, "Activo"),
    ]
    cur.executemany(
        "INSERT INTO proyectos (id, nombre, presupuesto_total, presupuesto_egresos, "
        "margen_esperado, gastos_acumulados, estado) VALUES (?,?,?,?,?,?,?)",
        proyectos,
    )

    # --- Trabajadores (10) ---
    trabajadores_raw = [
        (1, "Juan Pérez Soto", "12.345.678-9", 900_000, 1_260_000),
        (2, "María Fernández Rojas", "13.456.789-0", 1_100_000, 1_540_000),
        (3, "Carlos Muñoz Vega", "14.567.890-1", 850_000, 1_190_000),
        (4, "Ana González Díaz", "15.678.901-2", 950_000, 1_330_000),
        (5, "Pedro Sánchez Lagos", "16.789.012-3", 1_300_000, 1_820_000),
        (6, "Rosa Herrera Contreras", "17.890.123-4", 800_000, 1_120_000),
        (7, "Luis Torres Bravo", "18.901.234-5", 1_050_000, 1_470_000),
        (8, "Camila Ramírez Silva", "19.012.345-6", 920_000, 1_288_000),
        (9, "Diego Castro Morales", "20.123.456-7", 1_400_000, 1_960_000),
        (10, "Francisca Vidal Núñez", "21.234.567-8", 880_000, 1_232_000),
    ]
    trabajadores = [
        (tid, nom, rut, sb, cem, round(cem / HORAS_CONTRACTUALES_MES, 2))
        for (tid, nom, rut, sb, cem) in trabajadores_raw
    ]
    cur.executemany(
        "INSERT INTO trabajadores (id, nombre, rut, sueldo_base, costo_empresa_mensual, "
        "valor_hora) VALUES (?,?,?,?,?,?)",
        trabajadores,
    )

    # --- Órdenes de Compra (5) ---
    ordenes = [
        (101, 1, "Compra de fierro y moldajes", 12_000_000, 0, "Pendiente"),
        (102, 1, "Arriendo de grúa torre", 8_500_000, 0, "Pendiente"),
        (103, 2, "Estructura metálica galpón", 22_000_000, 0, "Pendiente"),
        (104, 2, "Instalaciones eléctricas", 6_000_000, 0, "Pendiente"),
        (105, 3, "Materiales de terminación (pisos/pintura)", 4_500_000, 0, "Pendiente"),
    ]
    cur.executemany(
        "INSERT INTO ordenes_compra (id, proyecto_id, detalle, monto_presupuestado, "
        "monto_ejecutado, estado) VALUES (?,?,?,?,?,?)",
        ordenes,
    )

    # --- Cuentas por Cobrar (facturas a clientes) ---
    cxc = [
        (1, "F-2201", 42_000_000, 42_000_000 * IVA_RATE, 42_000_000 * (1 + IVA_RATE),
         hoy - timedelta(days=40), hoy - timedelta(days=3), "Pendiente", None),
        (1, "F-2214", 30_000_000, 30_000_000 * IVA_RATE, 30_000_000 * (1 + IVA_RATE),
         hoy - timedelta(days=20), hoy + timedelta(days=2), "Pendiente", None),
        (2, "F-1187", 25_000_000, 25_000_000 * IVA_RATE, 25_000_000 * (1 + IVA_RATE),
         hoy - timedelta(days=35), hoy - timedelta(days=15), "Pagado", hoy - timedelta(days=14)),
        (2, "F-1199", 18_000_000, 18_000_000 * IVA_RATE, 18_000_000 * (1 + IVA_RATE),
         hoy - timedelta(days=10), hoy + timedelta(days=4), "Pendiente", None),
        (3, "F-0876", 14_000_000, 14_000_000 * IVA_RATE, 14_000_000 * (1 + IVA_RATE),
         hoy - timedelta(days=8), hoy + timedelta(days=20), "Pendiente", None),
    ]
    cur.executemany(
        "INSERT INTO cuentas_cobrar (proyecto_id, folio_factura, monto_neto, iva, "
        "monto_total, fecha_emision, fecha_vencimiento, estado, fecha_pago) "
        "VALUES (?,?,?,?,?,?,?,?,?)",
        cxc,
    )

    # --- Cuentas por Pagar (gastos/proveedores) ---
    def fc(total):
        neto = round(total / (1 + IVA_RATE))
        return neto, total - neto

    n1, i1 = fc(9_500_000)
    n2, i2 = fc(3_200_000)
    n3, i3 = fc(17_000_000)
    cxp = [
        # proyecto, oc, tipo, folio, neto, iva, total, ret_hon, liq_hon, f_emision, f_venc, estado, f_pago, foto, cuadratura
        (1, 101, "Factura de Compra", "FC-5501", n1, i1, 9_500_000, 0, 0,
         hoy - timedelta(days=15), hoy - timedelta(days=1), "Pendiente", None, None, "Cuadrado"),
        (1, 102, "Boleta de Terceros", "B-330", 3_200_000, 0, 3_200_000, 0, 0,
         hoy - timedelta(days=6), hoy + timedelta(days=3), "Pendiente", None, None, "Cuadrado"),
        (2, 103, "Factura de Compra", "FC-5620", n3, i3, 17_000_000, 0, 0,
         hoy - timedelta(days=25), hoy - timedelta(days=10), "Pagado", hoy - timedelta(days=9), None, "Descuadrado"),
        (3, 105, "Gasto sin Factura", "S/N-021", 480_000, 0, 480_000, 0, 0,
         hoy - timedelta(days=2), hoy + timedelta(days=1), "Pendiente", None, None, "Cuadrado"),
        (1, None, "Boleta de Honorarios", "BH-118", 1_500_000, 0, 1_500_000,
         round(1_500_000 * RETENCION_HONORARIOS_RATE), 1_500_000 - round(1_500_000 * RETENCION_HONORARIOS_RATE),
         hoy - timedelta(days=5), hoy + timedelta(days=25), "Pendiente", None, None, "Cuadrado"),
    ]
    cur.executemany(
        "INSERT INTO cuentas_pagar (proyecto_id, oc_id, tipo_documento, folio, monto_neto, "
        "iva, monto_total, retencion_honorarios, liquido_honorarios, fecha_emision, "
        "fecha_vencimiento, estado, fecha_pago, foto_path, cuadratura) "
        "VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
        cxp,
    )

    # Actualiza monto_ejecutado / estado de las OC según lo cargado arriba
    for oc_id in (101, 102, 103, 105):
        recalcular_oc(oc_id)

    # --- Horas trabajadas (mano de obra acumulada de demo) ---
    horas_demo = [
        (1, 1, hoy - timedelta(days=3), 8), (2, 1, hoy - timedelta(days=3), 8),
        (3, 1, hoy - timedelta(days=2), 9), (4, 2, hoy - timedelta(days=2), 8),
        (5, 2, hoy - timedelta(days=1), 8), (6, 3, hoy - timedelta(days=1), 6),
        (7, 1, hoy - timedelta(days=1), 8), (8, 2, hoy, 4),
        (9, 1, hoy, 8), (10, 3, hoy, 8),
    ]
    valores_hora = {t[0]: t[5] for t in trabajadores}
    horas_rows = [
        (tid, pid, f.isoformat(), h, round(h * valores_hora[tid], 2))
        for (tid, pid, f, h) in horas_demo
    ]
    cur.executemany(
        "INSERT INTO horas_trabajadas (trabajador_id, proyecto_id, fecha, horas, "
        "costo_calculado) VALUES (?,?,?,?,?)",
        horas_rows,
    )

    conn.commit()


# ============================================================================
# 2. MOTOR FINANCIERO (cálculos)
# ============================================================================

def calc_valor_hora(costo_empresa_mensual: float) -> float:
    return round(costo_empresa_mensual / HORAS_CONTRACTUALES_MES, 2)


def calc_factura_compra(monto_total: float):
    """Factura de compra: Neto = Total / 1.19 ; IVA Crédito Fiscal = Total - Neto."""
    neto = round(monto_total / (1 + IVA_RATE))
    iva = round(monto_total - neto)
    return neto, iva


def calc_venta_desde_neto(monto_neto: float):
    """Factura emitida a cliente: IVA Débito Fiscal 19% sobre el neto."""
    iva = round(monto_neto * IVA_RATE)
    total = monto_neto + iva
    return iva, total


def calc_boleta_honorarios(monto_bruto: float):
    retencion = round(monto_bruto * RETENCION_HONORARIOS_RATE)
    liquido = monto_bruto - retencion
    return retencion, liquido


def recalcular_oc(oc_id):
    """Recalcula el monto ejecutado y el estado (Cuadrado/Descuadrado/Pendiente) de una OC
    en base a los gastos de terreno/gerencia asociados a ella."""
    if oc_id is None:
        return
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT monto_presupuestado FROM ordenes_compra WHERE id = ?", (oc_id,))
    row = cur.fetchone()
    if row is None:
        return
    presupuestado = row[0] or 0
    cur.execute(
        "SELECT COALESCE(SUM(monto_total), 0) FROM cuentas_pagar WHERE oc_id = ?", (oc_id,)
    )
    ejecutado = cur.fetchone()[0] or 0

    if ejecutado == 0:
        estado = "Pendiente"
    elif ejecutado > presupuestado:
        estado = "Descuadrado"
    else:
        estado = "Cuadrado"

    cur.execute(
        "UPDATE ordenes_compra SET monto_ejecutado = ?, estado = ? WHERE id = ?",
        (ejecutado, estado, oc_id),
    )
    conn.commit()
    return estado, ejecutado, presupuestado


# ============================================================================
# 3. CONSULTAS / AGREGACIONES PARA DASHBOARD
# ============================================================================

def df_proyectos():
    return pd.read_sql_query("SELECT * FROM proyectos ORDER BY id", get_conn())


def df_trabajadores():
    return pd.read_sql_query("SELECT * FROM trabajadores ORDER BY id", get_conn())


def df_ordenes_compra():
    return pd.read_sql_query(
        "SELECT oc.*, p.nombre AS proyecto_nombre FROM ordenes_compra oc "
        "LEFT JOIN proyectos p ON p.id = oc.proyecto_id ORDER BY oc.id",
        get_conn(),
    )


def df_cuentas_cobrar():
    return pd.read_sql_query(
        "SELECT cc.*, p.nombre AS proyecto_nombre FROM cuentas_cobrar cc "
        "LEFT JOIN proyectos p ON p.id = cc.proyecto_id ORDER BY cc.id DESC",
        get_conn(),
    )


def df_cuentas_pagar():
    return pd.read_sql_query(
        "SELECT cp.*, p.nombre AS proyecto_nombre FROM cuentas_pagar cp "
        "LEFT JOIN proyectos p ON p.id = cp.proyecto_id ORDER BY cp.id DESC",
        get_conn(),
    )


def df_horas():
    return pd.read_sql_query(
        "SELECT h.*, t.nombre AS trabajador_nombre, p.nombre AS proyecto_nombre "
        "FROM horas_trabajadas h "
        "LEFT JOIN trabajadores t ON t.id = h.trabajador_id "
        "LEFT JOIN proyectos p ON p.id = h.proyecto_id ORDER BY h.id DESC",
        get_conn(),
    )


def resumen_financiero_proyectos():
    """Arma la grilla del dashboard: presupuesto, ingresos, egresos, margen, CxC/CxP pendientes."""
    proyectos = df_proyectos()
    cobrar = df_cuentas_cobrar()
    pagar = df_cuentas_pagar()
    horas = df_horas()

    filas = []
    for _, p in proyectos.iterrows():
        pid = p["id"]
        ingresos_facturados = cobrar.loc[cobrar.proyecto_id == pid, "monto_total"].sum()
        cxc_pendiente = cobrar.loc[
            (cobrar.proyecto_id == pid) & (cobrar.estado == "Pendiente"), "monto_total"
        ].sum()
        egresos_compras = pagar.loc[pagar.proyecto_id == pid, "monto_total"].sum()
        cxp_pendiente = pagar.loc[
            (pagar.proyecto_id == pid) & (pagar.estado == "Pendiente"), "monto_total"
        ].sum()
        costo_mano_obra = horas.loc[horas.proyecto_id == pid, "costo_calculado"].sum()

        gasto_real = (p["gastos_acumulados"] or 0) + egresos_compras + costo_mano_obra
        margen_real = ingresos_facturados - gasto_real
        margen_real_pct = (margen_real / ingresos_facturados * 100) if ingresos_facturados else 0.0
        pct_presupuesto = (gasto_real / p["presupuesto_egresos"] * 100) if p["presupuesto_egresos"] else 0.0

        filas.append({
            "ID": pid,
            "Proyecto": p["nombre"],
            "Estado": p["estado"],
            "Presupuesto Asignado": p["presupuesto_total"],
            "Presupuesto Egresos": p["presupuesto_egresos"],
            "Total Ingresos Facturados": ingresos_facturados,
            "Egresos Compras": egresos_compras,
            "Costo Mano de Obra": costo_mano_obra,
            "Gasto Real Total": gasto_real,
            "% Presupuesto Egresos Consumido": round(pct_presupuesto, 1),
            "Margen Real $": margen_real,
            "Margen Real %": round(margen_real_pct, 1),
            "CxC Pendiente": cxc_pendiente,
            "CxP Pendiente": cxp_pendiente,
            "Alerta Presupuesto": pct_presupuesto >= 85,
        })
    return pd.DataFrame(filas)


def alertas_vencimientos(dias=5):
    hoy = date.today()
    limite = hoy + timedelta(days=dias)
    cobrar = df_cuentas_cobrar()
    pagar = df_cuentas_pagar()

    def _prep(df, tipo):
        d = df[df.estado == "Pendiente"].copy()
        if d.empty:
            return d
        d["fecha_vencimiento_dt"] = pd.to_datetime(d["fecha_vencimiento"]).dt.date
        d = d[d["fecha_vencimiento_dt"] <= limite]
        d["tipo"] = tipo
        d["dias_restantes"] = d["fecha_vencimiento_dt"].apply(lambda f: (f - hoy).days)
        return d

    cxc_alert = _prep(cobrar, "Por Cobrar")
    cxp_alert = _prep(pagar, "Por Pagar")
    return cxc_alert, cxp_alert


# ============================================================================
# 4. CARGA MASIVA — PLANTILLAS Y PROCESAMIENTO
# ============================================================================

def generar_plantilla_excel(columnas, nombre_hoja):
    buf = io.BytesIO()
    df = pd.DataFrame(columns=columnas)
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name=nombre_hoja, index=False)
    buf.seek(0)
    return buf


def leer_archivo_subido(uploaded_file):
    nombre = uploaded_file.name.lower()
    if nombre.endswith(".csv"):
        return pd.read_csv(uploaded_file)
    return pd.read_excel(uploaded_file)


def procesar_carga_trabajadores(df):
    faltantes = [c for c in COLS_TRABAJADORES if c not in df.columns]
    if faltantes:
        return False, f"Faltan columnas obligatorias: {', '.join(faltantes)}"

    conn = get_conn()
    cur = conn.cursor()
    n = 0
    for _, r in df.iterrows():
        if pd.isna(r["ID"]):
            continue
        cem = float(r["Costo Empresa Mensual"] or 0)
        valor_hora = calc_valor_hora(cem)
        cur.execute(
            "INSERT INTO trabajadores (id, nombre, rut, sueldo_base, costo_empresa_mensual, valor_hora) "
            "VALUES (?,?,?,?,?,?) "
            "ON CONFLICT(id) DO UPDATE SET nombre=excluded.nombre, rut=excluded.rut, "
            "sueldo_base=excluded.sueldo_base, costo_empresa_mensual=excluded.costo_empresa_mensual, "
            "valor_hora=excluded.valor_hora",
            (int(r["ID"]), str(r["Nombre"]), str(r.get("Rut", "")),
             float(r.get("Sueldo Base", 0) or 0), cem, valor_hora),
        )
        n += 1
    conn.commit()
    return True, f"{n} trabajador(es) cargado(s)/actualizado(s) correctamente."


def procesar_carga_proyectos(df):
    faltantes = [c for c in COLS_PROYECTOS if c not in df.columns]
    if faltantes:
        return False, f"Faltan columnas obligatorias: {', '.join(faltantes)}"

    conn = get_conn()
    cur = conn.cursor()
    n = 0
    for _, r in df.iterrows():
        if pd.isna(r["ID"]):
            continue
        cur.execute(
            "INSERT INTO proyectos (id, nombre, presupuesto_total, presupuesto_egresos, "
            "margen_esperado, gastos_acumulados, estado) VALUES (?,?,?,?,?,?,?) "
            "ON CONFLICT(id) DO UPDATE SET nombre=excluded.nombre, "
            "presupuesto_total=excluded.presupuesto_total, presupuesto_egresos=excluded.presupuesto_egresos, "
            "margen_esperado=excluded.margen_esperado, gastos_acumulados=excluded.gastos_acumulados, "
            "estado=excluded.estado",
            (int(r["ID"]), str(r["Nombre"]),
             float(r.get("Presupuesto Total Asignado", 0) or 0),
             float(r.get("Presupuesto de Egresos", 0) or 0),
             float(r.get("Margen Esperado", 0) or 0),
             float(r.get("Gastos Acumulados", 0) or 0),
             str(r.get("Estado", "Activo") or "Activo")),
        )
        n += 1
    conn.commit()
    return True, f"{n} proyecto(s) cargado(s)/actualizado(s) correctamente."


def procesar_carga_oc(df):
    faltantes = [c for c in COLS_OC if c not in df.columns]
    if faltantes:
        return False, f"Faltan columnas obligatorias: {', '.join(faltantes)}"

    conn = get_conn()
    cur = conn.cursor()
    n = 0
    for _, r in df.iterrows():
        if pd.isna(r["ID"]):
            continue
        cur.execute(
            "INSERT INTO ordenes_compra (id, proyecto_id, detalle, monto_presupuestado, "
            "monto_ejecutado, estado) VALUES (?,?,?,?,0,?) "
            "ON CONFLICT(id) DO UPDATE SET proyecto_id=excluded.proyecto_id, "
            "detalle=excluded.detalle, monto_presupuestado=excluded.monto_presupuestado, "
            "estado=excluded.estado",
            (int(r["ID"]), int(r["Proyecto_ID"]), str(r["Detalle"]),
             float(r.get("Monto Presupuestado por Gerencia", 0) or 0),
             str(r.get("Estado", "Pendiente") or "Pendiente")),
        )
        n += 1
    conn.commit()
    return True, f"{n} orden(es) de compra cargada(s)/actualizada(s) correctamente."


# ============================================================================
# 5. FLUJO DE CAJA REAL (base pagado, no facturado)
# ============================================================================

def flujo_de_caja(fecha_inicio: date, fecha_fin: date):
    conn = get_conn()
    cobrar = pd.read_sql_query(
        "SELECT * FROM cuentas_cobrar WHERE estado = 'Pagado' AND fecha_pago IS NOT NULL", conn
    )
    pagar = pd.read_sql_query(
        "SELECT * FROM cuentas_pagar WHERE estado = 'Pagado' AND fecha_pago IS NOT NULL", conn
    )
    horas = pd.read_sql_query("SELECT * FROM horas_trabajadas", conn)
    proyectos = df_proyectos()

    if not cobrar.empty:
        cobrar["fecha_pago"] = pd.to_datetime(cobrar["fecha_pago"]).dt.date
        cobrar = cobrar[(cobrar.fecha_pago >= fecha_inicio) & (cobrar.fecha_pago <= fecha_fin)]
    if not pagar.empty:
        pagar["fecha_pago"] = pd.to_datetime(pagar["fecha_pago"]).dt.date
        pagar = pagar[(pagar.fecha_pago >= fecha_inicio) & (pagar.fecha_pago <= fecha_fin)]
    if not horas.empty:
        horas["fecha"] = pd.to_datetime(horas["fecha"]).dt.date
        horas = horas[(horas.fecha >= fecha_inicio) & (horas.fecha <= fecha_fin)]

    filas = []
    for _, p in proyectos.iterrows():
        pid = p["id"]
        ing = cobrar.loc[cobrar.proyecto_id == pid, "monto_total"].sum() if not cobrar.empty else 0
        egr_op = pagar.loc[pagar.proyecto_id == pid, "monto_total"].sum() if not pagar.empty else 0
        iva_credito = pagar.loc[pagar.proyecto_id == pid, "iva"].sum() if not pagar.empty else 0
        mo = horas.loc[horas.proyecto_id == pid, "costo_calculado"].sum() if not horas.empty else 0
        neto = ing - egr_op - mo
        filas.append({
            "Proyecto": p["nombre"],
            "(+) Ingresos Cobrados": ing,
            "(-) Egresos Operacionales": egr_op,
            "IVA Crédito Fiscal Acumulado": iva_credito,
            "(-) Costo Mano de Obra": mo,
            "(=) Flujo Neto de Caja": neto,
        })

    df = pd.DataFrame(filas)
    consolidado = {
        "Proyecto": "TOTAL CONSOLIDADO EMPRESA",
        "(+) Ingresos Cobrados": df["(+) Ingresos Cobrados"].sum() if not df.empty else 0,
        "(-) Egresos Operacionales": df["(-) Egresos Operacionales"].sum() if not df.empty else 0,
        "IVA Crédito Fiscal Acumulado": df["IVA Crédito Fiscal Acumulado"].sum() if not df.empty else 0,
        "(-) Costo Mano de Obra": df["(-) Costo Mano de Obra"].sum() if not df.empty else 0,
        "(=) Flujo Neto de Caja": df["(=) Flujo Neto de Caja"].sum() if not df.empty else 0,
    }
    return df, consolidado, cobrar, pagar


def exportar_flujo_excel(df_flujo, consolidado, cobrar, pagar):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df_out = pd.concat([df_flujo, pd.DataFrame([consolidado])], ignore_index=True)
        df_out.to_excel(writer, sheet_name="Flujo de Caja", index=False)
        if not cobrar.empty:
            cobrar.to_excel(writer, sheet_name="Ingresos Cobrados", index=False)
        if not pagar.empty:
            pagar.to_excel(writer, sheet_name="Egresos Pagados", index=False)
        resumen_financiero_proyectos().to_excel(writer, sheet_name="Resumen Proyectos", index=False)
    buf.seek(0)
    return buf


# ============================================================================
# 6. INICIALIZACIÓN
# ============================================================================

init_db()
seed_mock_data()

MONEDA = "$"


def fmt(n):
    try:
        return f"{MONEDA}{n:,.0f}".replace(",", ".")
    except (TypeError, ValueError):
        return f"{MONEDA}0"


# ============================================================================
# 7. VISTA DE TERRENO (sin login — acceso directo vía ?vista=terreno)
# ============================================================================

def vista_terreno():
    st.title("📱 Reporte de Terreno")
    st.caption("Completa tu reporte del día. No necesitas iniciar sesión.")

    tab_gasto, tab_horas = st.tabs(["🧾 Reportar Compra / Gasto", "⏱️ Reportar Horas del Día"])

    proyectos = df_proyectos()
    proyectos_activos = proyectos[proyectos.estado == "Activo"]

    # ---- FORMULARIO DE COMPRAS / EGRESOS ----
    with tab_gasto:
        if proyectos_activos.empty:
            st.warning("No hay proyectos activos cargados todavía.")
        else:
            proyecto_sel = st.selectbox(
                "Proyecto", proyectos_activos["nombre"].tolist(), key="terreno_proy_gasto"
            )
            proyecto_id = int(proyectos_activos.loc[proyectos_activos.nombre == proyecto_sel, "id"].iloc[0])

            ocs = df_ordenes_compra()
            ocs_proy = ocs[ocs.proyecto_id == proyecto_id]
            if ocs_proy.empty:
                st.info("Este proyecto no tiene Órdenes de Compra cargadas. Puedes reportar igualmente sin OC.")
                oc_id_sel = None
            else:
                opciones_oc = ["(Sin OC asociada)"] + [
                    f"OC-{row.id} · {row.detalle} (Presupuesto: {fmt(row.monto_presupuestado)})"
                    for _, row in ocs_proy.iterrows()
                ]
                oc_elegida = st.selectbox("Orden de Compra (OC)", opciones_oc, key="terreno_oc")
                oc_id_sel = None
                if oc_elegida != "(Sin OC asociada)":
                    oc_id_sel = int(oc_elegida.split("·")[0].replace("OC-", "").strip())

            with st.form("form_gasto_terreno", clear_on_submit=True):
                monto_total = st.number_input("Monto total gastado ($)", min_value=0, step=1000)
                tipo_doc = st.selectbox("Tipo de Documento", TIPOS_DOCUMENTO_PAGAR)
                folio = st.text_input("N° de Folio / Boleta")
                foto = st.file_uploader("Foto de la boleta o factura", type=["jpg", "jpeg", "png", "pdf"])
                enviar = st.form_submit_button("Enviar Reporte de Gasto", use_container_width=True)

            if enviar:
                if monto_total <= 0:
                    st.error("Ingresa un monto válido mayor a 0.")
                else:
                    # Cálculo automático de IVA / retención según tipo de documento
                    neto, iva, ret_hon, liq_hon = monto_total, 0, 0, 0
                    if tipo_doc == "Factura de Compra":
                        neto, iva = calc_factura_compra(monto_total)
                    elif tipo_doc == "Boleta de Honorarios":
                        ret_hon, liq_hon = calc_boleta_honorarios(monto_total)
                        neto = monto_total
                    else:  # Boleta de Terceros / Gasto sin Factura
                        neto = monto_total

                    foto_path = None
                    if foto is not None:
                        foto_path = os.path.join(
                            UPLOADS_DIR, f"{datetime.now():%Y%m%d%H%M%S}_{foto.name}"
                        )
                        with open(foto_path, "wb") as f:
                            f.write(foto.getbuffer())

                    conn = get_conn()
                    cur = conn.cursor()
                    cur.execute(
                        "INSERT INTO cuentas_pagar (proyecto_id, oc_id, tipo_documento, folio, "
                        "monto_neto, iva, monto_total, retencion_honorarios, liquido_honorarios, "
                        "fecha_emision, fecha_vencimiento, estado, fecha_pago, foto_path, cuadratura) "
                        "VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                        (proyecto_id, oc_id_sel, tipo_doc, folio, neto, iva, monto_total,
                         ret_hon, liq_hon, date.today().isoformat(), date.today().isoformat(),
                         "Pendiente", None, foto_path, "Cuadrado"),
                    )
                    conn.commit()

                    cuadratura = "Cuadrado"
                    if oc_id_sel is not None:
                        cuadratura, ejecutado, presupuestado = recalcular_oc(oc_id_sel)
                        cur.execute(
                            "UPDATE cuentas_pagar SET cuadratura = ? WHERE id = (SELECT MAX(id) FROM cuentas_pagar)",
                            (cuadratura,),
                        )
                        conn.commit()

                    if oc_id_sel is not None and cuadratura == "Descuadrado":
                        st.warning(
                            "⚠️ Reporte enviado, pero el monto **excede el presupuesto de la OC**. "
                            "Se notificó a Gerencia como 'Descuadrado'."
                        )
                    else:
                        st.success("✅ Gasto reportado correctamente.")

    # ---- FORMULARIO DE HORAS DIARIAS ----
    with tab_horas:
        trabajadores = df_trabajadores()
        if trabajadores.empty or proyectos_activos.empty:
            st.warning("No hay trabajadores o proyectos activos cargados todavía.")
        else:
            trabajador_sel = st.selectbox("Trabajador", trabajadores["nombre"].tolist(), key="terreno_trab")
            trabajador_id = int(trabajadores.loc[trabajadores.nombre == trabajador_sel, "id"].iloc[0])
            valor_hora = float(trabajadores.loc[trabajadores.nombre == trabajador_sel, "valor_hora"].iloc[0])
            st.caption(f"Valor Hora Costo Empresa: **{fmt(valor_hora)}**")

            st.write("Registra las horas trabajadas hoy, distribuidas por proyecto:")
            editor_df = pd.DataFrame({"Proyecto": [proyectos_activos["nombre"].iloc[0]], "Horas": [8.0]})
            horas_editadas = st.data_editor(
                editor_df,
                num_rows="dynamic",
                use_container_width=True,
                column_config={
                    "Proyecto": st.column_config.SelectboxColumn(
                        "Proyecto", options=proyectos_activos["nombre"].tolist(), required=True
                    ),
                    "Horas": st.column_config.NumberColumn("Horas", min_value=0.0, max_value=24.0, step=0.5),
                },
                key="editor_horas",
            )

            if st.button("Enviar Reporte de Horas", use_container_width=True):
                filas_validas = horas_editadas.dropna(subset=["Proyecto", "Horas"])
                filas_validas = filas_validas[filas_validas["Horas"] > 0]
                if filas_validas.empty:
                    st.error("Ingresa al menos una fila con horas mayores a 0.")
                else:
                    conn = get_conn()
                    cur = conn.cursor()
                    total_costo = 0.0
                    for _, fila in filas_validas.iterrows():
                        pid = int(proyectos_activos.loc[proyectos_activos.nombre == fila["Proyecto"], "id"].iloc[0])
                        costo = round(float(fila["Horas"]) * valor_hora, 2)
                        total_costo += costo
                        cur.execute(
                            "INSERT INTO horas_trabajadas (trabajador_id, proyecto_id, fecha, horas, "
                            "costo_calculado) VALUES (?,?,?,?,?)",
                            (trabajador_id, pid, date.today().isoformat(), float(fila["Horas"]), costo),
                        )
                    conn.commit()
                    st.success(
                        f"✅ Horas registradas correctamente. Costo total del día: **{fmt(total_costo)}**"
                    )


# ============================================================================
# 8. VISTA DE GERENCIA
# ============================================================================

def vista_gerencia():
    st.title("📊 Panel de Gerencia — Gestión Financiera por Proyectos")

    url_terreno = "Agrega **`?vista=terreno`** al final de la URL de esta app y compártela con los operarios."
    with st.expander("🔗 Link directo para Terreno (sin login)"):
        st.info(url_terreno)

    tabs = st.tabs([
        "🏠 Dashboard", "📥 Carga Masiva", "🧾 Órdenes de Compra",
        "💰 Cuentas por Cobrar", "💸 Cuentas por Pagar", "📈 Flujo de Caja",
    ])

    # ---------------- DASHBOARD ----------------
    with tabs[0]:
        resumen = resumen_financiero_proyectos()
        cxc_alert, cxp_alert = alertas_vencimientos(dias=5)
        ocs = df_ordenes_compra()
        oc_descuadradas = ocs[ocs.estado == "Descuadrado"]

        criticos = resumen[resumen["Alerta Presupuesto"]]
        col1, col2, col3 = st.columns(3)
        col1.metric("Proyectos Activos", int((df_proyectos().estado == "Activo").sum()))
        col2.metric("Ingresos Facturados (Total)", fmt(resumen["Total Ingresos Facturados"].sum()))
        col3.metric("Gasto Real (Total)", fmt(resumen["Gasto Real Total"].sum()))

        st.subheader("🚨 Alertas de Tesorería")
        acol1, acol2, acol3 = st.columns(3)

        with acol1:
            st.markdown("**Presupuesto por Proyecto**")
            if criticos.empty:
                st.success("Sin proyectos en estado crítico.")
            for _, r in criticos.iterrows():
                st.error(
                    f"🔴 Crítico: Presupuesto por Agotarse — **{r['Proyecto']}** "
                    f"({r['% Presupuesto Egresos Consumido']}% del presupuesto de egresos)"
                )

        with acol2:
            st.markdown("**Vencimientos (Cobrar / Pagar, próx. 5 días o vencidas)**")
            if cxc_alert.empty and cxp_alert.empty:
                st.success("Sin vencimientos próximos.")
            for _, r in cxc_alert.iterrows():
                etiqueta = "VENCIDA" if r["dias_restantes"] < 0 else f"vence en {r['dias_restantes']} día(s)"
                st.warning(f"💰 Por Cobrar {r['folio_factura']} ({r['proyecto_nombre']}) — {etiqueta} — {fmt(r['monto_total'])}")
            for _, r in cxp_alert.iterrows():
                etiqueta = "VENCIDA" if r["dias_restantes"] < 0 else f"vence en {r['dias_restantes']} día(s)"
                st.warning(f"💸 Por Pagar {r['folio']} ({r['proyecto_nombre']}) — {etiqueta} — {fmt(r['monto_total'])}")

        with acol3:
            st.markdown("**Órdenes de Compra Descuadradas**")
            if oc_descuadradas.empty:
                st.success("Todas las OC están cuadradas.")
            for _, r in oc_descuadradas.iterrows():
                st.error(
                    f"🔴 OC-{r['id']} ({r['proyecto_nombre']}): ejecutado {fmt(r['monto_ejecutado'])} "
                    f"vs. presupuestado {fmt(r['monto_presupuestado'])}"
                )

        st.subheader("📋 Dashboard Financiero de Proyectos")
        tabla = resumen.drop(columns=["Alerta Presupuesto"]).copy()
        for c in ["Presupuesto Asignado", "Presupuesto Egresos", "Total Ingresos Facturados",
                  "Egresos Compras", "Costo Mano de Obra", "Gasto Real Total",
                  "Margen Real $", "CxC Pendiente", "CxP Pendiente"]:
            tabla[c] = tabla[c].apply(fmt)
        st.dataframe(tabla, use_container_width=True, hide_index=True)

    # ---------------- CARGA MASIVA ----------------
    with tabs[1]:
        st.subheader("📥 Carga Masiva mediante Excel / CSV")
        st.caption("Descarga la plantilla, complétala y súbela. Se valida el formato de columnas exacto.")

        c1, c2, c3 = st.columns(3)
        with c1:
            st.markdown("#### Trabajadores")
            st.download_button(
                "⬇️ Descargar plantilla", generar_plantilla_excel(COLS_TRABAJADORES, "Trabajadores"),
                file_name="plantilla_trabajadores.xlsx", key="dl_trab",
            )
            up = st.file_uploader("Subir archivo de Trabajadores", type=["xlsx", "csv"], key="up_trab")
            if up is not None and st.button("Procesar Trabajadores", key="proc_trab"):
                try:
                    ok, msg = procesar_carga_trabajadores(leer_archivo_subido(up))
                    st.success(msg) if ok else st.error(msg)
                except Exception as e:
                    st.error(f"Error al procesar el archivo: {e}")

        with c2:
            st.markdown("#### Proyectos")
            st.download_button(
                "⬇️ Descargar plantilla", generar_plantilla_excel(COLS_PROYECTOS, "Proyectos"),
                file_name="plantilla_proyectos.xlsx", key="dl_proy",
            )
            up2 = st.file_uploader("Subir archivo de Proyectos", type=["xlsx", "csv"], key="up_proy")
            if up2 is not None and st.button("Procesar Proyectos", key="proc_proy"):
                try:
                    ok, msg = procesar_carga_proyectos(leer_archivo_subido(up2))
                    st.success(msg) if ok else st.error(msg)
                except Exception as e:
                    st.error(f"Error al procesar el archivo: {e}")

        with c3:
            st.markdown("#### Órdenes de Compra")
            st.download_button(
                "⬇️ Descargar plantilla", generar_plantilla_excel(COLS_OC, "Ordenes_Compra"),
                file_name="plantilla_ordenes_compra.xlsx", key="dl_oc",
            )
            up3 = st.file_uploader("Subir archivo de OC", type=["xlsx", "csv"], key="up_oc")
            if up3 is not None and st.button("Procesar Órdenes de Compra", key="proc_oc"):
                try:
                    ok, msg = procesar_carga_oc(leer_archivo_subido(up3))
                    st.success(msg) if ok else st.error(msg)
                except Exception as e:
                    st.error(f"Error al procesar el archivo: {e}")

    # ---------------- ÓRDENES DE COMPRA ----------------
    with tabs[2]:
        st.subheader("🧾 Ingreso Manual de Orden de Compra")
        proyectos = df_proyectos()
        with st.form("form_oc_manual", clear_on_submit=True):
            colA, colB = st.columns(2)
            with colA:
                proyecto_sel = st.selectbox("Proyecto", proyectos["nombre"].tolist())
                detalle = st.text_input("Detalle de la OC")
            with colB:
                nuevo_id = st.number_input(
                    "Código de OC", min_value=1,
                    value=int(df_ordenes_compra()["id"].max() + 1) if not df_ordenes_compra().empty else 1,
                )
                monto_max = st.number_input("Monto máximo aprobado ($)", min_value=0, step=1000)
            crear = st.form_submit_button("Crear Orden de Compra")

        if crear:
            proyecto_id = int(proyectos.loc[proyectos.nombre == proyecto_sel, "id"].iloc[0])
            conn = get_conn()
            cur = conn.cursor()
            try:
                cur.execute(
                    "INSERT INTO ordenes_compra (id, proyecto_id, detalle, monto_presupuestado, "
                    "monto_ejecutado, estado) VALUES (?,?,?,?,0,'Pendiente')",
                    (int(nuevo_id), proyecto_id, detalle, float(monto_max)),
                )
                conn.commit()
                st.success(f"OC-{int(nuevo_id)} creada correctamente.")
            except sqlite3.IntegrityError:
                st.error("Ya existe una OC con ese código. Elige otro.")

        st.subheader("Listado de Órdenes de Compra")
        oc_tabla = df_ordenes_compra().copy()
        for c in ["monto_presupuestado", "monto_ejecutado"]:
            oc_tabla[c] = oc_tabla[c].apply(fmt)
        st.dataframe(oc_tabla, use_container_width=True, hide_index=True)

    # ---------------- CUENTAS POR COBRAR ----------------
    with tabs[3]:
        st.subheader("💰 Nueva Factura por Cobrar")
        proyectos = df_proyectos()
        with st.form("form_cxc", clear_on_submit=True):
            c1, c2, c3 = st.columns(3)
            with c1:
                proyecto_sel = st.selectbox("Proyecto", proyectos["nombre"].tolist(), key="cxc_proy")
                folio = st.text_input("Folio Factura")
            with c2:
                monto_neto = st.number_input("Monto Neto ($)", min_value=0, step=1000)
                f_emision = st.date_input("Fecha Emisión", value=date.today())
            with c3:
                f_vencimiento = st.date_input("Fecha Vencimiento", value=date.today() + timedelta(days=30))
            enviar = st.form_submit_button("Registrar Factura")

        if enviar:
            iva, total = calc_venta_desde_neto(monto_neto)
            proyecto_id = int(proyectos.loc[proyectos.nombre == proyecto_sel, "id"].iloc[0])
            conn = get_conn()
            conn.execute(
                "INSERT INTO cuentas_cobrar (proyecto_id, folio_factura, monto_neto, iva, monto_total, "
                "fecha_emision, fecha_vencimiento, estado, fecha_pago) VALUES (?,?,?,?,?,?,?,?,?)",
                (proyecto_id, folio, monto_neto, iva, total, f_emision.isoformat(),
                 f_vencimiento.isoformat(), "Pendiente", None),
            )
            conn.commit()
            st.success(f"Factura registrada. Neto {fmt(monto_neto)} + IVA {fmt(iva)} = Total {fmt(total)}")

        st.subheader("Gestión de Cuentas por Cobrar")
        cxc = df_cuentas_cobrar()
        if cxc.empty:
            st.info("No hay cuentas por cobrar registradas.")
        else:
            editable = cxc[["id", "proyecto_nombre", "folio_factura", "monto_total",
                             "fecha_vencimiento", "estado", "fecha_pago"]].copy()
            editado = st.data_editor(
                editable,
                use_container_width=True,
                hide_index=True,
                disabled=["id", "proyecto_nombre", "folio_factura", "monto_total", "fecha_vencimiento"],
                column_config={
                    "estado": st.column_config.SelectboxColumn("Estado", options=["Pendiente", "Pagado"]),
                    "fecha_pago": st.column_config.TextColumn("Fecha Pago (AAAA-MM-DD)"),
                },
                key="editor_cxc",
            )
            if st.button("💾 Guardar cambios en Cuentas por Cobrar"):
                conn = get_conn()
                cur = conn.cursor()
                for _, r in editado.iterrows():
                    fecha_pago = r["fecha_pago"] if r["estado"] == "Pagado" and r["fecha_pago"] else \
                        (date.today().isoformat() if r["estado"] == "Pagado" else None)
                    cur.execute(
                        "UPDATE cuentas_cobrar SET estado = ?, fecha_pago = ? WHERE id = ?",
                        (r["estado"], fecha_pago, int(r["id"])),
                    )
                conn.commit()
                st.success("Cambios guardados correctamente.")
                st.rerun()

    # ---------------- CUENTAS POR PAGAR ----------------
    with tabs[4]:
        st.subheader("Gestión de Cuentas por Pagar")
        st.caption("Se pueblan automáticamente desde los reportes de Terreno. También puedes registrar manualmente.")

        with st.expander("➕ Registrar Cuenta por Pagar manual"):
            proyectos = df_proyectos()
            with st.form("form_cxp_manual", clear_on_submit=True):
                c1, c2, c3 = st.columns(3)
                with c1:
                    proyecto_sel = st.selectbox("Proyecto", proyectos["nombre"].tolist(), key="cxp_proy")
                    tipo_doc = st.selectbox("Tipo de Documento", TIPOS_DOCUMENTO_PAGAR, key="cxp_tipo")
                with c2:
                    folio = st.text_input("Folio")
                    monto_total = st.number_input("Monto Total ($)", min_value=0, step=1000, key="cxp_monto")
                with c3:
                    f_emision = st.date_input("Fecha Emisión", value=date.today(), key="cxp_fem")
                    f_vencimiento = st.date_input("Fecha Vencimiento", value=date.today() + timedelta(days=15), key="cxp_fven")
                enviar_cxp = st.form_submit_button("Registrar")

            if enviar_cxp:
                neto, iva, ret_hon, liq_hon = monto_total, 0, 0, 0
                if tipo_doc == "Factura de Compra":
                    neto, iva = calc_factura_compra(monto_total)
                elif tipo_doc == "Boleta de Honorarios":
                    ret_hon, liq_hon = calc_boleta_honorarios(monto_total)
                proyecto_id = int(proyectos.loc[proyectos.nombre == proyecto_sel, "id"].iloc[0])
                conn = get_conn()
                conn.execute(
                    "INSERT INTO cuentas_pagar (proyecto_id, oc_id, tipo_documento, folio, monto_neto, "
                    "iva, monto_total, retencion_honorarios, liquido_honorarios, fecha_emision, "
                    "fecha_vencimiento, estado, fecha_pago, foto_path, cuadratura) "
                    "VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                    (proyecto_id, None, tipo_doc, folio, neto, iva, monto_total, ret_hon, liq_hon,
                     f_emision.isoformat(), f_vencimiento.isoformat(), "Pendiente", None, None, "Cuadrado"),
                )
                conn.commit()
                st.success("Cuenta por pagar registrada correctamente.")
                st.rerun()

        cxp = df_cuentas_pagar()
        if cxp.empty:
            st.info("No hay cuentas por pagar registradas.")
        else:
            editable = cxp[["id", "proyecto_nombre", "tipo_documento", "folio", "monto_total",
                             "cuadratura", "fecha_vencimiento", "estado", "fecha_pago"]].copy()
            editado = st.data_editor(
                editable,
                use_container_width=True,
                hide_index=True,
                disabled=["id", "proyecto_nombre", "tipo_documento", "folio", "monto_total",
                          "cuadratura", "fecha_vencimiento"],
                column_config={
                    "estado": st.column_config.SelectboxColumn("Estado", options=["Pendiente", "Pagado"]),
                    "fecha_pago": st.column_config.TextColumn("Fecha Pago (AAAA-MM-DD)"),
                },
                key="editor_cxp",
            )
            if st.button("💾 Guardar cambios en Cuentas por Pagar"):
                conn = get_conn()
                cur = conn.cursor()
                for _, r in editado.iterrows():
                    fecha_pago = r["fecha_pago"] if r["estado"] == "Pagado" and r["fecha_pago"] else \
                        (date.today().isoformat() if r["estado"] == "Pagado" else None)
                    cur.execute(
                        "UPDATE cuentas_pagar SET estado = ?, fecha_pago = ? WHERE id = ?",
                        (r["estado"], fecha_pago, int(r["id"])),
                    )
                conn.commit()
                st.success("Cambios guardados correctamente.")
                st.rerun()

    # ---------------- FLUJO DE CAJA ----------------
    with tabs[5]:
        st.subheader("📈 Flujo de Caja Real por Proyectos (base pagado)")
        c1, c2 = st.columns(2)
        with c1:
            f_ini = st.date_input("Desde", value=date.today().replace(day=1))
        with c2:
            f_fin = st.date_input("Hasta", value=date.today())

        df_flujo, consolidado, cobrar_periodo, pagar_periodo = flujo_de_caja(f_ini, f_fin)

        st.markdown("#### Por Proyecto")
        tabla_flujo = df_flujo.copy()
        for c in tabla_flujo.columns:
            if c != "Proyecto":
                tabla_flujo[c] = tabla_flujo[c].apply(fmt)
        st.dataframe(tabla_flujo, use_container_width=True, hide_index=True)

        st.markdown("#### Consolidado de la Empresa")
        cons_col = st.columns(5)
        cons_col[0].metric("(+) Ingresos Cobrados", fmt(consolidado["(+) Ingresos Cobrados"]))
        cons_col[1].metric("(-) Egresos Operacionales", fmt(consolidado["(-) Egresos Operacionales"]))
        cons_col[2].metric("IVA Crédito Fiscal", fmt(consolidado["IVA Crédito Fiscal Acumulado"]))
        cons_col[3].metric("(-) Costo Mano de Obra", fmt(consolidado["(-) Costo Mano de Obra"]))
        cons_col[4].metric("(=) Flujo Neto de Caja", fmt(consolidado["(=) Flujo Neto de Caja"]))

        excel_buf = exportar_flujo_excel(df_flujo, consolidado, cobrar_periodo, pagar_periodo)
        st.download_button(
            "⬇️ Exportar Flujo de Caja a Excel",
            data=excel_buf,
            file_name=f"flujo_caja_{f_ini}_{f_fin}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


# ============================================================================
# 9. ENRUTAMIENTO PRINCIPAL
# ============================================================================

def main():
    params = st.query_params
    vista = params.get("vista", "gerencia")
    if isinstance(vista, list):  # compatibilidad versiones antiguas de Streamlit
        vista = vista[0] if vista else "gerencia"

    if vista == "terreno":
        vista_terreno()
    else:
        vista_gerencia()


if __name__ == "__main__":
    main()
