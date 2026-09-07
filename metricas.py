"""
============================================================================
 MÉTRICAS Y ANÁLISIS DE AUDITORÍA  ·  Fase 3
============================================================================

Toda la lógica analítica del sistema, separada de la interfaz para poder
mantenerla y probarla sola. La interfaz (Streamlit) llama a estas funciones.

Incluye:
  - Cotizaciones (dólar / IPC) por mes: tabla interna que el usuario carga.
  - Ajuste de precios por dólar e inflación (para comparar en el tiempo).
  - Evolución de precio de un artículo (nominal, en USD, y en pesos reales).
  - Top de artículos por monto / cantidad.
  - Variación de precios mes a mes y detección de saltos.
  - Comparativa de proveedores para un mismo artículo.

Todas las funciones reciben una conexión SQLAlchemy y devuelven DataFrames
de pandas, listos para graficar o mostrar en tabla.
"""

import pandas as pd
from sqlalchemy import text, create_engine


# ---------------------------------------------------------------------------
# CONEXIÓN
# ---------------------------------------------------------------------------

def get_engine(ruta_db="auditoria_compras.db"):
    return create_engine(f"sqlite:///{ruta_db}")


# ---------------------------------------------------------------------------
# TABLA DE COTIZACIONES (dólar / IPC) - se crea sola si no existe
# ---------------------------------------------------------------------------

def asegurar_tabla_cotizaciones(engine):
    """
    Crea la tabla cotizacion_mensual si no existe.
    - mes: texto 'YYYY-MM'
    - dolar: cotización del dólar de referencia para ese mes
    - ipc: índice de precios (nivel del IPC INDEC de ese mes, base cualquiera)
    El usuario la completa una vez por mes desde la interfaz.
    """
    with engine.begin() as c:
        c.execute(text("""
            CREATE TABLE IF NOT EXISTS cotizacion_mensual (
                mes   TEXT PRIMARY KEY,
                dolar REAL,
                ipc   REAL
            )
        """))


def guardar_cotizacion(engine, mes, dolar, ipc):
    """Inserta o actualiza la cotización de un mes."""
    asegurar_tabla_cotizaciones(engine)
    with engine.begin() as c:
        c.execute(text("""
            INSERT INTO cotizacion_mensual (mes, dolar, ipc)
            VALUES (:mes, :dolar, :ipc)
            ON CONFLICT(mes) DO UPDATE SET dolar=:dolar, ipc=:ipc
        """), {"mes": mes, "dolar": dolar, "ipc": ipc})


def leer_cotizaciones(engine):
    asegurar_tabla_cotizaciones(engine)
    return pd.read_sql("SELECT * FROM cotizacion_mensual ORDER BY mes", engine)


# ---------------------------------------------------------------------------
# CATÁLOGO / FILTROS
# ---------------------------------------------------------------------------

def listar_empresas(engine):
    return pd.read_sql("SELECT id_empresa, nombre FROM empresa ORDER BY nombre", engine)


def listar_articulos(engine, id_empresa, buscar=None, limite=200):
    """Lista artículos de una empresa, opcionalmente filtrando por texto."""
    q = """
        SELECT a.id_articulo, a.cod_articulo, a.descripcion, a.rubro,
               COUNT(cl.id_linea) AS n_compras
        FROM articulo a
        LEFT JOIN compra_linea cl ON cl.id_articulo = a.id_articulo
        WHERE a.id_empresa = :emp
    """
    params = {"emp": id_empresa}
    if buscar:
        q += " AND (a.descripcion LIKE :b OR a.cod_articulo LIKE :b)"
        params["b"] = f"%{buscar}%"
    q += " GROUP BY a.id_articulo ORDER BY n_compras DESC LIMIT :lim"
    params["lim"] = limite
    return pd.read_sql(text(q), engine, params=params)


# ---------------------------------------------------------------------------
# TOP DE ARTÍCULOS
# ---------------------------------------------------------------------------

def top_articulos(engine, id_empresa, por="monto", desde=None, hasta=None, limite=15):
    """
    Top de artículos por 'monto' (suma de p_total) o 'cantidad'.
    Filtra por rango de fechas opcional.
    """
    metrica = "SUM(cl.p_total)" if por == "monto" else "SUM(cl.cantidad)"
    q = f"""
        SELECT a.descripcion, a.cod_articulo,
               {metrica} AS valor,
               COUNT(cl.id_linea) AS n_compras,
               AVG(cl.p_unitario) AS precio_prom
        FROM compra_linea cl
        JOIN articulo a ON a.id_articulo = cl.id_articulo
        WHERE cl.id_empresa = :emp
    """
    params = {"emp": id_empresa, "lim": limite}
    if desde:
        q += " AND cl.f_comp >= :desde"; params["desde"] = desde
    if hasta:
        q += " AND cl.f_comp <= :hasta"; params["hasta"] = hasta
    q += " GROUP BY a.id_articulo ORDER BY valor DESC LIMIT :lim"
    return pd.read_sql(text(q), engine, params=params)


# ---------------------------------------------------------------------------
# EVOLUCIÓN DE PRECIO DE UN ARTÍCULO (con ajuste dólar / inflación)
# ---------------------------------------------------------------------------

def evolucion_precio(engine, id_articulo):
    """
    Devuelve la evolución mensual del precio unitario de un artículo:
      - precio_nominal: promedio ponderado por cantidad (el precio real pagado)
      - precio_usd: nominal / dólar del mes (si hay cotización)
      - precio_real: nominal deflactado por IPC a valores del último mes
    El promedio ponderado por cantidad evita que una compra chica a precio
    raro distorsione el mes.
    """
    q = """
        SELECT substr(cl.f_comp,1,7) AS mes,
               SUM(cl.p_unitario * cl.cantidad) AS suma_pond,
               SUM(cl.cantidad) AS suma_cant,
               COUNT(*) AS n
        FROM compra_linea cl
        WHERE cl.id_articulo = :art
          AND cl.f_comp IS NOT NULL
          AND cl.p_unitario IS NOT NULL
          AND cl.cantidad IS NOT NULL AND cl.cantidad > 0
        GROUP BY mes ORDER BY mes
    """
    df = pd.read_sql(text(q), engine, params={"art": id_articulo})
    if df.empty:
        return df
    df["precio_nominal"] = df["suma_pond"] / df["suma_cant"]

    # traer cotizaciones y unir por mes
    cot = leer_cotizaciones(engine)
    df = df.merge(cot, on="mes", how="left")

    # precio en dólares
    df["precio_usd"] = df.apply(
        lambda r: r["precio_nominal"] / r["dolar"] if pd.notna(r.get("dolar")) and r["dolar"] else None,
        axis=1,
    )

    # precio real (deflactado por IPC al último mes disponible con IPC)
    if "ipc" in df.columns and df["ipc"].notna().any():
        ipc_base = df.loc[df["ipc"].notna(), "ipc"].iloc[-1]  # IPC del mes más reciente
        df["precio_real"] = df.apply(
            lambda r: r["precio_nominal"] * (ipc_base / r["ipc"]) if pd.notna(r.get("ipc")) and r["ipc"] else None,
            axis=1,
        )
    else:
        df["precio_real"] = None

    return df[["mes", "precio_nominal", "precio_usd", "precio_real", "suma_cant", "n"]]


# ---------------------------------------------------------------------------
# VARIACIÓN DE PRECIOS: saltos mes a mes
# ---------------------------------------------------------------------------

def variacion_articulos(engine, id_empresa, umbral_pct=20, min_compras=3):
    """
    Detecta artículos con mayor variación de precio entre su primer y último
    mes con compras. Útil para auditar aumentos fuertes.
    Devuelve variación % nominal ordenada de mayor a menor.
    """
    q = """
        SELECT a.id_articulo, a.descripcion, a.cod_articulo,
               substr(cl.f_comp,1,7) AS mes,
               SUM(cl.p_unitario*cl.cantidad)/SUM(cl.cantidad) AS precio,
               COUNT(*) AS n
        FROM compra_linea cl
        JOIN articulo a ON a.id_articulo = cl.id_articulo
        WHERE cl.id_empresa = :emp
          AND cl.f_comp IS NOT NULL AND cl.p_unitario IS NOT NULL
          AND cl.cantidad > 0
        GROUP BY a.id_articulo, mes
    """
    df = pd.read_sql(text(q), engine, params={"emp": id_empresa})
    if df.empty:
        return df

    filas = []
    for art_id, g in df.groupby("id_articulo"):
        g = g.sort_values("mes")
        if len(g) < 2 or g["n"].sum() < min_compras:
            continue
        p_ini, p_fin = g["precio"].iloc[0], g["precio"].iloc[-1]
        if p_ini and p_ini > 0:
            var = (p_fin - p_ini) / p_ini * 100
            filas.append({
                "descripcion": g["descripcion"].iloc[0],
                "cod_articulo": g["cod_articulo"].iloc[0],
                "mes_ini": g["mes"].iloc[0], "precio_ini": round(p_ini, 2),
                "mes_fin": g["mes"].iloc[-1], "precio_fin": round(p_fin, 2),
                "variacion_pct": round(var, 1),
                "meses": len(g),
            })
    res = pd.DataFrame(filas)
    if res.empty:
        return res
    res = res[res["variacion_pct"].abs() >= umbral_pct]
    return res.sort_values("variacion_pct", ascending=False).reset_index(drop=True)


# ---------------------------------------------------------------------------
# COMPARATIVA DE PROVEEDORES PARA UN ARTÍCULO
# ---------------------------------------------------------------------------

def comparar_proveedores(engine, id_articulo):
    """
    Para un artículo, muestra qué precio promedio cobró cada proveedor.
    Sirve para detectar si un proveedor está más caro que otro.
    """
    q = """
        SELECT p.razon_social AS proveedor,
               AVG(cl.p_unitario) AS precio_prom,
               MIN(cl.p_unitario) AS precio_min,
               MAX(cl.p_unitario) AS precio_max,
               SUM(cl.cantidad) AS cant_total,
               COUNT(*) AS n_compras
        FROM compra_linea cl
        JOIN proveedor p ON p.id_proveedor = cl.id_proveedor
        WHERE cl.id_articulo = :art AND cl.p_unitario IS NOT NULL
        GROUP BY p.id_proveedor
        ORDER BY precio_prom ASC
    """
    return pd.read_sql(text(q), engine, params={"art": id_articulo})


# ---------------------------------------------------------------------------
# RESUMEN GENERAL / KPIs
# ---------------------------------------------------------------------------

def kpis_empresa(engine, id_empresa, desde=None, hasta=None):
    """Números gruesos para el tablero: total comprado, nº compras, etc."""
    q = """
        SELECT COUNT(*) AS n_compras,
               SUM(cl.p_total) AS total_comprado,
               COUNT(DISTINCT cl.id_articulo) AS n_articulos,
               COUNT(DISTINCT cl.id_proveedor) AS n_proveedores
        FROM compra_linea cl
        WHERE cl.id_empresa = :emp
    """
    params = {"emp": id_empresa}
    if desde:
        q += " AND cl.f_comp >= :desde"; params["desde"] = desde
    if hasta:
        q += " AND cl.f_comp <= :hasta"; params["hasta"] = hasta
    return pd.read_sql(text(q), engine, params=params).iloc[0].to_dict()


def compras_por_mes(engine, id_empresa):
    """Serie mensual de monto total comprado, para el gráfico del tablero."""
    q = """
        SELECT substr(f_comp,1,7) AS mes,
               SUM(p_total) AS total,
               COUNT(*) AS n_compras
        FROM compra_linea
        WHERE id_empresa = :emp AND f_comp IS NOT NULL
        GROUP BY mes ORDER BY mes
    """
    return pd.read_sql(text(q), engine, params={"emp": id_empresa})


def filas_con_alerta(engine, id_empresa):
    """Trae las líneas marcadas con alertas para revisión."""
    q = """
        SELECT a.descripcion, p.razon_social AS proveedor,
               cl.cantidad, cl.p_unitario, cl.p_total,
               cl.n_comp, cl.f_comp, cl.motivo_alerta
        FROM compra_linea cl
        LEFT JOIN articulo a ON a.id_articulo = cl.id_articulo
        LEFT JOIN proveedor p ON p.id_proveedor = cl.id_proveedor
        WHERE cl.id_empresa = :emp AND cl.motivo_alerta != 'OK'
        ORDER BY cl.f_comp DESC
    """
    return pd.read_sql(text(q), engine, params={"emp": id_empresa})
