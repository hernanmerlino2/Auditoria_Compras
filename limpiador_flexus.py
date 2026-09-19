"""
============================================================================
 LIMPIADOR DE EXPORTS DE FLEXUS  ·  Fase 1
 Sistema de Auditoría de Compras - Fly Kitchen S.A. / Drill S.A. / etc.
============================================================================

Qué hace este script:
  1. Lee un export .xlsx de Flexus (estructura de "Estadísticas").
  2. VALIDA que el archivo tenga la estructura esperada (21 columnas correctas).
     Si no coincide, RECHAZA el archivo y avisa (no toca nada).
  3. LIMPIA la basura estructural del export:
        - Filas de encabezado repetidas ('Proveedor', 'Cuit', ...).
        - Filas de fecha de paginación (una fecha sola en la 1ra columna).
  4. RUTEA las líneas cuyo Cod. Artículo es '*' a un circuito separado
     de GASTOS VARIOS (fondos fijos, rendiciones, materiales sueltos),
     porque no sirven para comparar precios de artículos.
  5. NORMALIZA tipos: números a float, fechas dd/mm/yyyy a fecha ISO,
     texto sin espacios sobrantes ni caracteres de control (_x000B_).
  6. Agrega columnas de control: id_empresa y hash de la línea (clave de
     negocio, para detectar duplicados al recargar el mismo archivo).
  7. Devuelve un REPORTE de lo que pasó: cuántas filas entraron, cuántas
     se limpiaron, cuántas fueron a gastos, cuántas quedaron listas.

Uso desde consola:
    python limpiador_flexus.py archivo.xlsx --empresa "Fly Kitchen S.A." --id-empresa 1

Salida:
    - <archivo>_LIMPIO.xlsx        (artículos reales, listos para la BD)
    - <archivo>_GASTOS.xlsx        (líneas con código '*')
    - <archivo>_RECHAZADAS.xlsx    (filas que no se pudieron procesar)
    - Reporte impreso en pantalla + <archivo>_REPORTE.txt
"""

import sys
import csv
import argparse
import hashlib
import unicodedata
from datetime import datetime
from pathlib import Path

import openpyxl
import pandas as pd


# ---------------------------------------------------------------------------
# LECTOR UNIFICADO: CSV (formato real de Flexus) o Excel (reporte interno)
# ---------------------------------------------------------------------------

def leer_filas(ruta):
    """
    Lee un archivo de Flexus y devuelve (header, filas) como listas de tuplas,
    sin importar si es CSV o Excel.

    - CSV: es el formato REAL que exporta Flexus. Separador ';', encoding
      Latin-1 (típico de sistemas Windows en Argentina), decimales con coma,
      y el encabezado en la PRIMERA fila.
    - Excel: es el reporte interno (PDF/Excel), con el encabezado en la fila 4
      y filas de basura (paginación, headers repetidos).

    Devuelve también 'fila_header' = índice (0-based) donde está el encabezado,
    para que el resto del proceso sepa desde dónde son datos.
    """
    ruta = Path(ruta)
    ext = ruta.suffix.lower()

    if ext == ".csv":
        # Detectar encoding: primero latin-1 (lo más común en Flexus), luego utf-8
        for enc in ("latin-1", "utf-8-sig", "cp1252"):
            try:
                with open(ruta, encoding=enc, newline="") as f:
                    filas = list(csv.reader(f, delimiter=";"))
                break
            except (UnicodeDecodeError, UnicodeError):
                continue
        else:
            raise ValueError("No se pudo leer el CSV con ningún encoding conocido.")
        # limpiar filas vacías del final y columna vacía sobrante por ';' final
        filas = [tuple(c for c in fila) for fila in filas if any(x.strip() for x in fila)]
        if not filas:
            raise ValueError("El CSV está vacío.")
        return filas, 0   # header en la fila 0

    else:  # Excel
        wb = openpyxl.load_workbook(ruta, read_only=True, data_only=True)
        ws = wb.active
        filas = [tuple(fila) for fila in ws.iter_rows(values_only=True)]
        wb.close()
        # en el reporte interno el header está en la fila 4 (índice 3)
        return filas, FILA_HEADER - 1



# ---------------------------------------------------------------------------
# CONFIGURACIÓN: la estructura que DEBE tener un export válido de Flexus
# ---------------------------------------------------------------------------

# Encabezado esperado (fila 4 del export). El orden importa.
COLUMNAS_ESPERADAS = [
    "Proveedor", "Cuit", "Cod. Artículo", "Artículo", "Cod. Rubro", "Rubro",
    "Marca", "U. Medida", "Cantidad", "P. Unitario", "Precio Compra",
    "P. Total", "Peso Artículo (Kg)", "Volumen Artículo (cm3)",
    "Peso Artículo Comp.", "Comentarios", "T. Comp.", "N.Comp.", "F.Comp.",
    "Codigo Depósito", "Descripción Depósito",
]

# Nombres internos "limpios" (los que van a la BD).
COLUMNAS_BD = [
    "proveedor", "cuit", "cod_articulo", "articulo", "cod_rubro", "rubro",
    "marca", "u_medida", "cantidad", "p_unitario", "precio_compra",
    "p_total", "peso_articulo_kg", "volumen_articulo_cm3",
    "peso_articulo_comp", "comentarios", "t_comp", "n_comp", "f_comp",
    "cod_deposito", "desc_deposito", "item_flexxus",
]

# Mapa: nombre de columna en el export de Flexus  ->  nombre interno de BD.
# IMPORTANTE: Flexus puede exportar las 21 columnas EN DISTINTO ORDEN según
# cómo esté configurado el reporte (Fly Kitchen y Drill difieren en el orden).
# Por eso mapeamos por NOMBRE y no por posición. La clave se normaliza
# (minúsculas, sin acentos) al comparar, así que acentos/mayúsculas no rompen.
MAPA_FLEXUS_A_BD = {
    "proveedor": "proveedor",
    "cuit": "cuit",
    "cod. articulo": "cod_articulo",
    "articulo": "articulo",
    "cod. rubro": "cod_rubro",
    "rubro": "rubro",
    "marca": "marca",
    "u. medida": "u_medida",
    "cantidad": "cantidad",
    "p. unitario": "p_unitario",
    "precio compra": "precio_compra",
    "p. total": "p_total",
    "peso articulo (kg)": "peso_articulo_kg",
    "volumen articulo (cm3)": "volumen_articulo_cm3",
    "peso articulo comp.": "peso_articulo_comp",
    "comentarios": "comentarios",
    "t. comp.": "t_comp",
    "n.comp.": "n_comp",
    "f.comp.": "f_comp",
    "codigo deposito": "cod_deposito",
    "descripcion deposito": "desc_deposito",
    "itemflexxus": "item_flexxus",
}

# Columnas que pueden faltar sin que el archivo se rechace (según el export).
# ITEMFLEXXUS solo viene en el CSV real de Flexus, no en el reporte interno.
COLUMNAS_OPCIONALES = {"item_flexxus"}

# Columnas que deben ser numéricas.
COLS_NUMERICAS = [
    "cantidad", "p_unitario", "precio_compra", "p_total",
    "peso_articulo_kg", "volumen_articulo_cm3", "peso_articulo_comp",
]

FILA_HEADER = 4          # el encabezado real está en la fila 4
CODIGO_GASTO = "*"       # marca de gasto vario / no comparable


# ---------------------------------------------------------------------------
# UTILIDADES DE NORMALIZACIÓN
# ---------------------------------------------------------------------------

def _norm(s):
    """minúsculas + sin acentos, para comparar etiquetas de forma robusta."""
    if s is None:
        return ""
    s = str(s).strip().lower()
    return "".join(c for c in unicodedata.normalize("NFD", s)
                   if unicodedata.category(c) != "Mn")


def limpiar_texto(valor):
    """Quita espacios sobrantes y caracteres de control tipo _x000B_."""
    if valor is None:
        return None
    s = str(valor)
    # openpyxl a veces deja el literal _x000B_ (salto de línea vertical)
    s = s.replace("_x000B_", " ").replace("\x0b", " ")
    # normaliza espacios internos y bordes
    s = " ".join(s.split())
    return s if s != "" else None


def a_numero(valor):
    """Convierte a float. Vacío o inválido -> None."""
    if valor is None or valor == "":
        return None
    if isinstance(valor, (int, float)):
        return float(valor)
    s = str(valor).strip().replace(".", "").replace(",", ".") \
        if "," in str(valor) else str(valor).strip()
    try:
        return float(s)
    except ValueError:
        return None


def a_fecha_iso(valor):
    """dd/mm/yyyy -> yyyy-mm-dd. Si ya es fecha, la formatea. Si no puede, None."""
    if valor is None or valor == "":
        return None
    if isinstance(valor, datetime):
        return valor.strftime("%Y-%m-%d")
    s = str(valor).strip()
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"):
        try:
            return datetime.strptime(s, fmt).strftime("%Y-%m-%d")
        except ValueError:
            continue
    return None


# Etiquetas de encabezado (principal y secundario partido) que Flexus
# puede dejar caer en la 1ra columna al cortar tablas anchas entre páginas.
ETIQUETAS_HEADER = {
    "proveedor", "campos", "cantidad", "p. unitario", "precio compra",
    "p. total", "peso artículo (kg)", "volumen artículo (cm3)",
    "peso artículo comp.", "n.comp.", "cuit", "cod. artículo",
}


def es_fila_fecha_paginacion(fila):
    """
    Detecta la fila de paginación que Flexus mete antes de cada encabezado
    repetido. Estas filas tienen dos marcas características:
      - primera celda = una fecha dd/mm/yyyy (fecha de impresión), y
      - alguna celda con el texto 'Página X de Y'.
    Basta con cualquiera de las dos para clasificarla como paginación.
    """
    # marca fuerte: 'Página X de Y' en cualquier celda
    for c in fila:
        if c is not None and _norm(c).startswith("pagina ") and " de " in _norm(c):
            return True
    # marca de respaldo: fecha sola en la 1ra celda, resto sin datos de negocio
    c0 = fila[0]
    if c0 is not None:
        s = str(c0).strip()
        es_fecha = (len(s) == 10 and s[2] == "/" and s[5] == "/")
        # 'resto de negocio' = columnas 1..14 (proveedor..comentarios previos)
        resto = fila[1:15] if len(fila) > 1 else ()
        resto_vacio = all(c is None or str(c).strip() == "" for c in resto)
        if es_fecha and resto_vacio:
            return True
    return False


def es_fila_resumen(fila):
    """
    Detecta el cuadro de totales que Flexus agrega al pie del reporte:
    filas con etiquetas 'Campos', 'Suma', 'Promedio' fuera de la 1ra columna.
    No son datos ni errores; son un resumen estadístico a descartar.
    """
    valores = {_norm(c) for c in fila if c is not None and str(c).strip() != ""}
    marcas = {"campos", "suma", "promedio", "maximo", "minimo"}
    return len(valores & marcas) >= 2 and _norm(fila[0]) in ("", "campos")


def es_header_repetido(fila, indice=None):
    """
    Detecta un encabezado repetido embebido como fila de datos. Robusto al
    orden de columnas: en lugar de posiciones fijas, cuenta cuántas celdas de
    la fila son exactamente etiquetas de encabezado de Flexus.
    """
    valores = [_norm(c) for c in fila if c is not None and str(c).strip() != ""]
    if not valores:
        return False
    # Si varias celdas coinciden con nombres de columnas de Flexus, es un header.
    coincidencias = sum(1 for v in valores if v in MAPA_FLEXUS_A_BD)
    if coincidencias >= 3:
        return True
    # header secundario partido: única celda y es una etiqueta conocida
    if len(valores) == 1 and valores[0] in ETIQUETAS_HEADER:
        return True
    return False


def hash_linea(id_empresa, t_comp, n_comp, cod_articulo, f_comp,
               p_unitario, cantidad, p_total):
    """
    Clave de negocio para anti-duplicados. Identifica una línea de factura
    de forma estable: si recargás el mismo Excel, este hash se repite y el
    sistema puede frenar la duplicación.

    OJO: una misma factura puede tener VARIOS renglones del mismo artículo al
    mismo precio pero con distinta cantidad (ej: 6, 13, 76 unidades). Eso es
    legítimo, no un duplicado. Por eso la clave incluye cantidad y p_total,
    que distinguen renglones reales dentro de una misma factura.
    """
    base = (f"{id_empresa}|{t_comp}|{n_comp}|{cod_articulo}|{f_comp}"
            f"|{p_unitario}|{cantidad}|{p_total}")
    return hashlib.sha1(base.encode("utf-8")).hexdigest()[:16]


# ---------------------------------------------------------------------------
# VALIDACIÓN DE ESTRUCTURA
# ---------------------------------------------------------------------------

def validar_estructura(fila_header):
    """
    Verifica que el encabezado contenga las columnas esperadas, SIN importar
    el orden (Flexus las exporta en orden variable) ni el formato (CSV o Excel).
    Mapea por nombre. Recibe el encabezado ya como lista de valores.

    Devuelve (ok, mensaje, indice) donde 'indice' es un dict
    {nombre_bd: posicion_en_la_fila} para leer cada columna por su nombre.
    """
    fila_header = [limpiar_texto(c) for c in fila_header]

    # construir índice nombre_bd -> posición, usando nombres normalizados
    indice = {}
    encontradas = set()
    for pos, nombre in enumerate(fila_header):
        clave = _norm(nombre)
        if clave in MAPA_FLEXUS_A_BD:
            nombre_bd = MAPA_FLEXUS_A_BD[clave]
            indice[nombre_bd] = pos
            encontradas.add(nombre_bd)

    # las columnas opcionales (ej. ITEMFLEXXUS) no cuentan como faltantes
    obligatorias = set(COLUMNAS_BD) - COLUMNAS_OPCIONALES
    faltantes = obligatorias - encontradas
    if faltantes:
        inv = {v: k for k, v in MAPA_FLEXUS_A_BD.items()}
        visibles = sorted(inv.get(f, f) for f in faltantes)
        msg = ("El encabezado NO coincide con un export de Flexus válido.\n"
               f"   Faltan columnas: {visibles}\n")
        return False, msg, None

    return True, "Estructura válida: columnas reconocidas (orden y formato flexibles).", indice


# ---------------------------------------------------------------------------
# PROCESO PRINCIPAL
# ---------------------------------------------------------------------------

def procesar(ruta_archivo, nombre_empresa, id_empresa, usuario_carga=None):
    """
    Procesa un archivo de Flexus (CSV real o Excel del reporte interno).
    - usuario_carga: nombre del usuario que sube el archivo (queda registrado
      en cada línea junto con la fecha/hora de carga).
    """
    ruta = Path(ruta_archivo)
    if not ruta.exists():
        raise FileNotFoundError(f"No existe el archivo: {ruta}")

    # lector unificado: devuelve todas las filas + índice del encabezado
    filas, idx_header = leer_filas(ruta)
    if idx_header >= len(filas):
        return {"ok": False, "mensaje": "El archivo no tiene encabezado legible."}

    fila_header = filas[idx_header]
    filas_datos = filas[idx_header + 1:]

    # --- 1. Validar estructura ANTES de tocar nada ---
    ok, msg_val, indice = validar_estructura(fila_header)
    if not ok:
        return {"ok": False, "mensaje": msg_val}

    # fecha/hora de esta carga (una sola para todo el archivo)
    momento_carga = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # --- 2. Recorrer filas de datos y clasificar ---
    limpias, gastos, rechazadas = [], [], []
    cont = {"header_rep": 0, "fecha_pag": 0, "vacias": 0,
            "resumen": 0, "total_leidas": 0}

    def celda(fila, nombre_bd):
        """Lee una celda por nombre de columna usando el índice, robusto a
        filas de largo variable."""
        pos = indice.get(nombre_bd)
        if pos is None or pos >= len(fila):
            return None
        return fila[pos]

    for fila in filas_datos:
        cont["total_leidas"] += 1

        # fila totalmente vacía
        if all(c is None or str(c).strip() == "" for c in fila):
            cont["vacias"] += 1
            continue
        # fila de fecha de paginación (chequear ANTES del header)
        if es_fila_fecha_paginacion(fila):
            cont["fecha_pag"] += 1
            continue
        # cuadro de totales/resumen del pie
        if es_fila_resumen(fila):
            cont["resumen"] += 1
            continue
        # header repetido embebido (principal o secundario partido)
        if es_header_repetido(fila, indice):
            cont["header_rep"] += 1
            continue

        # --- fila de datos: leer por nombre y normalizar ---
        reg = {c: celda(fila, c) for c in COLUMNAS_BD}

        for c in COLUMNAS_BD:
            if c in COLS_NUMERICAS:
                reg[c] = a_numero(reg[c])
            elif c == "f_comp":
                reg[c] = a_fecha_iso(reg[c])
            else:
                reg[c] = limpiar_texto(reg[c])

        # control mínimo de calidad: tiene que tener proveedor
        if reg["proveedor"] is None:
            rechazadas.append({**reg, "_motivo": "Sin proveedor"})
            continue

        # metadatos de control
        reg["id_empresa"] = id_empresa
        reg["empresa"] = nombre_empresa
        # trazabilidad de carga: quién y cuándo subió este registro
        reg["usuario_carga"] = usuario_carga
        reg["fecha_carga"] = momento_carga
        reg["hash_linea"] = hash_linea(
            id_empresa, reg["t_comp"], reg["n_comp"],
            reg["cod_articulo"], reg["f_comp"], reg["p_unitario"],
            reg["cantidad"], reg["p_total"],
        )

        # --- rutear: gasto vario vs artículo real ---
        if reg["cod_articulo"] == CODIGO_GASTO:
            gastos.append(reg)
        else:
            limpias.append(reg)

    return {
        "ok": True,
        "mensaje": msg_val,
        "limpias": limpias,
        "gastos": gastos,
        "rechazadas": rechazadas,
        "contadores": cont,
    }


# ---------------------------------------------------------------------------
# SALIDA: archivos + reporte
# ---------------------------------------------------------------------------

def escribir_salida(res, ruta_xlsx, nombre_empresa, id_empresa):
    ruta = Path(ruta_xlsx)
    base = ruta.with_suffix("")   # sin extensión

    df_limpio = pd.DataFrame(res["limpias"])
    df_gastos = pd.DataFrame(res["gastos"])
    df_rech = pd.DataFrame(res["rechazadas"])

    # -----------------------------------------------------------------
    # MARCA DE CONTROL: columna 'motivo_alerta'.
    # No elimina NADA. Solo señala filas que un humano debería revisar
    # en la etapa de aprobación. Una fila puede acumular varios motivos.
    # -----------------------------------------------------------------
    if not df_limpio.empty:
        alertas = [[] for _ in range(len(df_limpio))]

        # 1) posible duplicado: misma clave de negocio (hash) que otra fila.
        #    Se marca a TODAS las filas del grupo (keep=False), no solo una.
        dup_mask = df_limpio["hash_linea"].duplicated(keep=False)
        for i in df_limpio.index[dup_mask]:
            alertas[i].append("POSIBLE_DUPLICADO")

        # 2) precio unitario ausente
        for i in df_limpio.index[df_limpio["p_unitario"].isna()]:
            alertas[i].append("PRECIO_NULO")

        # 3) precio unitario en cero (compra sin valor: revisar)
        for i in df_limpio.index[df_limpio["p_unitario"] == 0]:
            alertas[i].append("PRECIO_CERO")

        # 4) cantidad ausente o cero
        cant_mask = df_limpio["cantidad"].isna() | (df_limpio["cantidad"] == 0)
        for i in df_limpio.index[cant_mask]:
            alertas[i].append("CANTIDAD_INVALIDA")

        df_limpio["motivo_alerta"] = ["; ".join(a) if a else "OK" for a in alertas]

    salidas = {}
    if not df_limpio.empty:
        p = f"{base}_LIMPIO.xlsx"; df_limpio.to_excel(p, index=False); salidas["limpio"] = p
    if not df_gastos.empty:
        p = f"{base}_GASTOS.xlsx"; df_gastos.to_excel(p, index=False); salidas["gastos"] = p
    if not df_rech.empty:
        p = f"{base}_RECHAZADAS.xlsx"; df_rech.to_excel(p, index=False); salidas["rechazadas"] = p

    # detectar duplicados internos (por hash) dentro del propio archivo
    dups = 0
    desglose_alertas = {}
    if not df_limpio.empty:
        dups = int(df_limpio["hash_linea"].duplicated().sum())
        # contar cada tipo de alerta
        for motivos in df_limpio["motivo_alerta"]:
            if motivos and motivos != "OK":
                for m in motivos.split("; "):
                    desglose_alertas[m] = desglose_alertas.get(m, 0) + 1
    total_marcadas = int((df_limpio["motivo_alerta"] != "OK").sum()) if not df_limpio.empty else 0

    c = res["contadores"]
    reporte = f"""
============================================================
 REPORTE DE LIMPIEZA - FLEXUS
============================================================
 Empresa           : {nombre_empresa} (id_empresa={id_empresa})
 Archivo           : {ruta.name}
 Fecha proceso     : {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}
 Validación        : {res['mensaje'].strip()}
------------------------------------------------------------
 Filas leídas (crudas)        : {c['total_leidas']:>8}
   - Headers repetidos        : {c['header_rep']:>8}  (eliminadas)
   - Filas de fecha/paginación: {c['fecha_pag']:>8}  (eliminadas)
   - Cuadro de totales/resumen: {c['resumen']:>8}  (eliminadas)
   - Filas vacías             : {c['vacias']:>8}  (eliminadas)
------------------------------------------------------------
 Artículos reales -> LIMPIO   : {len(res['limpias']):>8}
 Gastos varios (*) -> GASTOS  : {len(res['gastos']):>8}
 Rechazadas                   : {len(res['rechazadas']):>8}
------------------------------------------------------------
 Filas marcadas para revisar     : {total_marcadas:>7}  (columna 'motivo_alerta')
{chr(10).join(f"     - {k:<20}: {v:>6}" for k, v in desglose_alertas.items()) if desglose_alertas else "     (ninguna)"}
============================================================
 Nota: NINGUNA fila fue eliminada por estar marcada. Todas están en el
 archivo LIMPIO. La columna 'motivo_alerta' señala qué revisar en la
 etapa de aprobación humana (pipeline de 2 etapas). POSIBLE_DUPLICADO
 suele ser artículos distintos de la misma factura con igual precio y
 cantidad (ej. dos sabores), no una fila repetida.
============================================================
"""
    print(reporte)
    with open(f"{base}_REPORTE.txt", "w", encoding="utf-8") as fh:
        fh.write(reporte)

    print("Archivos generados:")
    for k, v in salidas.items():
        print(f"   [{k}] {v}")
    print(f"   [reporte] {base}_REPORTE.txt")
    return salidas


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def main():
    ap = argparse.ArgumentParser(description="Limpiador de exports de Flexus (Fase 1)")
    ap.add_argument("archivo", help="Ruta al .xlsx exportado de Flexus")
    ap.add_argument("--empresa", required=True, help='Nombre de la empresa, ej: "Fly Kitchen S.A."')
    ap.add_argument("--id-empresa", type=int, required=True, help="ID numérico de la empresa")
    args = ap.parse_args()

    print(f"\n> Procesando: {args.archivo}")
    print(f"> Empresa   : {args.empresa} (id={args.id_empresa})\n")

    res = procesar(args.archivo, args.empresa, args.id_empresa)

    if not res["ok"]:
        print("ARCHIVO RECHAZADO")
        print(res["mensaje"])
        sys.exit(1)

    escribir_salida(res, args.archivo, args.empresa, args.id_empresa)


if __name__ == "__main__":
    main()
