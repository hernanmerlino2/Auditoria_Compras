"""
============================================================================
 CARGADOR A BASE DE DATOS  ·  Fase 2
 Sistema de Auditoría de Compras
============================================================================

Toma los archivos generados por el limpiador (Fase 1) y los inserta en la
base SQLite, resolviendo empresa / proveedor / artículo y aplicando control
anti-duplicados por hash de línea.

Flujo:
  1. Asegura que la empresa exista (o la crea).
  2. Registra la carga en 'carga_archivo' (trazabilidad).
  3. Por cada fila de LIMPIO:
       - resuelve/crea proveedor y artículo (cache en memoria),
       - inserta en compra_linea SI el hash no existe ya (anti-duplicado),
       - alimenta precio_historico.
  4. Por cada fila de GASTOS: inserta en gasto_vario (también anti-duplicado).
  5. Actualiza los contadores de la carga y reporta.

Uso:
    python cargador.py --limpio ARCHIVO_LIMPIO.xlsx --gastos ARCHIVO_GASTOS.xlsx \\
                       --empresa "Fly Kitchen S.A." --db auditoria_compras.db

Nota: el nombre y el id de empresa salen de las columnas del propio archivo
LIMPIO (las agregó el limpiador), pero se puede forzar con --empresa.
"""

import argparse
import hashlib
from datetime import datetime

import pandas as pd

from esquema import (
    crear_engine, crear_esquema, get_session, sembrar_roles,
    Empresa, Proveedor, Articulo, CargaArchivo, CompraLinea,
    GastoVario, PrecioHistorico,
)


def _fecha(valor):
    """Convierte 'yyyy-mm-dd' (o Timestamp) a date. None si no se puede."""
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return None
    if isinstance(valor, str):
        try:
            return datetime.strptime(valor[:10], "%Y-%m-%d").date()
        except ValueError:
            return None
    try:
        return pd.to_datetime(valor).date()
    except Exception:
        return None


def _txt(valor):
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return None
    return str(valor).strip()


def _num(valor):
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return None
    try:
        return float(valor)
    except (ValueError, TypeError):
        return None


def hash_archivo(ruta):
    """SHA-256 del archivo, para detectar recarga del MISMO archivo físico."""
    h = hashlib.sha256()
    with open(ruta, "rb") as fh:
        for bloque in iter(lambda: fh.read(8192), b""):
            h.update(bloque)
    return h.hexdigest()


def obtener_o_crear_empresa(session, nombre, cuit=None):
    emp = session.query(Empresa).filter_by(nombre=nombre).first()
    if emp is None:
        emp = Empresa(nombre=nombre, cuit=cuit, activo=True)
        session.add(emp)
        session.flush()  # asigna id sin cerrar la transacción
    return emp


def cargar(ruta_limpio, ruta_gastos, nombre_empresa, ruta_db,
           id_usuario=None, usuario_carga=None):
    engine = crear_engine(ruta_db)
    crear_esquema(engine)
    session = get_session(engine)
    sembrar_roles(session)

    df = pd.read_excel(ruta_limpio)

    # nombre de empresa: prioridad al parámetro, si no la columna del archivo
    if not nombre_empresa and "empresa" in df.columns:
        nombre_empresa = str(df["empresa"].iloc[0])
    # usuario de carga: prioridad al parámetro, si no la columna del archivo
    if not usuario_carga and "usuario_carga" in df.columns and not df.empty:
        usuario_carga = str(df["usuario_carga"].iloc[0])

    emp = obtener_o_crear_empresa(session, nombre_empresa)
    id_emp = emp.id_empresa

    # registrar la carga
    carga = CargaArchivo(
        id_empresa=id_emp,
        id_usuario=id_usuario,
        usuario_carga=usuario_carga,
        nombre_archivo=str(ruta_limpio).split("/")[-1],
        hash_archivo=hash_archivo(ruta_limpio),
        fecha_carga=datetime.utcnow(),
    )
    session.add(carga)
    session.flush()

    # caches para no re-consultar la BD en cada fila
    cache_prov = {(p.cuit if p.cuit else f"SC:{p.razon_social}"): p.id_proveedor
                  for p in session.query(Proveedor).filter_by(id_empresa=id_emp)}
    cache_art = {a.cod_articulo: a.id_articulo
                 for a in session.query(Articulo).filter_by(id_empresa=id_emp)}
    # hashes ya presentes en la BD para esta empresa (anti-duplicado entre cargas)
    hashes_existentes = {h[0] for h in session.query(CompraLinea.hash_linea)
                         .filter_by(id_empresa=id_emp)}

    ok = dup = 0
    hashes_en_lote = set()

    for _, fila in df.iterrows():
        h = _txt(fila.get("hash_linea"))
        # anti-duplicado: contra la BD y contra el propio lote
        if h in hashes_existentes or h in hashes_en_lote:
            dup += 1
            continue
        hashes_en_lote.add(h)

        # --- resolver proveedor (clave: cuit, o razón social si no hay cuit) ---
        cuit = _txt(fila.get("cuit"))
        razon = _txt(fila.get("proveedor"))
        clave_prov = cuit if cuit else f"SC:{razon}"
        id_prov = cache_prov.get(clave_prov)
        if id_prov is None:
            prov = Proveedor(id_empresa=id_emp, cuit=cuit, razon_social=razon)
            session.add(prov)
            session.flush()
            id_prov = prov.id_proveedor
            cache_prov[clave_prov] = id_prov

        # --- resolver artículo ---
        cod = _txt(fila.get("cod_articulo"))
        if cod is None or cod == "":
            cod = "SIN_CODIGO"   # placeholder: artículo real sin código en Flexus
        id_art = cache_art.get(cod)
        if id_art is None:
            art = Articulo(
                id_empresa=id_emp, cod_articulo=cod,
                descripcion=_txt(fila.get("articulo")),
                cod_rubro=_txt(fila.get("cod_rubro")),
                rubro=_txt(fila.get("rubro")),
                marca=_txt(fila.get("marca")),
                u_medida=_txt(fila.get("u_medida")),
            )
            session.add(art)
            session.flush()
            id_art = art.id_articulo
            cache_art[cod] = id_art

        # --- insertar línea de compra ---
        f_comp = _fecha(fila.get("f_comp"))
        p_unit = _num(fila.get("p_unitario"))
        linea = CompraLinea(
            id_empresa=id_emp, id_proveedor=id_prov, id_articulo=id_art,
            id_carga=carga.id_carga,
            cantidad=_num(fila.get("cantidad")),
            p_unitario=p_unit,
            precio_compra=_num(fila.get("precio_compra")),
            p_total=_num(fila.get("p_total")),
            u_medida=_txt(fila.get("u_medida")),
            peso_articulo_kg=_num(fila.get("peso_articulo_kg")),
            volumen_articulo_cm3=_num(fila.get("volumen_articulo_cm3")),
            peso_articulo_comp=_num(fila.get("peso_articulo_comp")),
            comentarios=_txt(fila.get("comentarios")),
            cod_deposito=_txt(fila.get("cod_deposito")),
            desc_deposito=_txt(fila.get("desc_deposito")),
            t_comp=_txt(fila.get("t_comp")),
            n_comp=_txt(fila.get("n_comp")),
            f_comp=f_comp,
            item_flexxus=_txt(fila.get("item_flexxus")),
            hash_linea=h,
            motivo_alerta=_txt(fila.get("motivo_alerta")),
            usuario_carga=_txt(fila.get("usuario_carga")),
            fecha_carga=_txt(fila.get("fecha_carga")),
        )
        session.add(linea)

        # --- alimentar histórico de precios ---
        if p_unit is not None and f_comp is not None:
            session.add(PrecioHistorico(
                id_empresa=id_emp, id_articulo=id_art, id_proveedor=id_prov,
                precio=p_unit, f_comp=f_comp,
            ))
        ok += 1

    # --- cargar GASTOS (circuito separado) ---
    gastos_ok = gastos_dup = 0
    if ruta_gastos:
        dfg = pd.read_excel(ruta_gastos)
        hashes_gasto = {h[0] for h in session.query(GastoVario.hash_linea)
                        .filter_by(id_empresa=id_emp)}
        hashes_gasto_lote = set()
        for _, fila in dfg.iterrows():
            h = _txt(fila.get("hash_linea"))
            if h in hashes_gasto or h in hashes_gasto_lote:
                gastos_dup += 1
                continue
            hashes_gasto_lote.add(h)

            cuit = _txt(fila.get("cuit"))
            razon = _txt(fila.get("proveedor"))
            clave_prov = cuit if cuit else f"SC:{razon}"
            id_prov = cache_prov.get(clave_prov)
            if id_prov is None:
                prov = Proveedor(id_empresa=id_emp, cuit=cuit, razon_social=razon)
                session.add(prov)
                session.flush()
                id_prov = prov.id_proveedor
                cache_prov[clave_prov] = id_prov

            session.add(GastoVario(
                id_empresa=id_emp, id_proveedor=id_prov, id_carga=carga.id_carga,
                descripcion=_txt(fila.get("articulo")),
                cantidad=_num(fila.get("cantidad")),
                p_unitario=_num(fila.get("p_unitario")),
                precio_compra=_num(fila.get("precio_compra")),
                p_total=_num(fila.get("p_total")),
                u_medida=_txt(fila.get("u_medida")),
                comentarios=_txt(fila.get("comentarios")),
                t_comp=_txt(fila.get("t_comp")),
                n_comp=_txt(fila.get("n_comp")),
                f_comp=_fecha(fila.get("f_comp")),
                item_flexxus=_txt(fila.get("item_flexxus")),
                hash_linea=h,
                usuario_carga=_txt(fila.get("usuario_carga")),
                fecha_carga=_txt(fila.get("fecha_carga")),
            ))
            gastos_ok += 1

    # actualizar contadores de la carga
    carga.filas_ok = ok
    carga.filas_duplicadas = dup + gastos_dup
    carga.filas_gastos = gastos_ok
    session.commit()

    reporte = f"""
============================================================
 CARGA A BASE DE DATOS COMPLETADA
============================================================
 Empresa            : {nombre_empresa} (id={id_emp})
 Base de datos      : {ruta_db}
 Carga #            : {carga.id_carga}
------------------------------------------------------------
 Compras insertadas         : {ok:>8}
 Compras duplicadas (omitidas): {dup:>7}
 Gastos insertados          : {gastos_ok:>8}
 Gastos duplicados (omitidos) : {gastos_dup:>7}
------------------------------------------------------------
 Proveedores en BD (empresa): {len(cache_prov):>8}
 Artículos en BD (empresa)  : {len(cache_art):>8}
============================================================
"""
    print(reporte)
    session.close()
    return {"ok": ok, "dup": dup, "gastos_ok": gastos_ok,
            "gastos_dup": gastos_dup, "id_carga": carga.id_carga}


def main():
    ap = argparse.ArgumentParser(description="Cargador a BD (Fase 2)")
    ap.add_argument("--limpio", required=True, help="Ruta al archivo _LIMPIO.xlsx")
    ap.add_argument("--gastos", help="Ruta al archivo _GASTOS.xlsx (opcional)")
    ap.add_argument("--empresa", help="Nombre de empresa (si no, se toma del archivo)")
    ap.add_argument("--db", default="auditoria_compras.db", help="Ruta al archivo .db")
    args = ap.parse_args()
    cargar(args.limpio, args.gastos, args.empresa, args.db)


if __name__ == "__main__":
    main()
