"""
============================================================================
 INTERFAZ GRÁFICA (Streamlit)  ·  Fase 3
 Sistema de Auditoría de Compras
============================================================================

Interfaz web local para:
  1. Iniciar sesión (usuario / rol).
  2. Cargar un Excel de Flexus: elegir empresa -> limpiar -> guardar en la BD,
     todo desde un botón (integra Fase 1 + Fase 2).
  3. Ver reportes de auditoría: tablero, top de artículos, evolución de precio
     por producto (con ajuste por dólar e inflación), variaciones y alertas.
  4. Cargar las cotizaciones mensuales (dólar / IPC).

Cómo se ejecuta (desde la carpeta del proyecto, en la terminal):
    streamlit run app.py

Eso abre el navegador en http://localhost:8501
"""

import io
import hashlib
from datetime import date

import pandas as pd
import streamlit as st

import metricas as m
from esquema import crear_engine, crear_esquema, get_session, sembrar_roles, Usuario, Rol
# Fase 1 y 2
import limpiador_flexus as limpiador
import cargador as cargador_mod

RUTA_DB = "auditoria_compras.db"

st.set_page_config(page_title="Auditoría de Compras", page_icon="📊", layout="wide")


# ---------------------------------------------------------------------------
# SEGURIDAD MÍNIMA (hash de contraseñas)
# ---------------------------------------------------------------------------

def hash_pass(p):
    return hashlib.sha256(p.encode("utf-8")).hexdigest()


def crear_usuario_admin_si_no_hay(engine):
    """Si no hay ningún usuario, crea uno admin por defecto (admin / admin123)."""
    ses = get_session(engine)
    sembrar_roles(ses)
    if ses.query(Usuario).count() == 0:
        rol_admin = ses.query(Rol).filter_by(nombre="Administrador").first()
        ses.add(Usuario(
            nombre="Administrador", email="admin",
            password_hash=hash_pass("admin123"),
            id_rol=rol_admin.id_rol, activo=True,
        ))
        ses.commit()
    ses.close()


def validar_login(engine, email, password):
    ses = get_session(engine)
    u = ses.query(Usuario).filter_by(email=email, activo=True).first()
    ok = u is not None and u.password_hash == hash_pass(password)
    datos = None
    if ok:
        rol = ses.query(Rol).filter_by(id_rol=u.id_rol).first()
        datos = {"nombre": u.nombre, "rol": rol.nombre, "id_empresa": u.id_empresa}
    ses.close()
    return ok, datos


# ---------------------------------------------------------------------------
# PANTALLA DE LOGIN
# ---------------------------------------------------------------------------

def pantalla_login(engine):
    st.title("📊 Auditoría de Compras")
    st.caption("Ingresá para continuar")
    with st.container(border=True):
        email = st.text_input("Usuario", value="", placeholder="admin")
        password = st.text_input("Contraseña", type="password", placeholder="admin123")
        if st.button("Ingresar", type="primary"):
            ok, datos = validar_login(engine, email, password)
            if ok:
                st.session_state["auth"] = datos
                st.rerun()
            else:
                st.error("Usuario o contraseña incorrectos.")
    st.info("Primera vez: usuario **admin** · contraseña **admin123** (cambiala después).")


# ---------------------------------------------------------------------------
# SECCIÓN: CARGA DE ARCHIVO
# ---------------------------------------------------------------------------

def seccion_carga(engine):
    st.header("📥 Cargar archivo de Flexus")
    st.write("Elegí la empresa, subí el Excel exportado de Flexus y cargalo a la base. "
             "El sistema limpia, valida y guarda automáticamente.")

    empresas = m.listar_empresas(engine)
    opciones = dict(zip(empresas["nombre"], empresas["id_empresa"]))
    col1, col2 = st.columns([1, 1])
    with col1:
        nombre_emp = st.selectbox("Empresa", list(opciones.keys()) + ["+ Nueva empresa"])
    with col2:
        if nombre_emp == "+ Nueva empresa":
            nombre_emp = st.text_input("Nombre de la nueva empresa")

    archivo = st.file_uploader(
        "Archivo de Flexus (.csv o .xlsx)", type=["csv", "xlsx"])

    # usuario que está cargando (queda registrado en cada línea)
    usuario_actual = st.session_state.get("auth", {}).get("nombre", "desconocido")

    if archivo and nombre_emp and nombre_emp != "+ Nueva empresa":
        st.caption(f"Se registrará esta carga a nombre de **{usuario_actual}** "
                   f"con la fecha y hora actual.")
        if st.button("🚀 Procesar y cargar", type="primary"):
            with st.spinner("Limpiando y cargando..."):
                # guardar el archivo subido a disco temporal
                ruta_tmp = f"_tmp_{archivo.name}"
                with open(ruta_tmp, "wb") as f:
                    f.write(archivo.getbuffer())

                # id de empresa: si es nueva, se creará en la carga
                id_emp = opciones.get(nombre_emp, 0)

                # FASE 1: limpiar (registra usuario y fecha de carga)
                res = limpiador.procesar(ruta_tmp, nombre_emp, id_emp or 0,
                                         usuario_carga=usuario_actual)
                if not res["ok"]:
                    st.error("Archivo rechazado en la validación:")
                    st.code(res["mensaje"])
                    return
                salidas = limpiador.escribir_salida(res, ruta_tmp, nombre_emp, id_emp or 0)

                # FASE 2: cargar a BD
                ruta_limpio = salidas.get("limpio")
                ruta_gastos = salidas.get("gastos")
                if not ruta_limpio:
                    st.warning("No se generaron filas limpias para cargar.")
                    return
                r = cargador_mod.cargar(ruta_limpio, ruta_gastos, nombre_emp, RUTA_DB,
                                        usuario_carga=usuario_actual)

            st.success("¡Carga completada!")
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Compras insertadas", f"{r['ok']:,}")
            c2.metric("Duplicadas omitidas", f"{r['dup']:,}")
            c3.metric("Gastos insertados", f"{r['gastos_ok']:,}")
            c4.metric("Gastos duplicados", f"{r['gastos_dup']:,}")
            st.caption("Las duplicadas se omiten para no contaminar el histórico. "
                       "Recargar el mismo archivo no genera datos repetidos.")


# ---------------------------------------------------------------------------
# SECCIÓN: TABLERO
# ---------------------------------------------------------------------------

def seccion_tablero(engine, id_empresa):
    st.header("📈 Tablero")
    kpis = m.kpis_empresa(engine, id_empresa)
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total comprado", f"${(kpis['total_comprado'] or 0):,.0f}")
    c2.metric("Nº de compras", f"{int(kpis['n_compras'] or 0):,}")
    c3.metric("Artículos", f"{int(kpis['n_articulos'] or 0):,}")
    c4.metric("Proveedores", f"{int(kpis['n_proveedores'] or 0):,}")

    st.subheader("Evolución mensual de compras")
    serie = m.compras_por_mes(engine, id_empresa)
    if not serie.empty:
        st.bar_chart(serie.set_index("mes")["total"])
    else:
        st.info("No hay datos de compras para esta empresa.")


# ---------------------------------------------------------------------------
# SECCIÓN: TOP DE ARTÍCULOS
# ---------------------------------------------------------------------------

def seccion_top(engine, id_empresa):
    st.header("🏆 Artículos más relevantes")
    col1, col2 = st.columns([1, 1])
    with col1:
        por = st.radio("Ordenar por", ["monto", "cantidad"], horizontal=True)
    with col2:
        limite = st.slider("Cantidad a mostrar", 5, 50, 15)

    df = m.top_articulos(engine, id_empresa, por=por, limite=limite)
    if df.empty:
        st.info("Sin datos.")
        return
    etiqueta = "Monto comprado ($)" if por == "monto" else "Cantidad total"
    df_show = df.rename(columns={"valor": etiqueta, "descripcion": "Artículo",
                                 "n_compras": "Nº compras", "precio_prom": "Precio prom."})
    st.dataframe(df_show, use_container_width=True, hide_index=True)
    st.bar_chart(df.set_index("descripcion")["valor"])


# ---------------------------------------------------------------------------
# SECCIÓN: EVOLUCIÓN DE PRECIO POR ARTÍCULO
# ---------------------------------------------------------------------------

def seccion_evolucion(engine, id_empresa):
    st.header("💰 Evolución de precio por artículo")
    st.caption("Precio unitario promedio ponderado por cantidad. "
               "Si cargaste cotizaciones, se puede ver ajustado por dólar e inflación.")

    buscar = st.text_input("Buscar artículo (nombre o código)")
    arts = m.listar_articulos(engine, id_empresa, buscar=buscar, limite=100)
    if arts.empty:
        st.info("No se encontraron artículos.")
        return

    arts["etiqueta"] = arts["descripcion"].fillna("(sin desc)") + " · " + \
        arts["cod_articulo"].astype(str) + "  (" + arts["n_compras"].astype(str) + " compras)"
    sel = st.selectbox("Artículo", arts["etiqueta"])
    id_art = int(arts.loc[arts["etiqueta"] == sel, "id_articulo"].iloc[0])

    df = m.evolucion_precio(engine, id_art)
    if df.empty:
        st.info("Este artículo no tiene suficientes datos de precio.")
        return

    vista = st.radio("Ver precio en",
                     ["Nominal ($)", "Dólares (USD)", "Real (ajustado por inflación)"],
                     horizontal=True)
    col_map = {"Nominal ($)": "precio_nominal", "Dólares (USD)": "precio_usd",
               "Real (ajustado por inflación)": "precio_real"}
    col = col_map[vista]

    if df[col].notna().any():
        st.line_chart(df.set_index("mes")[col])
    else:
        st.warning("No hay cotizaciones cargadas para esta vista. "
                   "Cargá dólar/IPC en la sección 'Cotizaciones'.")

    # variación total del período
    serie = df[col].dropna()
    if len(serie) >= 2 and serie.iloc[0]:
        var = (serie.iloc[-1] - serie.iloc[0]) / serie.iloc[0] * 100
        st.metric(f"Variación del período ({vista})", f"{var:+.1f}%")

    st.dataframe(df, use_container_width=True, hide_index=True)

    # comparativa de proveedores
    st.subheader("Comparativa de proveedores para este artículo")
    comp = m.comparar_proveedores(engine, id_art)
    if not comp.empty:
        st.dataframe(comp, use_container_width=True, hide_index=True)


# ---------------------------------------------------------------------------
# SECCIÓN: VARIACIONES / AUDITORÍA
# ---------------------------------------------------------------------------

def seccion_variaciones(engine, id_empresa):
    st.header("🔎 Auditoría de variaciones de precio")
    st.caption("Artículos cuyo precio cambió más entre su primer y último mes con compras.")
    col1, col2 = st.columns(2)
    with col1:
        umbral = st.slider("Variación mínima (%)", 10, 200, 50, step=10)
    with col2:
        min_c = st.slider("Mínimo de compras", 2, 20, 5)

    df = m.variacion_articulos(engine, id_empresa, umbral_pct=umbral, min_compras=min_c)
    if df.empty:
        st.info("No hay artículos que superen ese umbral.")
        return
    st.write(f"**{len(df)} artículos** con variación ≥ {umbral}%.")
    st.caption("⚠️ Variaciones extremas (miles de %) suelen indicar errores de carga "
               "o de unidad de medida en Flexus, no aumentos reales. Revisá esos casos.")
    st.dataframe(df, use_container_width=True, hide_index=True)


# ---------------------------------------------------------------------------
# SECCIÓN: ALERTAS
# ---------------------------------------------------------------------------

def seccion_alertas(engine, id_empresa):
    st.header("⚠️ Filas marcadas para revisión")
    df = m.filas_con_alerta(engine, id_empresa)
    if df.empty:
        st.success("No hay filas marcadas. Todo OK.")
        return
    st.write(f"**{len(df)} líneas** marcadas por el limpiador (no se eliminaron).")
    st.dataframe(df, use_container_width=True, hide_index=True)


# ---------------------------------------------------------------------------
# SECCIÓN: COTIZACIONES
# ---------------------------------------------------------------------------

def seccion_cotizaciones(engine):
    st.header("💵 Cotizaciones mensuales (dólar / IPC)")
    st.caption("Cargá una vez por mes el dólar de referencia y el índice IPC del INDEC. "
               "El sistema los usa para ajustar los precios por dólar e inflación.")

    with st.container(border=True):
        c1, c2, c3, c4 = st.columns([1, 1, 1, 1])
        with c1:
            mes = st.text_input("Mes (YYYY-MM)", value=str(date.today())[:7])
        with c2:
            dolar = st.number_input("Dólar", min_value=0.0, step=1.0)
        with c3:
            ipc = st.number_input("IPC (índice)", min_value=0.0, step=0.1)
        with c4:
            st.write("")
            st.write("")
            if st.button("Guardar", type="primary"):
                m.guardar_cotizacion(engine, mes, dolar or None, ipc or None)
                st.success(f"Cotización de {mes} guardada.")

    st.subheader("Cotizaciones cargadas")
    cot = m.leer_cotizaciones(engine)
    if cot.empty:
        st.info("Todavía no cargaste cotizaciones.")
    else:
        st.dataframe(cot, use_container_width=True, hide_index=True)


# ---------------------------------------------------------------------------
# SECCIÓN: GESTIÓN DE USUARIOS (solo admin)
# ---------------------------------------------------------------------------

def seccion_usuarios(engine):
    st.header("👥 Gestión de usuarios")
    st.caption("Creá los usuarios para los administrativos. El rol 'Consulta' "
               "les permite ver reportes y cargar cotizaciones, pero NO cargar "
               "archivos de Flexus.")

    ses = get_session(engine)
    roles = ses.query(Rol).all()
    roles_dict = {r.nombre: r.id_rol for r in roles}

    with st.container(border=True):
        st.subheader("Crear nuevo usuario")
        c1, c2 = st.columns(2)
        with c1:
            nombre = st.text_input("Nombre y apellido")
            email = st.text_input("Usuario (para ingresar)")
        with c2:
            rol_sel = st.selectbox("Rol", list(roles_dict.keys()),
                                   index=list(roles_dict.keys()).index("Consulta")
                                   if "Consulta" in roles_dict else 0)
            password = st.text_input("Contraseña inicial", type="password")

        if st.button("Crear usuario", type="primary"):
            if not (nombre and email and password):
                st.error("Completá nombre, usuario y contraseña.")
            elif ses.query(Usuario).filter_by(email=email).first():
                st.error(f"Ya existe un usuario con el nombre '{email}'.")
            else:
                ses.add(Usuario(
                    nombre=nombre, email=email,
                    password_hash=hash_pass(password),
                    id_rol=roles_dict[rol_sel], activo=True,
                ))
                ses.commit()
                st.success(f"Usuario '{email}' creado con rol {rol_sel}.")

    st.subheader("Usuarios existentes")
    usuarios = ses.query(Usuario).all()
    if usuarios:
        filas = []
        for u in usuarios:
            rol_nombre = next((r.nombre for r in roles if r.id_rol == u.id_rol), "?")
            filas.append({"Usuario": u.email, "Nombre": u.nombre,
                          "Rol": rol_nombre, "Activo": "Sí" if u.activo else "No"})
        st.dataframe(pd.DataFrame(filas), use_container_width=True, hide_index=True)
    ses.close()


# ---------------------------------------------------------------------------
# APP PRINCIPAL
# ---------------------------------------------------------------------------

def preparar_base_demo():
    """
    Para el demo en la nube: si la base no tiene datos, los carga desde los
    CSV comprimidos de la carpeta demo_data/. Se usa CSV (texto) en vez de un
    .db binario porque los binarios se corrompen al subirse a GitHub por la web.
    Solo corre una vez: si ya hay compras cargadas, no hace nada.
    """
    import os
    import pandas as pd
    from esquema import crear_engine, crear_esquema

    engine = crear_engine(RUTA_DB)
    crear_esquema(engine)

    # ¿ya hay datos? entonces no recargar
    try:
        n = pd.read_sql("SELECT COUNT(*) AS n FROM compra_linea", engine)["n"].iloc[0]
        if n and n > 0:
            return
    except Exception:
        pass

    carpeta = "demo_data"
    if not os.path.isdir(carpeta):
        return  # sin datos de demo; la app arranca vacía (modo real)

    # orden de carga respetando dependencias (empresa antes que sus hijos)
    tablas = ["empresa", "proveedor", "articulo", "carga_archivo",
              "compra_linea", "gasto_vario", "precio_historico"]
    for t in tablas:
        ruta = os.path.join(carpeta, f"{t}.csv.gz")
        if os.path.exists(ruta):
            df = pd.read_csv(ruta, compression="gzip")
            if not df.empty:
                df.to_sql(t, engine, if_exists="append", index=False)


def main():
    preparar_base_demo()
    engine = crear_engine(RUTA_DB)
    crear_esquema(engine)
    crear_usuario_admin_si_no_hay(engine)

    if "auth" not in st.session_state:
        pantalla_login(engine)
        return

    datos = st.session_state["auth"]
    es_admin = (datos["rol"] == "Administrador")

    with st.sidebar:
        st.markdown(f"### 👤 {datos['nombre']}")
        st.caption(f"Rol: {datos['rol']}")
        if es_admin:
            st.success("Modo administrador: podés cargar archivos.")
        else:
            st.info("Modo consulta: podés ver reportes y cargar cotizaciones.")
        st.divider()

        empresas = m.listar_empresas(engine)
        if empresas.empty:
            st.warning("No hay empresas todavía." +
                       (" Cargá un archivo primero." if es_admin
                        else " Pedile al administrador que cargue datos."))
            id_empresa = None
        else:
            nombre_sel = st.selectbox("Empresa a analizar", empresas["nombre"])
            id_empresa = int(empresas.loc[empresas["nombre"] == nombre_sel, "id_empresa"].iloc[0])

        st.divider()

        # El menú depende del rol:
        #  - Administrador: ve todo, incluida la carga de archivos.
        #  - Otros roles: ven reportes y pueden cargar cotizaciones (dólar/IPC),
        #    pero NO cargar archivos analizables.
        secciones_comunes = [
            "📈 Tablero", "🏆 Top artículos", "💰 Evolución de precio",
            "🔎 Variaciones", "⚠️ Alertas", "💵 Cotizaciones",
        ]
        if es_admin:
            opciones_menu = ["📥 Cargar archivo"] + secciones_comunes + ["👥 Usuarios"]
        else:
            opciones_menu = secciones_comunes

        seccion = st.radio("Sección", opciones_menu)
        st.divider()
        if st.button("Cerrar sesión"):
            del st.session_state["auth"]
            st.rerun()

    # ruteo de secciones
    if seccion == "📥 Cargar archivo":
        # doble barrera de seguridad: aunque alguien fuerce la sección,
        # solo el admin puede ejecutar la carga.
        if not es_admin:
            st.error("No tenés permisos para cargar archivos. "
                     "Esta función es exclusiva del administrador.")
        else:
            seccion_carga(engine)
    elif id_empresa is None:
        st.info("Todavía no hay datos cargados para analizar.")
    elif seccion == "📈 Tablero":
        seccion_tablero(engine, id_empresa)
    elif seccion == "🏆 Top artículos":
        seccion_top(engine, id_empresa)
    elif seccion == "💰 Evolución de precio":
        seccion_evolucion(engine, id_empresa)
    elif seccion == "🔎 Variaciones":
        seccion_variaciones(engine, id_empresa)
    elif seccion == "⚠️ Alertas":
        seccion_alertas(engine, id_empresa)
    elif seccion == "💵 Cotizaciones":
        seccion_cotizaciones(engine)
    elif seccion == "👥 Usuarios":
        if not es_admin:
            st.error("Solo el administrador puede gestionar usuarios.")
        else:
            seccion_usuarios(engine)


if __name__ == "__main__":
    main()
