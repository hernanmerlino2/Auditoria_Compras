"""
============================================================================
 ESQUEMA DE BASE DE DATOS  ·  Fase 2
 Sistema de Auditoría de Compras - Fly Kitchen / Drill / etc.
============================================================================

Define todas las tablas con SQLAlchemy. Motor: SQLite (un archivo .db por PC).
Al usar SQLAlchemy, migrar a PostgreSQL en el futuro NO requiere reescribir
esta lógica: solo cambia la cadena de conexión.

Tablas:
  - empresa           : catálogo de empresas (Fly Kitchen, Drill, Mítica...)
  - rol               : Administrador / Auditor / Consulta
  - usuario           : login del sistema, con rol y empresa opcional
  - proveedor         : proveedores por empresa
  - articulo          : artículos reales (sin los '*')
  - carga_archivo     : trazabilidad de cada import (auditoría)
  - compra_linea      : HECHO central. Una fila por línea de factura.
  - gasto_vario       : circuito separado para las líneas con código '*'
  - precio_historico  : evolución de precio por artículo/proveedor/fecha

Relación clave: casi todo cuelga de 'empresa' vía id_empresa, para poder
filtrar y aislar los datos de cada compañía.
"""

from datetime import datetime

from sqlalchemy import (
    create_engine, Column, Integer, String, Float, Date, DateTime,
    ForeignKey, UniqueConstraint, Boolean, Index,
)
from sqlalchemy.orm import declarative_base, relationship, sessionmaker

Base = declarative_base()


# ---------------------------------------------------------------------------
# SEGURIDAD Y CATÁLOGO
# ---------------------------------------------------------------------------

class Empresa(Base):
    __tablename__ = "empresa"
    id_empresa = Column(Integer, primary_key=True, autoincrement=True)
    nombre = Column(String(200), nullable=False, unique=True)
    cuit = Column(String(20))
    activo = Column(Boolean, default=True)

    proveedores = relationship("Proveedor", back_populates="empresa")
    articulos = relationship("Articulo", back_populates="empresa")


class Rol(Base):
    __tablename__ = "rol"
    id_rol = Column(Integer, primary_key=True, autoincrement=True)
    nombre = Column(String(50), nullable=False, unique=True)  # Administrador/Auditor/Consulta

    usuarios = relationship("Usuario", back_populates="rol")


class Usuario(Base):
    __tablename__ = "usuario"
    id_usuario = Column(Integer, primary_key=True, autoincrement=True)
    # id_empresa nullable = usuario con acceso a TODAS las empresas
    id_empresa = Column(Integer, ForeignKey("empresa.id_empresa"), nullable=True)
    id_rol = Column(Integer, ForeignKey("rol.id_rol"), nullable=False)
    nombre = Column(String(120), nullable=False)
    email = Column(String(200), nullable=False, unique=True)
    password_hash = Column(String(255), nullable=False)  # nunca texto plano
    activo = Column(Boolean, default=True)
    creado = Column(DateTime, default=datetime.utcnow)

    rol = relationship("Rol", back_populates="usuarios")


# ---------------------------------------------------------------------------
# DIMENSIONES
# ---------------------------------------------------------------------------

class Proveedor(Base):
    __tablename__ = "proveedor"
    id_proveedor = Column(Integer, primary_key=True, autoincrement=True)
    id_empresa = Column(Integer, ForeignKey("empresa.id_empresa"), nullable=False)
    cuit = Column(String(20))
    razon_social = Column(String(250), nullable=False)

    empresa = relationship("Empresa", back_populates="proveedores")
    # un mismo proveedor (por CUIT) no se repite dentro de una empresa
    __table_args__ = (
        UniqueConstraint("id_empresa", "cuit", name="uq_prov_empresa_cuit"),
    )


class Articulo(Base):
    __tablename__ = "articulo"
    id_articulo = Column(Integer, primary_key=True, autoincrement=True)
    id_empresa = Column(Integer, ForeignKey("empresa.id_empresa"), nullable=False)
    cod_articulo = Column(String(50), nullable=False)
    descripcion = Column(String(300))
    cod_rubro = Column(String(50))
    rubro = Column(String(150))
    marca = Column(String(150))
    u_medida = Column(String(50))

    empresa = relationship("Empresa", back_populates="articulos")
    # el código de artículo es único dentro de cada empresa
    __table_args__ = (
        UniqueConstraint("id_empresa", "cod_articulo", name="uq_art_empresa_cod"),
    )


# ---------------------------------------------------------------------------
# TRAZABILIDAD DE CARGAS
# ---------------------------------------------------------------------------

class CargaArchivo(Base):
    __tablename__ = "carga_archivo"
    id_carga = Column(Integer, primary_key=True, autoincrement=True)
    id_empresa = Column(Integer, ForeignKey("empresa.id_empresa"), nullable=False)
    id_usuario = Column(Integer, ForeignKey("usuario.id_usuario"), nullable=True)
    nombre_archivo = Column(String(300), nullable=False)
    hash_archivo = Column(String(64))  # para detectar recarga del MISMO archivo
    fecha_carga = Column(DateTime, default=datetime.utcnow)
    filas_ok = Column(Integer, default=0)
    filas_gastos = Column(Integer, default=0)
    filas_duplicadas = Column(Integer, default=0)
    filas_rechazadas = Column(Integer, default=0)


# ---------------------------------------------------------------------------
# HECHO CENTRAL: LÍNEAS DE COMPRA
# ---------------------------------------------------------------------------

class CompraLinea(Base):
    __tablename__ = "compra_linea"
    # id_linea = ID técnico autoincremental (uno nuevo por cada registro)
    id_linea = Column(Integer, primary_key=True, autoincrement=True)
    id_empresa = Column(Integer, ForeignKey("empresa.id_empresa"), nullable=False)
    id_proveedor = Column(Integer, ForeignKey("proveedor.id_proveedor"))
    id_articulo = Column(Integer, ForeignKey("articulo.id_articulo"))
    id_carga = Column(Integer, ForeignKey("carga_archivo.id_carga"))

    # datos de la línea
    cantidad = Column(Float)
    p_unitario = Column(Float)
    precio_compra = Column(Float)
    p_total = Column(Float)
    u_medida = Column(String(50))
    peso_articulo_kg = Column(Float)
    volumen_articulo_cm3 = Column(Float)
    peso_articulo_comp = Column(Float)
    comentarios = Column(String(500))
    cod_deposito = Column(String(50))
    desc_deposito = Column(String(150))

    # comprobante
    t_comp = Column(String(20))
    n_comp = Column(String(50))
    f_comp = Column(Date)

    # control
    hash_linea = Column(String(32), nullable=False)  # clave de negocio anti-duplicados
    motivo_alerta = Column(String(200))

    # el hash de negocio es único por empresa: frena recargar la misma línea
    __table_args__ = (
        UniqueConstraint("id_empresa", "hash_linea", name="uq_compra_empresa_hash"),
        Index("ix_compra_empresa_fecha", "id_empresa", "f_comp"),
        Index("ix_compra_articulo", "id_articulo"),
    )


class GastoVario(Base):
    """Circuito separado para las líneas con código '*' (no comparables)."""
    __tablename__ = "gasto_vario"
    id_gasto = Column(Integer, primary_key=True, autoincrement=True)
    id_empresa = Column(Integer, ForeignKey("empresa.id_empresa"), nullable=False)
    id_proveedor = Column(Integer, ForeignKey("proveedor.id_proveedor"))
    id_carga = Column(Integer, ForeignKey("carga_archivo.id_carga"))

    descripcion = Column(String(300))  # el 'Artículo' del *; ej. 'FONDO FIJO...'
    cantidad = Column(Float)
    p_unitario = Column(Float)
    precio_compra = Column(Float)
    p_total = Column(Float)
    u_medida = Column(String(50))
    comentarios = Column(String(500))
    t_comp = Column(String(20))
    n_comp = Column(String(50))
    f_comp = Column(Date)
    hash_linea = Column(String(32), nullable=False)

    __table_args__ = (
        UniqueConstraint("id_empresa", "hash_linea", name="uq_gasto_empresa_hash"),
    )


# ---------------------------------------------------------------------------
# HISTÓRICO DE PRECIOS
# ---------------------------------------------------------------------------

class PrecioHistorico(Base):
    __tablename__ = "precio_historico"
    id = Column(Integer, primary_key=True, autoincrement=True)
    id_empresa = Column(Integer, ForeignKey("empresa.id_empresa"), nullable=False)
    id_articulo = Column(Integer, ForeignKey("articulo.id_articulo"), nullable=False)
    id_proveedor = Column(Integer, ForeignKey("proveedor.id_proveedor"))
    precio = Column(Float, nullable=False)     # p_unitario pagado
    f_comp = Column(Date, nullable=False)

    __table_args__ = (
        Index("ix_precio_art_fecha", "id_articulo", "f_comp"),
    )


# ---------------------------------------------------------------------------
# CONEXIÓN
# ---------------------------------------------------------------------------

def crear_engine(ruta_db="auditoria_compras.db", echo=False):
    """Crea el engine SQLite. Cambiar aquí la URL para migrar a PostgreSQL."""
    return create_engine(f"sqlite:///{ruta_db}", echo=echo)


def crear_esquema(engine):
    """Crea todas las tablas si no existen."""
    Base.metadata.create_all(engine)


def get_session(engine):
    """Devuelve una sesión para operar sobre la BD."""
    Session = sessionmaker(bind=engine)
    return Session()


# semillas mínimas: roles estándar
ROLES_ESTANDAR = ["Administrador", "Auditor", "Consulta"]


def sembrar_roles(session):
    """Inserta los roles estándar si aún no existen."""
    existentes = {r.nombre for r in session.query(Rol).all()}
    for nombre in ROLES_ESTANDAR:
        if nombre not in existentes:
            session.add(Rol(nombre=nombre))
    session.commit()


if __name__ == "__main__":
    # crear la base vacía con el esquema y los roles semilla
    eng = crear_engine()
    crear_esquema(eng)
    ses = get_session(eng)
    sembrar_roles(ses)
    print("Base creada con esquema y roles estándar:", ROLES_ESTANDAR)
    ses.close()
