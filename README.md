# Sistema de Auditoría de Compras

App web para limpiar exports de Flexus, cargarlos a una base de datos y
generar reportes de auditoría de compras (evolución de precios, ajuste por
dólar e inflación, top de artículos, variaciones y alertas).

## Archivos del proyecto

| Archivo | Qué hace |
|---|---|
| `app.py` | La interfaz web (Streamlit). Es el archivo principal. |
| `esquema.py` | Define las tablas de la base de datos. |
| `cargador.py` | Carga los datos limpios a la base. |
| `limpiador_flexus.py` | Limpia y valida los Excel de Flexus. |
| `metricas.py` | Toda la lógica de reportes y análisis. |
| `requirements.txt` | Lista de librerías que el servidor debe instalar. |

> IMPORTANTE: los nombres de archivo deben quedar EXACTAMENTE así.
> No renombrar a `esquema_2.py`, `cargador_3.py`, etc. — la app los busca
> por estos nombres y si no coinciden, falla.

## Cómo correrlo localmente (en tu PC)

1. Instalar dependencias (una sola vez):
   ```
   pip install -r requirements.txt
   ```
2. Levantar la app:
   ```
   streamlit run app.py
   ```
3. Se abre solo en el navegador, en http://localhost:8501

## Usuario inicial

- Usuario: **admin**
- Contraseña: **admin123**

Cambiala después de entrar. Desde la sección "Usuarios" (visible solo para
el admin) se dan de alta los administrativos con rol "Consulta".

## Roles

- **Administrador**: carga archivos de Flexus, gestiona usuarios, ve todo.
- **Consulta / Auditor**: ven reportes y cargan cotizaciones (dólar/IPC),
  pero NO pueden cargar archivos de Flexus.
