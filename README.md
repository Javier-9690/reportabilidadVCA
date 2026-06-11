# Reportabilidad VCA 5400 - PostgreSQL

Aplicación Flask para Render que procesa curva de poblamiento y reportabilidades por empresa, cruza por ID y guarda el resultado en PostgreSQL.

## Funciones

- Importar curva de poblamiento.
- Leer hoja `Fcst_Autorizado VCA`.
- Importar una o varias reportabilidades/dotaciones.
- Cruzar `ID de la solicitud` contra `N° DE ID`.
- Detectar IDs planificados no reportados.
- Detectar empresas con dotación planificada sin reportabilidad.
- Exportar Excel final.
- Guardar historial en PostgreSQL.
- Descargar reportes guardados desde la BD.
- Eliminar reportabilidades guardadas.

## Render

Build Command:

```bash
pip install --upgrade pip && pip install -r requirements.txt
```

Start Command:

```bash
gunicorn --workers 1 --threads 2 --timeout 240 app:app
```

Variables de entorno necesarias:

```text
PYTHON_VERSION=3.11.11
DATABASE_URL=<Internal Database URL de PostgreSQL en Render>
SECRET_KEY=<valor seguro>
```

## Tablas creadas automáticamente

- `reportabilidad_procesos`
- `reportabilidad_formato`
- `reportabilidad_missing_ids`
- `reportabilidad_missing_empresas`

La app crea estas tablas al iniciar si `DATABASE_URL` está configurada.
