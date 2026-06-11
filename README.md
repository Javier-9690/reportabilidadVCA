# Reportabilidad VCA 5400 - Semanas acumulativas

Aplicación Flask para Render.com orientada a controlar reportabilidades semanales de Campamento 5400.

## Funcionalidad principal

- Crear una semana con su curva de poblamiento.
- Leer la hoja `Fcst_Autorizado VCA` desde la curva.
- Guardar la planificación semanal en PostgreSQL.
- Cargar reportabilidades de empresas de forma incremental durante varios días.
- Eliminar una reportabilidad específica sin borrar la semana completa.
- Reemplazar la curva de una semana manteniendo las reportabilidades ya cargadas.
- Recalcular automáticamente:
  - IDs planificados.
  - IDs reportados.
  - IDs faltantes.
  - Empresas con dotación planificada sin reportabilidad.
- Exportar el Excel final normalizado.

## Formato final exportado

Columnas:

- ID
- MODULO
- RUT (CON GUION)
- NOMBRE COMPLETO
- EMPRESA
- NUMERO DE CONTRATO
- GERENCIA
- SISTEMA DE TURNO
- CO MEL
- GENERO
- NOMBRE DE TURNO

## Render

Build Command:

```bash
pip install --upgrade pip && pip install -r requirements.txt
```

Start Command:

```bash
gunicorn --workers 1 --threads 2 --timeout 240 app:app
```

Variables de entorno:

```text
PYTHON_VERSION=3.11.11
DATABASE_URL=<cadena PostgreSQL de Render>
SECRET_KEY=<clave segura>
```

## Dependencias

No usa pandas ni openpyxl para lectura de archivos de entrada. La lectura XLSX se realiza directamente sobre el XML interno del archivo para evitar timeouts con planillas pesadas.

## Modelo acumulativo

La unidad principal ya no es una carga aislada, sino una `semana`:

1. Se crea la semana con una curva.
2. La semana queda guardada en PostgreSQL.
3. Puedes agregar reportabilidades a esa misma semana durante varios días.
4. Cada archivo cargado queda individualizado.
5. Puedes eliminar un archivo puntual si una empresa envió una versión errónea.
6. El Excel se genera siempre desde el estado actual de la semana.

## Corrección de exportación

- La hoja `Formato_Final` conserva todas las filas/personas cargadas desde las reportabilidades. El ID se usa para medir cobertura contra curva, pero no elimina personas repetidas bajo un mismo ID.
- La hoja `IDs_Planificados_No_Reportados` incluye la columna `CO MEL`, tomada desde la curva, para facilitar el contacto del responsable.
