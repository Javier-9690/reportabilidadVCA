# Reportabilidad VCA 5400

Aplicación Flask para Render.com que cruza la curva de poblamiento con la reportabilidad/dotación de empresas.

## Funciones

- Carga curva de poblamiento Excel.
- Lee automáticamente la hoja `Fcst_Autorizado VCA`.
- Detecta la semana de planificación desde el archivo o encabezado (`Semana 26`, `W26`, etc.).
- Carga una o varias planillas de reportabilidad/dotación.
- Cruza por ID:
  - Curva: `ID de la solicitud`.
  - Reportabilidad: `N° DE ID` o equivalente.
- Informa empresas con dotación planificada que no enviaron reportabilidad.
- Exporta Excel final con:
  - `Formato_Final`
  - `Empresas_Sin_Reportabilidad`
  - `IDs_Planificados_No_Reportados`
  - `Resumen`

## Excel final

Columnas generadas en `Formato_Final`:

1. ID
2. MODULO
3. RUT (CON GUION)
4. NOMBRE COMPLETO
5. EMPRESA
6. NUMERO DE CONTRATO
7. GERENCIA
8. SISTEMA DE TURNO
9. CO MEL
10. GENERO
11. NOMBRE DE TURNO

## Render.com

### Environment Variable obligatoria

Agrega en Render:

```text
PYTHON_VERSION=3.11.11
```

### Build Command

```bash
pip install --upgrade pip && pip install --only-binary=:all: -r requirements.txt
```

### Start Command

```bash
gunicorn --workers 1 --threads 2 --timeout 180 app:app
```

## Nota técnica

Esta versión no usa `pandas` ni `numpy`. El procesamiento se realiza con `openpyxl` para evitar problemas de compilación y consumo de memoria en Render.

Si Render sigue mostrando rutas con `python3.14` en el log, significa que el servicio no tomó la variable `PYTHON_VERSION` o está desplegando otra raíz de proyecto.
