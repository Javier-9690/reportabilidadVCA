# Reportabilidad VCA 5400

Aplicación Flask para Render.com que cruza una curva de poblamiento con archivos de reportabilidad/dotación por empresa.

## Qué hace

- Carga la curva de poblamiento.
- Toma automáticamente la hoja `Fcst_Autorizado VCA`.
- Carga una o varias reportabilidades de empresas.
- Reconoce el ID de la curva como `ID de la solicitud`.
- Reconoce el ID de reportabilidad como `N° DE ID` o equivalente.
- Indica empresas que tienen dotación planificada en curva pero no enviaron reportabilidad.
- Exporta un Excel con formato final y hojas de control.

## Salida Excel

La hoja `Formato_Final` contiene:

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

Además incluye:

- `Empresas_Sin_Reportabilidad`
- `IDs_Planificados_No_Reportados`
- `Resumen`

## Versión optimizada

Esta versión no usa pandas, numpy ni openpyxl para leer los archivos de entrada. Lee directamente el XML interno de los XLSX/XLSM para evitar timeouts en Render con curvas pesadas o con muchos estilos.

## Render

Build Command:

```bash
pip install --upgrade pip && pip install -r requirements.txt
```

Start Command:

```bash
gunicorn --workers 1 --threads 2 --timeout 240 app:app
```

Environment Variable recomendada:

```text
PYTHON_VERSION=3.11.11
```

## Archivos aceptados

- `.xlsx`
- `.xlsm`

Si tienes un archivo `.xls`, guárdalo desde Excel como `.xlsx` antes de subirlo.
