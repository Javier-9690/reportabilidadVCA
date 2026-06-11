# Reportabilidad VCA 5400

Aplicación Flask para Render.com que cruza una curva de poblamiento con archivos de reportabilidad/dotación por empresa.

## Ajuste visual

Esta versión mantiene la funcionalidad rápida sin pandas ni openpyxl para lectura, pero usa solamente el color corporativo rojo y los logos importados en `static/img/`.

No incluye menús ni opciones falsas sin funcionalidad.

## Qué hace

- Carga la curva de poblamiento.
- Lee automáticamente la hoja `Fcst_Autorizado VCA`.
- Carga una o varias reportabilidades de empresa.
- Reconoce el ID de la curva como `ID de la solicitud`.
- Reconoce el ID de reportabilidad como `N° DE ID` o equivalente.
- Indica empresas que tienen dotación planificada en curva pero no enviaron reportabilidad.
- Exporta un Excel con formato final y hojas de control.

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
