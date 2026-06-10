# Reportabilidad VCA 5400

Aplicación web Flask lista para Render.com. Permite subir una curva de poblamiento y una o varias planillas de reportabilidad/dotación, cruza los ID planificados contra los ID reportados y genera un Excel de salida con formato estándar.

## Funciones

- Lee automáticamente la hoja `Fcst_Autorizado VCA` de la curva.
- Detecta `ID de la solicitud`, empresa, número de contrato y columnas de dotación diaria.
- Lee reportabilidad de empresa desde `Hoja1` o la primera hoja compatible.
- Reconoce el ID de reportabilidad desde `N° DE ID` o nombres equivalentes.
- Informa empresas con dotación planificada que no aparecen en la reportabilidad cargada.
- Informa IDs planificados no reportados.
- Exporta un Excel con estas hojas:
  - `Formato_Final`
  - `Empresas_Sin_Reportabilidad`
  - `IDs_Planificados_No_Reportados`
  - `Resumen`

## Columnas del formato final

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

## Ejecución local

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
python app.py
```

Luego abre `http://127.0.0.1:5000`.

## Deploy en Render.com

1. Sube esta carpeta a un repositorio GitHub.
2. En Render, crea un **Web Service**.
3. Conecta el repositorio.
4. Verifica que los archivos estén en la raíz del repositorio: `app.py`, `requirements.txt`, `.python-version`, `render.yaml`, `templates/` y `static/`.
5. Usa estos comandos:
   - Build Command: `pip install --upgrade pip && pip install -r requirements.txt`
   - Start Command: `gunicorn app:app`
6. Configura la variable de entorno `PYTHON_VERSION=3.11.11` si Render no toma automáticamente el archivo `.python-version`.

## Nota operativa

Render usa almacenamiento efímero en su plan estándar. Los archivos generados quedan disponibles para descarga inmediatamente después del procesamiento, pero no debe asumirse conservación permanente.


## Corrección de despliegue en Render

Este paquete incluye `.python-version` con `3.11.11` y `render.yaml` con `PYTHON_VERSION=3.11.11`. Esto evita que Render use Python 3.14 por defecto y compile pandas desde fuente.

Si el despliegue anterior falló con `metadata-generation-failed` en pandas, vuelve a subir este paquete, asegúrate de que `.python-version` esté en la raíz del repositorio y ejecuta un nuevo deploy manual.
