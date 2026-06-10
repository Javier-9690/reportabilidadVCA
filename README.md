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
4. Usa estos comandos:
   - Build Command: `pip install -r requirements.txt`
   - Start Command: `gunicorn app:app`
5. Variable recomendada:
   - `SECRET_KEY`: una clave aleatoria.

## Nota operativa

Render usa almacenamiento efímero en su plan estándar. Los archivos generados quedan disponibles para descarga inmediatamente después del procesamiento, pero no debe asumirse conservación permanente.
