import os
import re
import uuid
import json
import unicodedata
from datetime import datetime
from pathlib import Path

import pandas as pd
from flask import Flask, render_template, request, redirect, url_for, send_file, flash

BASE_DIR = Path(__file__).resolve().parent
UPLOAD_DIR = BASE_DIR / 'uploads'
OUTPUT_DIR = BASE_DIR / 'outputs'
UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'reportabilidad-vca-5400')
app.config['MAX_CONTENT_LENGTH'] = 60 * 1024 * 1024

TARGET_COLUMNS = [
    'ID', 'MODULO', 'RUT (CON GUION)', 'NOMBRE COMPLETO', 'EMPRESA',
    'NUMERO DE CONTRATO', 'GERENCIA', 'SISTEMA DE TURNO', 'CO MEL',
    'GENERO', 'NOMBRE DE TURNO'
]

SYNONYMS = {
    'ID': ['n de id', 'n° de id', 'nº de id', 'id', 'id de la solicitud', 'numero de id', 'n de solicitud'],
    'MODULO': ['modulo', 'módulo', 'habitacion', 'habitación', 'cama a utilizar', 'cama'],
    'RUT (CON GUION)': ['rut', 'run'],
    'NOMBRE COMPLETO': ['nombre huesped', 'nombre huésped', 'nombre completo', 'trabajador', 'persona', 'nombre'],
    'EMPRESA': ['empresa', 'razon social', 'razón social'],
    'NUMERO DE CONTRATO': ['n de contrato', 'n° de contrato', 'nº de contrato', 'numero de contrato', 'número contrato', 'numero contrato', 'contrato'],
    'GERENCIA': ['gerencia', 'gerencia general', 'area', 'área'],
    'SISTEMA DE TURNO': ['sistema de turno', 'turno'],
    'CO MEL': ['co mel', 'nombre contract owner', 'contract owner', 'correo co', 'spa co', 'spa/ co'],
    'GENERO': ['genero', 'género', 'sexo'],
    'NOMBRE DE TURNO': ['nombre turno', 'nombre de turno', 'turno actual'],
}

CURVE_ID_COLS = ['id de la solicitud', 'n° de id', 'nº de id', 'n de id', 'id']
CURVE_COMPANY_COLS = ['empresa', 'razon social', 'razón social']
CURVE_CONTRACT_COLS = ['número contrato', 'numero contrato', 'n de contrato', 'n° de contrato', 'contrato']


def clean_text(value):
    if pd.isna(value):
        return ''
    text = str(value).strip()
    text = re.sub(r'\s+', ' ', text)
    return text


def normalize(value):
    text = clean_text(value).lower()
    text = unicodedata.normalize('NFKD', text).encode('ascii', 'ignore').decode('ascii')
    text = text.replace('°', '').replace('º', '')
    text = re.sub(r'[^a-z0-9]+', ' ', text)
    return re.sub(r'\s+', ' ', text).strip()


def normalize_id(value):
    text = clean_text(value).upper()
    text = text.replace('.0', '') if text.endswith('.0') else text
    text = re.sub(r'\s+', '', text)
    return text


def normalize_company(value):
    text = normalize(value)
    # Quita razones sociales frecuentes para comparar "El Sauce" vs "Constructora El Sauce S.A".
    legal_tokens = [
        'constructora', 'sociedad', 'anonima', 'servicios', 'servicio', 'empresa',
        's', 'a', 'sa', 'spa', 'sac', 'ltda', 'limitada', 'chile'
    ]
    for token in legal_tokens:
        text = re.sub(rf'\b{token}\b', ' ', text)
    return re.sub(r'\s+', ' ', text).strip()


def format_rut(value):
    raw = clean_text(value).upper().replace('.', '').replace(' ', '')
    if not raw:
        return ''
    raw = raw.replace('–', '-').replace('—', '-')
    if '-' in raw:
        num, dv = raw.rsplit('-', 1)
    else:
        num, dv = raw[:-1], raw[-1]
    num = re.sub(r'\D', '', num)
    dv = re.sub(r'[^0-9K]', '', dv)
    if not num or not dv:
        return raw
    return f'{num}-{dv}'


def find_header_row(path, sheet_name=0, required_any=None, max_rows=30):
    raw = pd.read_excel(path, sheet_name=sheet_name, header=None, nrows=max_rows, dtype=str)
    required_any = [normalize(x) for x in (required_any or [])]
    best_idx, best_score = 0, -1
    for idx, row in raw.iterrows():
        cells = [normalize(x) for x in row.tolist()]
        score = sum(any(req == c or req in c for c in cells) for req in required_any)
        score += sum(1 for c in cells if c)
        if score > best_score:
            best_idx, best_score = idx, score
    return best_idx


def read_excel_smart(path, sheet_name=0, required_any=None):
    header = find_header_row(path, sheet_name, required_any)
    df = pd.read_excel(path, sheet_name=sheet_name, header=header, dtype=str)
    df = df.dropna(how='all')
    df.columns = [clean_text(c) for c in df.columns]
    df = df.loc[:, [c for c in df.columns if c and not c.startswith('Unnamed')]]
    return df


def find_sheet(path, preferred):
    xls = pd.ExcelFile(path)
    for name in xls.sheet_names:
        if normalize(name) == normalize(preferred):
            return name
    for name in xls.sheet_names:
        if normalize(preferred) in normalize(name):
            return name
    return xls.sheet_names[0]


def pick_column(df, candidates):
    norm_map = {normalize(c): c for c in df.columns}
    for cand in candidates:
        nc = normalize(cand)
        if nc in norm_map:
            return norm_map[nc]
    for cand in candidates:
        nc = normalize(cand)
        for ncol, original in norm_map.items():
            if nc and (nc in ncol or ncol in nc):
                return original
    return None


def numeric_plan_columns(df):
    cols = []
    for c in df.columns:
        nc = normalize(c)
        if re.fullmatch(r'\d+(\.\d+)?', clean_text(c)) or re.fullmatch(r'\d{5}', clean_text(c)) or nc in ['lunes','martes','miercoles','jueves','viernes','sabado','domingo']:
            series = pd.to_numeric(df[c], errors='coerce')
            if series.notna().sum() > 0:
                cols.append(c)
    if not cols:
        # fallback: all mostly numeric columns except obvious identifiers/contracts
        for c in df.columns:
            if c in df.columns[:17]:
                continue
            series = pd.to_numeric(df[c], errors='coerce')
            if series.notna().sum() > 0:
                cols.append(c)
    return cols


def process_files(curve_path, report_paths):
    curve_sheet = find_sheet(curve_path, 'Fcst_Autorizado VCA')
    curve = read_excel_smart(curve_path, curve_sheet, required_any=['ID de la solicitud', 'Empresa', 'Número Contrato'])
    curve_id_col = pick_column(curve, CURVE_ID_COLS)
    curve_company_col = pick_column(curve, CURVE_COMPANY_COLS)
    curve_contract_col = pick_column(curve, CURVE_CONTRACT_COLS)
    if not curve_id_col or not curve_company_col:
        raise ValueError('No se pudo detectar ID de solicitud y Empresa en la hoja Fcst_Autorizado VCA.')

    plan_cols = numeric_plan_columns(curve)
    curve['_ID_NORM'] = curve[curve_id_col].map(normalize_id)
    curve['_EMPRESA_NORM'] = curve[curve_company_col].map(normalize_company)
    if plan_cols:
        plan_values = curve[plan_cols].apply(pd.to_numeric, errors='coerce').fillna(0)
        curve['_DOTACION_PLANIFICADA'] = plan_values.sum(axis=1)
    else:
        curve['_DOTACION_PLANIFICADA'] = 1
    planned = curve[(curve['_ID_NORM'] != '') & (curve['_DOTACION_PLANIFICADA'] > 0)].copy()

    reports = []
    for path in report_paths:
        sheet = find_sheet(path, 'Hoja1')
        report = read_excel_smart(path, sheet, required_any=['N° DE ID', 'RUT', 'Nombre Huésped', 'Empresa'])
        report['_ARCHIVO_ORIGEN'] = Path(path).name
        reports.append(report)
    if reports:
        report = pd.concat(reports, ignore_index=True)
    else:
        report = pd.DataFrame()

    if report.empty:
        transformed = pd.DataFrame(columns=TARGET_COLUMNS)
        reported_ids = set()
        reported_companies = set()
    else:
        report_id_col = pick_column(report, SYNONYMS['ID'])
        if not report_id_col:
            raise ValueError('No se pudo detectar la columna ID en la reportabilidad. Se espera N° DE ID o equivalente.')
        report['_ID_NORM'] = report[report_id_col].map(normalize_id)
        emp_col = pick_column(report, SYNONYMS['EMPRESA'])
        report['_EMPRESA_NORM'] = report[emp_col].map(normalize_company) if emp_col else ''
        reported_ids = set(report['_ID_NORM'].dropna()) - {''}
        reported_companies = set(report['_EMPRESA_NORM'].dropna()) - {''}

        output = {}
        for target in TARGET_COLUMNS:
            col = pick_column(report, SYNONYMS[target])
            if col:
                output[target] = report[col].map(clean_text)
            else:
                output[target] = ''
        transformed = pd.DataFrame(output)
        transformed['ID'] = transformed['ID'].map(normalize_id)
        transformed['RUT (CON GUION)'] = transformed['RUT (CON GUION)'].map(format_rut)
        transformed = transformed.dropna(how='all')
        transformed = transformed[transformed['ID'] != '']
        transformed = transformed[TARGET_COLUMNS]

    planned_ids = set(planned['_ID_NORM'].dropna()) - {''}
    missing_ids = planned[~planned['_ID_NORM'].isin(reported_ids)].copy()
    company_original = planned.groupby('_EMPRESA_NORM', dropna=False)[curve_company_col].first().rename('EMPRESA')
    missing_companies = planned.groupby('_EMPRESA_NORM', dropna=False).agg(
        ids_planificados=('_ID_NORM', 'nunique'),
        dotacion_planificada=('_DOTACION_PLANIFICADA', 'sum')
    ).join(company_original).reset_index()
    missing_companies = missing_companies[~missing_companies['_EMPRESA_NORM'].isin(reported_companies)].copy()
    missing_companies = missing_companies[['EMPRESA', 'ids_planificados', 'dotacion_planificada']].sort_values('EMPRESA')

    comparison_cols = [curve_id_col, curve_company_col]
    if curve_contract_col:
        comparison_cols.append(curve_contract_col)
    comparison_cols += ['_DOTACION_PLANIFICADA']
    missing_ids_out = missing_ids[comparison_cols].rename(columns={
        curve_id_col: 'ID',
        curve_company_col: 'EMPRESA',
        curve_contract_col or '': 'NUMERO DE CONTRATO',
        '_DOTACION_PLANIFICADA': 'DOTACION_PLANIFICADA'
    })
    if 'NUMERO DE CONTRATO' not in missing_ids_out.columns:
        missing_ids_out['NUMERO DE CONTRATO'] = ''
    missing_ids_out = missing_ids_out[['ID','EMPRESA','NUMERO DE CONTRATO','DOTACION_PLANIFICADA']].sort_values(['EMPRESA','ID'])

    summary = {
        'curve_sheet': curve_sheet,
        'planned_ids': int(len(planned_ids)),
        'reported_ids': int(len(reported_ids)),
        'matched_ids': int(len(planned_ids.intersection(reported_ids))),
        'missing_ids': int(len(planned_ids - reported_ids)),
        'planned_companies': int(planned['_EMPRESA_NORM'].nunique()),
        'reported_companies': int(len(reported_companies)),
        'missing_companies': int(len(missing_companies)),
        'transformed_rows': int(len(transformed)),
    }
    return transformed, missing_ids_out, missing_companies, summary


def write_output(path, transformed, missing_ids, missing_companies, summary):
    with pd.ExcelWriter(path, engine='xlsxwriter') as writer:
        transformed.to_excel(writer, sheet_name='Formato_Final', index=False)
        missing_companies.to_excel(writer, sheet_name='Empresas_Sin_Reportabilidad', index=False)
        missing_ids.to_excel(writer, sheet_name='IDs_Planificados_No_Reportados', index=False)
        pd.DataFrame([summary]).to_excel(writer, sheet_name='Resumen', index=False)
        workbook = writer.book
        header_fmt = workbook.add_format({'bold': True, 'font_color': 'white', 'bg_color': '#0f2f55', 'border': 1, 'align': 'center', 'valign': 'vcenter'})
        cell_fmt = workbook.add_format({'border': 1, 'valign': 'top'})
        warn_fmt = workbook.add_format({'border': 1, 'bg_color': '#fff4dc'})
        for sheet_name, df in [('Formato_Final', transformed), ('Empresas_Sin_Reportabilidad', missing_companies), ('IDs_Planificados_No_Reportados', missing_ids), ('Resumen', pd.DataFrame([summary]))]:
            ws = writer.sheets[sheet_name]
            ws.freeze_panes(1, 0)
            ws.autofilter(0, 0, max(len(df), 1), max(len(df.columns)-1, 0))
            for col_num, value in enumerate(df.columns):
                ws.write(0, col_num, value, header_fmt)
                width = min(max(len(str(value)) + 2, 14), 38)
                if len(df) > 0:
                    width = min(max(width, int(df.iloc[:, col_num].astype(str).str.len().quantile(.9)) + 2), 45)
                ws.set_column(col_num, col_num, width, warn_fmt if 'Sin' in sheet_name or 'No_Reportados' in sheet_name else cell_fmt)


@app.route('/', methods=['GET'])
def index():
    return render_template('index.html')


@app.route('/procesar', methods=['POST'])
def procesar():
    try:
        curva = request.files.get('curva')
        reports = request.files.getlist('reportabilidad')
        if not curva or not curva.filename:
            flash('Debes subir la curva de poblamiento.', 'error')
            return redirect(url_for('index'))
        reports = [f for f in reports if f and f.filename]
        if not reports:
            flash('Debes subir al menos un archivo de reportabilidad/dotación.', 'error')
            return redirect(url_for('index'))
        job = uuid.uuid4().hex[:10]
        curve_path = UPLOAD_DIR / f'{job}_curva.xlsx'
        curva.save(curve_path)
        report_paths = []
        for i, f in enumerate(reports, start=1):
            path = UPLOAD_DIR / f'{job}_report_{i}.xlsx'
            f.save(path)
            report_paths.append(path)
        transformed, missing_ids, missing_companies, summary = process_files(curve_path, report_paths)
        output_path = OUTPUT_DIR / f'reportabilidad_vca_{job}.xlsx'
        write_output(output_path, transformed, missing_ids, missing_companies, summary)
        meta = {
            'job': job,
            'output': output_path.name,
            'summary': summary,
            'preview': transformed.head(20).to_dict(orient='records'),
            'missing_ids': missing_ids.head(50).to_dict(orient='records'),
            'missing_companies': missing_companies.head(50).to_dict(orient='records'),
            'created_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        }
        with open(OUTPUT_DIR / f'{job}.json', 'w', encoding='utf-8') as fh:
            json.dump(meta, fh, ensure_ascii=False, indent=2)
        return redirect(url_for('resultado', job=job))
    except Exception as exc:
        flash(f'Error al procesar archivos: {exc}', 'error')
        return redirect(url_for('index'))


@app.route('/resultado/<job>')
def resultado(job):
    meta_path = OUTPUT_DIR / f'{job}.json'
    if not meta_path.exists():
        flash('No se encontró el procesamiento solicitado.', 'error')
        return redirect(url_for('index'))
    with open(meta_path, encoding='utf-8') as fh:
        meta = json.load(fh)
    return render_template('resultado.html', meta=meta, target_columns=TARGET_COLUMNS)


@app.route('/descargar/<job>')
def descargar(job):
    meta_path = OUTPUT_DIR / f'{job}.json'
    if not meta_path.exists():
        flash('No se encontró el archivo de salida.', 'error')
        return redirect(url_for('index'))
    with open(meta_path, encoding='utf-8') as fh:
        meta = json.load(fh)
    path = OUTPUT_DIR / meta['output']
    return send_file(path, as_attachment=True, download_name='Reportabilidad_VCA_Formato_Final.xlsx')


if __name__ == '__main__':
    app.run(debug=True)
