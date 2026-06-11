import json
import os
import re
import unicodedata
import uuid
import zipfile
import xml.etree.ElementTree as ET
from datetime import date, datetime
from pathlib import Path
from typing import Any, Dict, Iterable, Iterator, List, Optional, Sequence, Tuple
from xml.sax.saxutils import escape

from flask import Flask, flash, redirect, render_template, request, send_file, url_for
from werkzeug.utils import secure_filename

BASE_DIR = Path(__file__).resolve().parent
UPLOAD_DIR = BASE_DIR / "uploads"
OUTPUT_DIR = BASE_DIR / "outputs"
UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "reportabilidad-vca-5400")
app.config["MAX_CONTENT_LENGTH"] = int(os.environ.get("MAX_UPLOAD_MB", "80")) * 1024 * 1024

# Lector XLSX ultraliviano: no usa pandas ni openpyxl para leer entradas.
# Esto evita que Render se quede detenido cargando estilos o miles de celdas formateadas.
DEFAULT_SCAN_MAX_COLS = int(os.environ.get("SCAN_MAX_COLS", "260"))
REPORT_SCAN_MAX_COLS = int(os.environ.get("REPORT_SCAN_MAX_COLS", "120"))

TARGET_COLUMNS = [
    "ID",
    "MODULO",
    "RUT (CON GUION)",
    "NOMBRE COMPLETO",
    "EMPRESA",
    "NUMERO DE CONTRATO",
    "GERENCIA",
    "SISTEMA DE TURNO",
    "CO MEL",
    "GENERO",
    "NOMBRE DE TURNO",
]

SYNONYMS = {
    "ID": ["n de id", "n° de id", "nº de id", "id", "id de la solicitud", "numero de id", "n de solicitud"],
    "MODULO": ["modulo", "módulo", "habitacion", "habitación", "cama a utilizar", "cama"],
    "RUT (CON GUION)": ["rut", "run"],
    "NOMBRE COMPLETO": ["nombre huesped", "nombre huésped", "nombre completo", "trabajador", "persona", "nombre"],
    "EMPRESA": ["empresa", "razon social", "razón social"],
    "NUMERO DE CONTRATO": ["n de contrato", "n° de contrato", "nº de contrato", "numero de contrato", "número contrato", "numero contrato", "contrato"],
    "GERENCIA": ["gerencia", "gerencia general", "area", "área"],
    "SISTEMA DE TURNO": ["sistema de turno", "turno"],
    "CO MEL": ["co mel", "nombre contract owner", "contract owner", "correo co", "spa co", "spa/ co", "spa co"],
    "GENERO": ["genero", "género", "sexo"],
    "NOMBRE DE TURNO": ["nombre turno", "nombre de turno", "turno actual"],
}

CURVE_ID_COLS = ["id de la solicitud", "n° de id", "nº de id", "n de id", "id"]
CURVE_COMPANY_COLS = ["empresa", "razon social", "razón social"]
CURVE_CONTRACT_COLS = ["número contrato", "numero contrato", "n de contrato", "n° de contrato", "contrato"]
WEEKDAY_NAMES = {"lunes", "martes", "miercoles", "jueves", "viernes", "sabado", "domingo"}
ALLOWED_EXTENSIONS = {".xlsx", ".xlsm"}

NS = {
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "rel": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "pkgrel": "http://schemas.openxmlformats.org/package/2006/relationships",
}
REL_ID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"


def clean_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, date):
        return value.strftime("%Y-%m-%d")
    text = str(value).strip()
    if text.endswith(".0") and re.fullmatch(r"\d+\.0", text):
        text = text[:-2]
    return re.sub(r"\s+", " ", text)


def normalize(value: Any) -> str:
    text = clean_text(value).lower()
    text = unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode("ascii")
    text = text.replace("°", "").replace("º", "")
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def normalize_id(value: Any) -> str:
    text = clean_text(value).upper()
    text = text.replace("–", "-").replace("—", "-")
    text = re.sub(r"\s+", "", text)
    return text


def normalize_company(value: Any) -> str:
    text = normalize(value)
    legal_tokens = [
        "constructora", "sociedad", "anonima", "servicios", "servicio", "empresa",
        "mineria", "minera", "chile", "spa", "sa", "s", "a", "sac", "ltda", "limitada",
    ]
    for token in legal_tokens:
        text = re.sub(rf"\b{re.escape(token)}\b", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def format_rut(value: Any) -> str:
    raw = clean_text(value).upper().replace(".", "").replace(" ", "")
    raw = raw.replace("–", "-").replace("—", "-")
    if not raw:
        return ""
    if "-" in raw:
        num, dv = raw.rsplit("-", 1)
    else:
        num, dv = raw[:-1], raw[-1]
    num = re.sub(r"\D", "", num)
    dv = re.sub(r"[^0-9K]", "", dv)
    if not num or not dv:
        return raw
    return f"{num}-{dv}"


def to_number(value: Any) -> Optional[float]:
    if value is None or value == "":
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = clean_text(value).replace(".", "").replace(",", ".")
    try:
        return float(text)
    except ValueError:
        return None


def safe_filename(prefix: str, original: str, job: str, index: Optional[int] = None) -> Path:
    suffix = Path(original).suffix.lower()
    if suffix not in ALLOWED_EXTENSIONS:
        raise ValueError("Solo se aceptan archivos Excel modernos: .xlsx o .xlsm. Guarda el archivo .xls como .xlsx antes de subirlo.")
    safe = secure_filename(original) or f"archivo{suffix}"
    name = f"{job}_{prefix}_{index}_{safe}" if index is not None else f"{job}_{prefix}_{safe}"
    return UPLOAD_DIR / name


def col_to_index(col_letters: str) -> int:
    total = 0
    for char in col_letters:
        if char.isalpha():
            total = total * 26 + (ord(char.upper()) - 64)
    return total


def index_to_col(index: int) -> str:
    letters = ""
    while index:
        index, rem = divmod(index - 1, 26)
        letters = chr(65 + rem) + letters
    return letters or "A"


def cell_ref_to_col(ref: str) -> int:
    match = re.match(r"([A-Za-z]+)", ref or "")
    return col_to_index(match.group(1)) if match else 0


def sheet_info(path: Path) -> List[Tuple[str, str]]:
    with zipfile.ZipFile(path) as z:
        workbook = ET.fromstring(z.read("xl/workbook.xml"))
        rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        rid_to_target = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels}
        out: List[Tuple[str, str]] = []
        for sheet in workbook.find("main:sheets", NS):
            name = sheet.attrib["name"]
            rid = sheet.attrib.get(REL_ID)
            target = rid_to_target.get(rid, "")
            if target.startswith("/"):
                xml_path = target.lstrip("/")
            elif target.startswith("xl/"):
                xml_path = target
            else:
                xml_path = "xl/" + target
            out.append((name, xml_path))
        return out


def find_sheet(path: Path, preferred: str) -> str:
    sheets = sheet_info(path)
    for name, _ in sheets:
        if normalize(name) == normalize(preferred):
            return name
    for name, _ in sheets:
        if normalize(preferred) in normalize(name):
            return name
    if not sheets:
        raise ValueError(f"El archivo {path.name} no contiene hojas visibles.")
    return sheets[0][0]


def sheet_xml_path(path: Path, sheet_name: str) -> str:
    for name, xml_path in sheet_info(path):
        if name == sheet_name:
            return xml_path
    raise ValueError(f"No se encontró la hoja {sheet_name} en {path.name}.")


def read_shared_strings(z: zipfile.ZipFile) -> List[str]:
    if "xl/sharedStrings.xml" not in z.namelist():
        return []
    strings: List[str] = []
    with z.open("xl/sharedStrings.xml") as fh:
        for _, elem in ET.iterparse(fh, events=("end",)):
            if elem.tag.endswith("}si"):
                parts: List[str] = []
                for child in elem.iter():
                    if child.tag.endswith("}t") and child.text:
                        parts.append(child.text)
                strings.append("".join(parts))
                elem.clear()
    return strings


def cell_value(cell: ET.Element, shared: Sequence[str]) -> str:
    cell_type = cell.attrib.get("t")
    if cell_type == "s":
        v = cell.find("main:v", NS)
        if v is not None and v.text is not None:
            try:
                return shared[int(v.text)]
            except (ValueError, IndexError):
                return v.text
        return ""
    if cell_type == "inlineStr":
        parts: List[str] = []
        for child in cell.iter():
            if child.tag.endswith("}t") and child.text:
                parts.append(child.text)
        return "".join(parts)
    if cell_type == "str":
        v = cell.find("main:v", NS)
        return v.text if v is not None and v.text is not None else ""
    if cell_type == "b":
        v = cell.find("main:v", NS)
        return "VERDADERO" if v is not None and v.text == "1" else "FALSO"
    v = cell.find("main:v", NS)
    return v.text if v is not None and v.text is not None else ""


def iter_sheet_rows(path: Path, sheet_name: str, max_cols: int = 260, max_row: Optional[int] = None) -> Iterator[Tuple[int, Dict[int, str]]]:
    xml_path = sheet_xml_path(path, sheet_name)
    with zipfile.ZipFile(path) as z:
        shared = read_shared_strings(z)
        with z.open(xml_path) as fh:
            for _, row in ET.iterparse(fh, events=("end",)):
                if not row.tag.endswith("}row"):
                    continue
                row_idx = int(row.attrib.get("r", "0") or 0)
                if max_row is not None and row_idx > max_row:
                    row.clear()
                    break
                values: Dict[int, str] = {}
                for cell in row:
                    if not cell.tag.endswith("}c"):
                        continue
                    col_idx = cell_ref_to_col(cell.attrib.get("r", ""))
                    if col_idx <= 0 or col_idx > max_cols:
                        continue
                    value = cell_value(cell, shared)
                    if value not in (None, ""):
                        values[col_idx] = value
                yield row_idx, values
                row.clear()


def score_header_row(values_by_col: Dict[int, Any], required_any: Sequence[str]) -> int:
    cells = [normalize(v) for c, v in values_by_col.items() if c <= 80]
    score = len([c for c in cells if c])
    for req in required_any:
        nreq = normalize(req)
        if any(nreq == c or (nreq and nreq in c) for c in cells):
            score += 100
    business_words = ["empresa", "contrato", "rut", "gerencia", "id", "nombre", "turno"]
    score += 20 * sum(any(word in c for c in cells) for word in business_words)
    return score


def find_header_row(path: Path, sheet_name: str, required_any: Sequence[str], max_rows: int = 40, max_cols: int = 260) -> int:
    best_row = 1
    best_score = -1
    for row_idx, values in iter_sheet_rows(path, sheet_name, max_cols=max_cols, max_row=max_rows):
        score = score_header_row(values, required_any)
        if score > best_score:
            best_row = row_idx
            best_score = score
    return best_row


def unique_header(base: Any, used: Dict[str, int], col_idx: int) -> str:
    header = clean_text(base) or f"COL_{col_idx}"
    if header not in used:
        used[header] = 1
        return header
    used[header] += 1
    return f"{header}__{used[header]}"


def read_sheet_records(
    path: Path,
    sheet_name: str,
    required_any: Sequence[str],
    max_cols: int,
    max_rows: Optional[int] = None,
) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]], int]:
    header_row = find_header_row(path, sheet_name, required_any, max_cols=max_cols)
    records: List[Dict[str, Any]] = []
    metas: List[Dict[str, Any]] = []
    used: Dict[str, int] = {}
    end_row = None if max_rows is None else header_row + max_rows

    for row_idx, values in iter_sheet_rows(path, sheet_name, max_cols=max_cols, max_row=end_row):
        if row_idx < header_row:
            continue
        if row_idx == header_row:
            for idx in range(1, max_cols + 1):
                raw = values.get(idx, "")
                header = unique_header(raw, used, idx)
                if header.startswith("COL_") and idx < 18:
                    continue
                metas.append({"col": idx, "key": header, "raw": raw, "norm": normalize(raw)})
            continue
        record = {meta["key"]: values.get(meta["col"], "") for meta in metas}
        if any(clean_text(v) for v in record.values()):
            records.append(record)
    return records, metas, header_row


def pick_column_from_headers(headers: Iterable[str], candidates: Sequence[str]) -> Optional[str]:
    norm_map = {normalize(h): h for h in headers}
    for candidate in candidates:
        nc = normalize(candidate)
        if nc in norm_map:
            return norm_map[nc]
    for candidate in candidates:
        nc = normalize(candidate)
        for ncol, original in norm_map.items():
            if nc and (nc in ncol or ncol in nc):
                return original
    return None


def extract_target_week(path: Path, sheet_name: str) -> Optional[int]:
    patterns = [path.name]
    for _, values in iter_sheet_rows(path, sheet_name, max_cols=12, max_row=6):
        patterns.extend(clean_text(v) for v in values.values() if v is not None)
    joined = " ".join(patterns)
    for regex in [r"Semana[_\s-]*(\d{1,2})", r"\bW\s*(\d{1,2})\b"]:
        match = re.search(regex, joined, flags=re.IGNORECASE)
        if match:
            return int(match.group(1))
    return None


def detect_plan_keys(path: Path, sheet_name: str, metas: List[Dict[str, Any]], header_row: int) -> Tuple[List[str], str]:
    target_week = extract_target_week(path, sheet_name)
    col_to_key = {m["col"]: m["key"] for m in metas}
    if target_week:
        semana_match: Optional[int] = None
        w_match: Optional[int] = None
        for row_idx, values in iter_sheet_rows(path, sheet_name, max_cols=DEFAULT_SCAN_MAX_COLS, max_row=max(header_row, 6)):
            if row_idx >= header_row:
                break
            for col_idx, value in values.items():
                text = normalize(value)
                if text == f"semana {target_week}" and col_idx >= 18:
                    semana_match = col_idx
                    break
                if text == f"w{target_week}" and col_idx >= 18:
                    w_match = col_idx
            if semana_match:
                break
        start_col = semana_match or w_match
        if start_col:
            keys = [col_to_key[c] for c in range(start_col, start_col + 7) if c in col_to_key]
            if keys:
                return keys, f"Semana {target_week}"

    keys: List[str] = []
    for meta in metas:
        col_idx = int(meta["col"])
        nraw = meta["norm"]
        if col_idx >= 18 and (nraw in WEEKDAY_NAMES or re.fullmatch(r"\d+( \d+)?", nraw)):
            keys.append(meta["key"])
    return keys, "columnas de fechas detectadas"


def max_planned_value(record: Dict[str, Any], plan_keys: Sequence[str]) -> float:
    values = [to_number(record.get(key)) for key in plan_keys]
    numeric = [v for v in values if v is not None and v > 0]
    return max(numeric) if numeric else 0.0


def sum_planned_days(record: Dict[str, Any], plan_keys: Sequence[str]) -> float:
    values = [to_number(record.get(key)) for key in plan_keys]
    return sum(v for v in values if v is not None and v > 0)


def row_to_output(report_record: Dict[str, Any], headers: Sequence[str]) -> Dict[str, str]:
    output: Dict[str, str] = {}
    for target in TARGET_COLUMNS:
        col = pick_column_from_headers(headers, SYNONYMS[target])
        value = report_record.get(col, "") if col else ""
        output[target] = clean_text(value)
    output["ID"] = normalize_id(output["ID"])
    output["RUT (CON GUION)"] = format_rut(output["RUT (CON GUION)"])
    return output


def process_files(curve_path: Path, report_paths: Sequence[Path]) -> Tuple[List[Dict[str, str]], List[Dict[str, Any]], List[Dict[str, Any]], Dict[str, Any]]:
    curve_sheet = find_sheet(curve_path, "Fcst_Autorizado VCA")
    curve_records, curve_metas, curve_header_row = read_sheet_records(
        curve_path,
        curve_sheet,
        required_any=["ID de la solicitud", "Empresa", "Número Contrato"],
        max_cols=DEFAULT_SCAN_MAX_COLS,
    )
    curve_headers = [m["key"] for m in curve_metas]
    curve_id_col = pick_column_from_headers(curve_headers, CURVE_ID_COLS)
    curve_company_col = pick_column_from_headers(curve_headers, CURVE_COMPANY_COLS)
    curve_contract_col = pick_column_from_headers(curve_headers, CURVE_CONTRACT_COLS)
    if not curve_id_col or not curve_company_col:
        raise ValueError("No se pudo detectar ID de la solicitud y Empresa en la hoja Fcst_Autorizado VCA.")

    plan_keys, plan_scope = detect_plan_keys(curve_path, curve_sheet, curve_metas, curve_header_row)
    if not plan_keys:
        raise ValueError("No se pudieron detectar columnas de dotación planificada en la curva.")

    planned_rows: List[Dict[str, Any]] = []
    for record in curve_records:
        row_id = normalize_id(record.get(curve_id_col))
        company = clean_text(record.get(curve_company_col))
        planned_max = max_planned_value(record, plan_keys)
        if row_id and planned_max > 0:
            planned_rows.append(
                {
                    "ID": row_id,
                    "EMPRESA": company,
                    "EMPRESA_NORM": normalize_company(company),
                    "NUMERO DE CONTRATO": clean_text(record.get(curve_contract_col)) if curve_contract_col else "",
                    "DOTACION_PLANIFICADA": int(planned_max) if planned_max.is_integer() else planned_max,
                    "DOTACION_DIAS_SUMADA": sum_planned_days(record, plan_keys),
                }
            )

    transformed: List[Dict[str, str]] = []
    reported_ids: set[str] = set()
    reported_companies: set[str] = set()

    for report_path in report_paths:
        report_sheet = find_sheet(report_path, "Hoja1")
        report_records, report_metas, _ = read_sheet_records(
            report_path,
            report_sheet,
            required_any=["N° DE ID", "RUT", "Nombre Huésped", "Empresa"],
            max_cols=REPORT_SCAN_MAX_COLS,
        )
        report_headers = [m["key"] for m in report_metas]
        report_id_col = pick_column_from_headers(report_headers, SYNONYMS["ID"])
        report_emp_col = pick_column_from_headers(report_headers, SYNONYMS["EMPRESA"])
        if not report_id_col:
            raise ValueError(f"No se pudo detectar la columna ID en {report_path.name}. Se espera N° DE ID o equivalente.")
        for record in report_records:
            row_id = normalize_id(record.get(report_id_col))
            if not row_id:
                continue
            reported_ids.add(row_id)
            if report_emp_col:
                reported_companies.add(normalize_company(record.get(report_emp_col)))
            out = row_to_output(record, report_headers)
            if out["ID"]:
                transformed.append(out)

    planned_ids = {row["ID"] for row in planned_rows}
    missing_ids = sorted(
        [row for row in planned_rows if row["ID"] not in reported_ids],
        key=lambda r: (normalize_company(r.get("EMPRESA")), r.get("ID", "")),
    )

    company_bucket: Dict[str, Dict[str, Any]] = {}
    for row in planned_rows:
        norm = row["EMPRESA_NORM"]
        if not norm:
            continue
        bucket = company_bucket.setdefault(
            norm,
            {"EMPRESA": row["EMPRESA"], "ids_planificados": set(), "dotacion_planificada": 0.0},
        )
        bucket["ids_planificados"].add(row["ID"])
        bucket["dotacion_planificada"] += float(row["DOTACION_PLANIFICADA"] or 0)

    missing_companies: List[Dict[str, Any]] = []
    for norm, bucket in company_bucket.items():
        if norm not in reported_companies:
            dot = bucket["dotacion_planificada"]
            missing_companies.append(
                {
                    "EMPRESA": bucket["EMPRESA"],
                    "ids_planificados": len(bucket["ids_planificados"]),
                    "dotacion_planificada": int(dot) if float(dot).is_integer() else round(dot, 2),
                }
            )
    missing_companies.sort(key=lambda r: normalize_company(r["EMPRESA"]))

    summary = {
        "curve_sheet": curve_sheet,
        "plan_scope": plan_scope,
        "planned_ids": len(planned_ids),
        "reported_ids": len(reported_ids),
        "matched_ids": len(planned_ids.intersection(reported_ids)),
        "missing_ids": len(planned_ids - reported_ids),
        "planned_companies": len(company_bucket),
        "reported_companies": len(reported_companies),
        "missing_companies": len(missing_companies),
        "transformed_rows": len(transformed),
        "plan_columns_used": len(plan_keys),
        "report_files": len(report_paths),
    }
    return transformed, missing_ids, missing_companies, summary


def xml_text(value: Any) -> str:
    return escape(clean_text(value), {'"': '&quot;'})


def sheet_xml(rows: List[Dict[str, Any]], columns: Sequence[str], title: str) -> str:
    data: List[List[Any]] = [list(columns)]
    data.extend([[row.get(col, "") for col in columns] for row in rows])
    max_row = max(1, len(data))
    max_col = max(1, len(columns))
    dim = f"A1:{index_to_col(max_col)}{max_row}"
    col_widths: List[int] = []
    for col in columns:
        max_len = len(clean_text(col))
        for row in rows[:250]:
            max_len = max(max_len, len(clean_text(row.get(col, ""))))
        col_widths.append(min(max(max_len + 2, 12), 45))
    cols_xml = "".join(f'<col min="{i}" max="{i}" width="{w}" customWidth="1"/>' for i, w in enumerate(col_widths, start=1))
    row_xml: List[str] = []
    for r_idx, row in enumerate(data, start=1):
        cells: List[str] = []
        for c_idx, value in enumerate(row, start=1):
            ref = f"{index_to_col(c_idx)}{r_idx}"
            style = ' s="1"' if r_idx == 1 else ' s="2"'
            cells.append(f'<c r="{ref}" t="inlineStr"{style}><is><t>{xml_text(value)}</t></is></c>')
        height = ' ht="24" customHeight="1"' if r_idx == 1 else ""
        row_xml.append(f'<row r="{r_idx}"{height}>{"".join(cells)}</row>')
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <dimension ref="{dim}"/>
  <sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>
  <cols>{cols_xml}</cols>
  <sheetData>{''.join(row_xml)}</sheetData>
  <autoFilter ref="{dim}"/>
</worksheet>'''


def workbook_xml(sheet_names: Sequence[str]) -> str:
    sheets = []
    for idx, name in enumerate(sheet_names, start=1):
        sheets.append(f'<sheet name="{xml_text(name)}" sheetId="{idx}" r:id="rId{idx}"/>')
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>{''.join(sheets)}</sheets>
</workbook>'''


def workbook_rels_xml(sheet_count: int) -> str:
    rels = []
    for idx in range(1, sheet_count + 1):
        rels.append(f'<Relationship Id="rId{idx}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{idx}.xml"/>')
    rels.append(f'<Relationship Id="rId{sheet_count + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>')
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{''.join(rels)}</Relationships>'''


def content_types_xml(sheet_count: int) -> str:
    overrides = ['<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>', '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>']
    for idx in range(1, sheet_count + 1):
        overrides.append(f'<Override PartName="/xl/worksheets/sheet{idx}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>')
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  {''.join(overrides)}
</Types>'''


def root_rels_xml() -> str:
    return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>'''


def styles_xml() -> str:
    return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="3"><font><sz val="11"/><name val="Calibri"/></font><font><b/><color rgb="FFFFFFFF"/><sz val="11"/><name val="Calibri"/></font><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF0F2F55"/><bgColor indexed="64"/></patternFill></fill></fills>
  <borders count="2"><border/><border><left style="thin"><color rgb="FFD9E4EF"/></left><right style="thin"><color rgb="FFD9E4EF"/></right><top style="thin"><color rgb="FFD9E4EF"/></top><bottom style="thin"><color rgb="FFD9E4EF"/></bottom></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf><xf numFmtId="0" fontId="2" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment vertical="top"/></xf></cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>'''


def write_output(path: Path, transformed: List[Dict[str, Any]], missing_ids: List[Dict[str, Any]], missing_companies: List[Dict[str, Any]], summary: Dict[str, Any]) -> None:
    sheets = [
        ("Formato_Final", transformed, TARGET_COLUMNS),
        ("Empresas_Sin_Reportabilidad", missing_companies, ["EMPRESA", "ids_planificados", "dotacion_planificada"]),
        ("IDs_Planificados_No_Reportados", missing_ids, ["ID", "EMPRESA", "NUMERO DE CONTRATO", "DOTACION_PLANIFICADA"]),
        ("Resumen", [summary], list(summary.keys())),
    ]
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types_xml(len(sheets)))
        z.writestr("_rels/.rels", root_rels_xml())
        z.writestr("xl/workbook.xml", workbook_xml([name for name, _, _ in sheets]))
        z.writestr("xl/_rels/workbook.xml.rels", workbook_rels_xml(len(sheets)))
        z.writestr("xl/styles.xml", styles_xml())
        for idx, (name, rows, columns) in enumerate(sheets, start=1):
            z.writestr(f"xl/worksheets/sheet{idx}.xml", sheet_xml(rows, columns, name))


@app.route("/", methods=["GET"])
def index():
    return render_template("index.html")


@app.route("/healthz", methods=["GET"])
def healthz():
    return {"status": "ok"}


@app.route("/procesar", methods=["POST"])
def procesar():
    try:
        curva = request.files.get("curva")
        reports = request.files.getlist("reportabilidad")
        if not curva or not curva.filename:
            flash("Debes subir la curva de poblamiento.", "error")
            return redirect(url_for("index"))
        reports = [f for f in reports if f and f.filename]
        if not reports:
            flash("Debes subir al menos un archivo de reportabilidad/dotación.", "error")
            return redirect(url_for("index"))

        job = uuid.uuid4().hex[:10]
        curve_path = safe_filename("curva", curva.filename, job)
        curva.save(curve_path)

        report_paths: List[Path] = []
        for idx, file_storage in enumerate(reports, start=1):
            path = safe_filename("report", file_storage.filename, job, idx)
            file_storage.save(path)
            report_paths.append(path)

        transformed, missing_ids, missing_companies, summary = process_files(curve_path, report_paths)
        output_path = OUTPUT_DIR / f"reportabilidad_vca_{job}.xlsx"
        write_output(output_path, transformed, missing_ids, missing_companies, summary)

        meta = {
            "job": job,
            "output": output_path.name,
            "summary": summary,
            "preview": transformed[:20],
            "missing_ids": missing_ids[:80],
            "missing_companies": missing_companies[:80],
            "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        }
        with open(OUTPUT_DIR / f"{job}.json", "w", encoding="utf-8") as fh:
            json.dump(meta, fh, ensure_ascii=False, indent=2)
        return redirect(url_for("resultado", job=job))
    except Exception as exc:
        app.logger.exception("Error al procesar archivos")
        flash(f"Error al procesar archivos: {exc}", "error")
        return redirect(url_for("index"))


@app.route("/resultado/<job>")
def resultado(job: str):
    meta_path = OUTPUT_DIR / f"{job}.json"
    if not meta_path.exists():
        flash("No se encontró el procesamiento solicitado.", "error")
        return redirect(url_for("index"))
    with open(meta_path, encoding="utf-8") as fh:
        meta = json.load(fh)
    return render_template("resultado.html", meta=meta, target_columns=TARGET_COLUMNS)


@app.route("/descargar/<job>")
def descargar(job: str):
    meta_path = OUTPUT_DIR / f"{job}.json"
    if not meta_path.exists():
        flash("No se encontró el archivo de salida.", "error")
        return redirect(url_for("index"))
    with open(meta_path, encoding="utf-8") as fh:
        meta = json.load(fh)
    path = OUTPUT_DIR / meta["output"]
    return send_file(path, as_attachment=True, download_name="Reportabilidad_VCA_Formato_Final.xlsx")


if __name__ == "__main__":
    app.run(debug=True)
