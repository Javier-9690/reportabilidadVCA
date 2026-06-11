import json
import os
import re
import ssl
import unicodedata
import uuid
import zipfile
import xml.etree.ElementTree as ET
from datetime import date, datetime
from difflib import SequenceMatcher
from pathlib import Path
from urllib.parse import parse_qs, unquote, urlparse
from typing import Any, Dict, Iterable, Iterator, List, Optional, Sequence, Tuple
from xml.sax.saxutils import escape

from flask import Flask, flash, redirect, render_template, request, send_file, url_for
from werkzeug.utils import secure_filename

try:
    import pg8000.dbapi as pgdb
except Exception:  # Permite ejecutar la app localmente sin PostgreSQL instalado.
    pgdb = None

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

REPAIR_COLUMNS = [
    "ID REPORTADO",
    "FILAS_AFECTADAS",
    "EMPRESA REPORTADA",
    "NOMBRE COMPLETO",
    "ARCHIVO",
    "ID SUGERIDO 1",
    "CONFIANZA 1",
    "MOTIVO 1",
    "EMPRESA CURVA 1",
    "CONTRATO CURVA 1",
    "SPA/ CO 1",
    "ID SUGERIDO 2",
    "CONFIANZA 2",
    "MOTIVO 2",
    "EMPRESA CURVA 2",
    "CONTRATO CURVA 2",
    "SPA/ CO 2",
    "ID SUGERIDO 3",
    "CONFIANZA 3",
    "MOTIVO 3",
    "EMPRESA CURVA 3",
    "CONTRATO CURVA 3",
    "SPA/ CO 3",
    "ACCION PROPUESTA",
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


def id_repair_key(value: Any) -> str:
    """Normaliza caracteres que suelen confundirse en IDs escritos manualmente."""
    text = normalize_id(value)
    return text.translate(str.maketrans({"O": "0", "I": "1", "L": "1", "S": "5", "B": "8", "Z": "2"}))


def split_id_parts(value: Any) -> Tuple[str, str]:
    text = normalize_id(value)
    match = re.match(r"^([A-Z]+)(.*)$", text)
    if not match:
        return "", text
    return match.group(1), match.group(2)


def is_single_transposition(a: str, b: str) -> bool:
    if len(a) != len(b) or a == b:
        return False
    diffs = [i for i, (ca, cb) in enumerate(zip(a, b)) if ca != cb]
    return len(diffs) == 2 and diffs[1] == diffs[0] + 1 and a[diffs[0]] == b[diffs[1]] and a[diffs[1]] == b[diffs[0]]


def id_similarity_score(report_id: str, planned_id: str) -> Tuple[int, str]:
    """Calcula similitud entre un ID reportado y uno planificado, con explicación legible."""
    a = normalize_id(report_id)
    b = normalize_id(planned_id)
    if not a or not b:
        return 0, "ID vacío"
    if a == b:
        return 100, "match exacto"

    ak = id_repair_key(a)
    bk = id_repair_key(b)
    base = SequenceMatcher(None, a, b).ratio() * 100
    key_score = SequenceMatcher(None, ak, bk).ratio() * 100
    score = max(base, key_score)
    reasons: List[str] = []

    prefix_a, rest_a = split_id_parts(a)
    prefix_b, rest_b = split_id_parts(b)
    key_prefix_a, key_rest_a = split_id_parts(ak)
    key_prefix_b, key_rest_b = split_id_parts(bk)

    if prefix_a and prefix_a == prefix_b:
        score += 10
        reasons.append("mismo prefijo")
    elif prefix_a[:2] and prefix_a[:2] == prefix_b[:2]:
        score += 5
        reasons.append("prefijo similar")
    else:
        score -= 12
        reasons.append("prefijo distinto")

    if len(rest_a) == len(rest_b) and rest_a:
        diffs = sum(1 for ca, cb in zip(rest_a, rest_b) if ca != cb)
        if diffs == 1:
            score += 12
            reasons.append("1 caracter distinto")
        elif diffs == 2 and is_single_transposition(rest_a, rest_b):
            score += 12
            reasons.append("posible transposición")
        elif diffs <= 2:
            score += 6
            reasons.append(f"{diffs} caracteres distintos")
    elif abs(len(a) - len(b)) == 1:
        score += 4
        reasons.append("posible caracter faltante/sobrante")
    else:
        score -= min(abs(len(a) - len(b)) * 4, 18)

    if ak == bk and a != b:
        score = max(score, 96)
        reasons.append("posible confusión O/0, I/1, S/5 u otro")
    elif key_prefix_a == key_prefix_b and key_rest_a and key_rest_b and len(key_rest_a) == len(key_rest_b):
        key_diffs = sum(1 for ca, cb in zip(key_rest_a, key_rest_b) if ca != cb)
        if key_diffs == 1:
            score += 5
            reasons.append("muy cercano tras normalizar caracteres")

    score = int(round(max(0, min(99, score))))
    reason = "; ".join(dict.fromkeys(reasons)) or "similitud textual"
    return score, reason


def confidence_label(score: int) -> str:
    if score >= 92:
        return f"Alta ({score}%)"
    if score >= 82:
        return f"Media ({score}%)"
    if score >= 70:
        return f"Baja ({score}%)"
    return f"Revisar ({score}%)"


def build_id_repair_suggestions(
    transformed_rows: List[Dict[str, Any]],
    plan_rows: List[Dict[str, Any]],
    reported_ids: set[str],
) -> List[Dict[str, Any]]:
    """Detecta IDs reportados que no existen en la curva y propone candidatos planificados.

    No corrige automáticamente: entrega sugerencias para validar contra empresa, contrato y SPA/CO.
    """
    planned_ids = {normalize_id(r.get("ID")) for r in plan_rows if normalize_id(r.get("ID"))}
    unmatched_ids = sorted([rid for rid in reported_ids if rid and rid not in planned_ids])
    if not unmatched_ids:
        return []

    plan_by_id = {normalize_id(r.get("ID")): r for r in plan_rows if normalize_id(r.get("ID"))}
    candidate_ids = sorted(planned_ids)
    rows_by_bad_id: Dict[str, List[Dict[str, Any]]] = {}
    for row in transformed_rows:
        rid = normalize_id(row.get("ID"))
        if rid in unmatched_ids:
            rows_by_bad_id.setdefault(rid, []).append(row)

    suggestions: List[Dict[str, Any]] = []
    for bad_id in unmatched_ids:
        rows = rows_by_bad_id.get(bad_id, [])
        first = rows[0] if rows else {}
        reported_company_norm = normalize_company(first.get("EMPRESA", ""))
        prefix_bad, _ = split_id_parts(bad_id)
        same_prefix = [cid for cid in candidate_ids if split_id_parts(cid)[0] == prefix_bad]
        similar_prefix = [cid for cid in candidate_ids if prefix_bad and split_id_parts(cid)[0][:2] == prefix_bad[:2]]
        same_company = [cid for cid in candidate_ids if reported_company_norm and plan_by_id.get(cid, {}).get("EMPRESA_NORM") == reported_company_norm]
        same_company_same_prefix = [cid for cid in same_company if split_id_parts(cid)[0] == prefix_bad]
        pool = same_company_same_prefix or same_company or same_prefix or similar_prefix or candidate_ids
        scored: List[Tuple[int, str, str]] = []
        for cid in pool:
            score, reason = id_similarity_score(bad_id, cid)
            plan = plan_by_id.get(cid, {})
            if reported_company_norm and plan.get("EMPRESA_NORM") == reported_company_norm:
                score = min(99, score + 10)
                reason = f"{reason}; misma empresa"
            reason = f"{reason}; {'ID sugerido ya reportado' if cid in reported_ids else 'ID sugerido aún no reportado'}"
            if score >= 62:
                scored.append((score, reason, cid))
        scored.sort(key=lambda item: item[0], reverse=True)
        top = scored[:3]
        item: Dict[str, Any] = {
            "ID REPORTADO": bad_id,
            "FILAS_AFECTADAS": len(rows) or 1,
            "EMPRESA REPORTADA": first.get("EMPRESA", ""),
            "NOMBRE COMPLETO": first.get("NOMBRE COMPLETO", ""),
            "ARCHIVO": first.get("ARCHIVO", ""),
            "ACCION PROPUESTA": "Sin candidato confiable. Revisar manualmente contra curva y reportabilidad.",
        }
        for idx in range(1, 4):
            item[f"ID SUGERIDO {idx}"] = ""
            item[f"CONFIANZA {idx}"] = ""
            item[f"MOTIVO {idx}"] = ""
            item[f"EMPRESA CURVA {idx}"] = ""
            item[f"CONTRATO CURVA {idx}"] = ""
            item[f"SPA/ CO {idx}"] = ""
        for idx, (score, reason, cid) in enumerate(top, start=1):
            plan = plan_by_id.get(cid, {})
            item[f"ID SUGERIDO {idx}"] = cid
            item[f"CONFIANZA {idx}"] = confidence_label(score)
            item[f"MOTIVO {idx}"] = reason
            item[f"EMPRESA CURVA {idx}"] = plan.get("EMPRESA", "")
            item[f"CONTRATO CURVA {idx}"] = plan.get("NUMERO DE CONTRATO", "")
            item[f"SPA/ CO {idx}"] = export_contact_value(plan)
        if top:
            item["ACCION PROPUESTA"] = f"Validar si {bad_id} corresponde a {top[0][2]}. Si corresponde, corregir el ID en la reportabilidad original y recargar el archivo."
        suggestions.append(item)
    return suggestions


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
    curve_co_mel_col = pick_column_from_headers(curve_headers, SYNONYMS["CO MEL"])
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
                    "CO MEL": clean_text(record.get(curve_co_mel_col)) if curve_co_mel_col else "",
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
            {"EMPRESA": row["EMPRESA"], "ids_planificados": set(), "dotacion_planificada": 0.0, "contactos_spa_co": set()},
        )
        bucket["ids_planificados"].add(row["ID"])
        contact = export_contact_value(row)
        if contact:
            bucket["contactos_spa_co"].add(contact)
        bucket["dotacion_planificada"] += float(row["DOTACION_PLANIFICADA"] or 0)

    missing_companies: List[Dict[str, Any]] = []
    for norm, bucket in company_bucket.items():
        if norm not in reported_companies:
            dot = bucket["dotacion_planificada"]
            missing_companies.append(
                {
                    "EMPRESA": bucket["EMPRESA"],
                    "SPA/ CO": " | ".join(sorted(bucket.get("contactos_spa_co", set()))),
                    "ids_planificados": len(bucket["ids_planificados"]),
                    "dotacion_planificada": int(dot) if float(dot).is_integer() else round(dot, 2),
                }
            )
    missing_companies.sort(key=lambda r: normalize_company(r["EMPRESA"]))

    id_repairs = build_id_repair_suggestions(transformed, planned_rows, reported_ids)

    summary = {
        "curve_sheet": curve_sheet,
        "plan_scope": plan_scope,
        "planned_ids": len(planned_ids),
        "reported_ids": len(reported_ids),
        "matched_ids": len(planned_ids.intersection(reported_ids)),
        "missing_ids": len(planned_ids - reported_ids),
        "unmatched_reported_ids": len(reported_ids - planned_ids),
        "repair_suggestions": len(id_repairs),
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


def export_contact_value(row: Dict[str, Any]) -> str:
    """Devuelve el contacto SPA/CO de la curva.

    En la BD se guarda como co_mel/CO MEL por compatibilidad con versiones anteriores,
    pero en el Excel de faltantes se muestra con el nombre operativo de la curva: SPA/ CO.
    """
    return (
        clean_text(row.get("SPA/ CO"))
        or clean_text(row.get("SPA/CO"))
        or clean_text(row.get("CO MEL"))
        or clean_text(row.get("co_mel"))
    )


def write_output(path: Path, transformed: List[Dict[str, Any]], missing_ids: List[Dict[str, Any]], missing_companies: List[Dict[str, Any]], summary: Dict[str, Any], id_repairs: Optional[List[Dict[str, Any]]] = None) -> None:
    missing_ids_export = []
    for row in missing_ids:
        enriched = dict(row)
        enriched["SPA/ CO"] = export_contact_value(row)
        missing_ids_export.append(enriched)

    missing_companies_export = []
    for row in missing_companies:
        enriched = dict(row)
        enriched["SPA/ CO"] = export_contact_value(row)
        missing_companies_export.append(enriched)

    sheets = [
        ("Formato_Final", transformed, TARGET_COLUMNS),
        ("IDs_Reportados_Sin_Match", id_repairs or [], REPAIR_COLUMNS),
        ("Empresas_Sin_Reportabilidad", missing_companies_export, ["EMPRESA", "SPA/ CO", "ids_planificados", "dotacion_planificada"]),
        ("IDs_Planificados_No_Reportados", missing_ids_export, ["ID", "EMPRESA", "NUMERO DE CONTRATO", "SPA/ CO", "DOTACION_PLANIFICADA"]),
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


# -----------------------------
# Persistencia acumulativa en PostgreSQL
# -----------------------------
DATABASE_URL = os.environ.get("DATABASE_URL") or os.environ.get("POSTGRES_URL") or os.environ.get("POSTGRESQL_URL")
_DB_INITIALIZED = False
_DB_LAST_ERROR: Optional[str] = None


def db_enabled() -> bool:
    return bool(DATABASE_URL and pgdb is not None)


def db_status() -> Dict[str, Any]:
    return {
        "enabled": bool(DATABASE_URL),
        "driver_loaded": pgdb is not None,
        "initialized": _DB_INITIALIZED,
        "last_error": _DB_LAST_ERROR,
    }


def get_db_connection():
    if not DATABASE_URL:
        raise RuntimeError("DATABASE_URL no está configurada. Conecta la PostgreSQL de Render al servicio web.")
    if pgdb is None:
        raise RuntimeError("No está instalado el driver pg8000. Revisa requirements.txt.")

    parsed = urlparse(DATABASE_URL)
    query = parse_qs(parsed.query)
    sslmode = (query.get("sslmode", [os.environ.get("PGSSLMODE", "")])[0] or "").lower()
    db_ssl = (os.environ.get("DB_SSL", "") or "").lower()
    use_ssl = sslmode in {"require", "verify-ca", "verify-full"} or db_ssl in {"1", "true", "yes", "require"}

    ssl_context = None
    if use_ssl:
        ssl_context = ssl.create_default_context()
        if sslmode == "require":
            ssl_context.check_hostname = False
            ssl_context.verify_mode = ssl.CERT_NONE

    return pgdb.connect(
        host=parsed.hostname,
        port=parsed.port or 5432,
        database=(parsed.path or "/").lstrip("/"),
        user=unquote(parsed.username or ""),
        password=unquote(parsed.password or ""),
        ssl_context=ssl_context,
        timeout=int(os.environ.get("DB_TIMEOUT", "20")),
    )


def execute_db(sql: str, params: Optional[Sequence[Any]] = None, fetch: bool = False) -> Optional[List[Tuple[Any, ...]]]:
    conn = get_db_connection()
    try:
        cur = conn.cursor()
        cur.execute(sql, params or ())
        rows = cur.fetchall() if fetch else None
        conn.commit()
        cur.close()
        return rows
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def init_db() -> bool:
    """Crea el modelo acumulativo. Mantiene tablas históricas anteriores si existían."""
    global _DB_INITIALIZED, _DB_LAST_ERROR
    if _DB_INITIALIZED:
        return True
    if not DATABASE_URL:
        _DB_LAST_ERROR = "DATABASE_URL no configurada"
        return False
    if pgdb is None:
        _DB_LAST_ERROR = "pg8000 no instalado"
        return False
    try:
        conn = get_db_connection()
        cur = conn.cursor()
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS reportabilidad_semanas (
                semana_id TEXT PRIMARY KEY,
                created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
                updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
                nombre_semana TEXT NOT NULL,
                curve_filename TEXT NOT NULL,
                curve_sheet TEXT,
                plan_scope TEXT,
                plan_columns_used INTEGER DEFAULT 0,
                plan_summary_json TEXT NOT NULL DEFAULT '{}'
            );
            """
        )
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS reportabilidad_planificacion (
                id BIGSERIAL PRIMARY KEY,
                semana_id TEXT NOT NULL REFERENCES reportabilidad_semanas(semana_id) ON DELETE CASCADE,
                id_solicitud TEXT NOT NULL,
                empresa TEXT,
                empresa_norm TEXT,
                numero_contrato TEXT,
                co_mel TEXT,
                dotacion_planificada NUMERIC,
                dotacion_dias_sumada NUMERIC,
                UNIQUE (semana_id, id_solicitud)
            );
            """
        )
        cur.execute("ALTER TABLE reportabilidad_planificacion ADD COLUMN IF NOT EXISTS co_mel TEXT;")
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS reportabilidad_archivos (
                archivo_id TEXT PRIMARY KEY,
                semana_id TEXT NOT NULL REFERENCES reportabilidad_semanas(semana_id) ON DELETE CASCADE,
                created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
                filename TEXT NOT NULL,
                rows_imported INTEGER NOT NULL DEFAULT 0,
                ids_reported INTEGER NOT NULL DEFAULT 0,
                company_display TEXT,
                company_norm TEXT
            );
            """
        )
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS reportabilidad_formato_v2 (
                id BIGSERIAL PRIMARY KEY,
                semana_id TEXT NOT NULL REFERENCES reportabilidad_semanas(semana_id) ON DELETE CASCADE,
                archivo_id TEXT NOT NULL REFERENCES reportabilidad_archivos(archivo_id) ON DELETE CASCADE,
                fila INTEGER NOT NULL,
                id_solicitud TEXT,
                modulo TEXT,
                rut TEXT,
                nombre_completo TEXT,
                empresa TEXT,
                numero_contrato TEXT,
                gerencia TEXT,
                sistema_turno TEXT,
                co_mel TEXT,
                genero TEXT,
                nombre_turno TEXT
            );
            """
        )
        cur.execute("CREATE INDEX IF NOT EXISTS idx_rep_sem_updated ON reportabilidad_semanas(updated_at DESC);")
        cur.execute("CREATE INDEX IF NOT EXISTS idx_rep_plan_sem ON reportabilidad_planificacion(semana_id, id_solicitud);")
        cur.execute("CREATE INDEX IF NOT EXISTS idx_rep_arch_sem ON reportabilidad_archivos(semana_id, created_at DESC);")
        cur.execute("CREATE INDEX IF NOT EXISTS idx_rep_fmt_sem ON reportabilidad_formato_v2(semana_id, id_solicitud);")
        cur.execute("CREATE INDEX IF NOT EXISTS idx_rep_fmt_arch ON reportabilidad_formato_v2(archivo_id);")
        conn.commit()
        cur.close()
        conn.close()
        _DB_INITIALIZED = True
        _DB_LAST_ERROR = None
        return True
    except Exception as exc:
        _DB_LAST_ERROR = str(exc)
        app.logger.exception("No se pudo inicializar PostgreSQL")
        return False


@app.before_request
def before_request_init_db():
    if DATABASE_URL and not _DB_INITIALIZED:
        init_db()


def parse_summary(value: Any) -> Dict[str, Any]:
    if isinstance(value, dict):
        return value
    try:
        return json.loads(value or "{}")
    except Exception:
        return {}


def json_dumps(value: Any) -> str:
    return json.dumps(value, ensure_ascii=False)


# -----------------------------
# Lectura lógica de curva y reportabilidades
# -----------------------------
def parse_curve_planning(curve_path: Path) -> Tuple[List[Dict[str, Any]], Dict[str, Any]]:
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
    curve_co_mel_col = pick_column_from_headers(curve_headers, SYNONYMS["CO MEL"])
    if not curve_id_col or not curve_company_col:
        raise ValueError("No se pudo detectar ID de la solicitud y Empresa en la hoja Fcst_Autorizado VCA.")

    plan_keys, plan_scope = detect_plan_keys(curve_path, curve_sheet, curve_metas, curve_header_row)
    if not plan_keys:
        raise ValueError("No se pudieron detectar columnas de dotación planificada en la curva.")

    planned_by_id: Dict[str, Dict[str, Any]] = {}
    for record in curve_records:
        row_id = normalize_id(record.get(curve_id_col))
        company = clean_text(record.get(curve_company_col))
        planned_max = max_planned_value(record, plan_keys)
        if not row_id or planned_max <= 0:
            continue
        # Si el mismo ID aparece más de una vez en la curva, conserva el mayor valor planificado.
        current = planned_by_id.get(row_id)
        candidate = {
            "ID": row_id,
            "EMPRESA": company,
            "EMPRESA_NORM": normalize_company(company),
            "NUMERO DE CONTRATO": clean_text(record.get(curve_contract_col)) if curve_contract_col else "",
            "CO MEL": clean_text(record.get(curve_co_mel_col)) if curve_co_mel_col else "",
            "DOTACION_PLANIFICADA": int(planned_max) if float(planned_max).is_integer() else planned_max,
            "DOTACION_DIAS_SUMADA": sum_planned_days(record, plan_keys),
        }
        if current is None or float(candidate["DOTACION_PLANIFICADA"] or 0) > float(current["DOTACION_PLANIFICADA"] or 0):
            planned_by_id[row_id] = candidate

    planned_rows = list(planned_by_id.values())
    planned_companies = {r["EMPRESA_NORM"] for r in planned_rows if r.get("EMPRESA_NORM")}
    summary = {
        "curve_sheet": curve_sheet,
        "plan_scope": plan_scope,
        "planned_ids": len(planned_rows),
        "planned_companies": len(planned_companies),
        "plan_columns_used": len(plan_keys),
    }
    return planned_rows, summary


def parse_reportability_file(report_path: Path) -> Tuple[List[Dict[str, str]], Dict[str, Any]]:
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

    rows: List[Dict[str, str]] = []
    company_counter: Dict[str, Dict[str, Any]] = {}
    ids: set[str] = set()
    for record in report_records:
        row_id = normalize_id(record.get(report_id_col))
        if not row_id:
            continue
        out = row_to_output(record, report_headers)
        if not out["ID"]:
            continue
        rows.append(out)
        ids.add(out["ID"])
        company_display = clean_text(record.get(report_emp_col)) if report_emp_col else out.get("EMPRESA", "")
        company_norm = normalize_company(company_display)
        if company_norm:
            bucket = company_counter.setdefault(company_norm, {"display": company_display, "count": 0})
            bucket["count"] += 1

    company_display = ""
    company_norm = ""
    if company_counter:
        company_norm, bucket = sorted(company_counter.items(), key=lambda item: item[1]["count"], reverse=True)[0]
        company_display = bucket["display"]

    stats = {
        "sheet": report_sheet,
        "rows_imported": len(rows),
        "ids_reported": len(ids),
        "company_display": company_display,
        "company_norm": company_norm,
    }
    return rows, stats


# -----------------------------
# Escritura y lectura de semana acumulativa
# -----------------------------
def create_week_in_db(semana_id: str, nombre_semana: str, curve_filename: str, planned_rows: List[Dict[str, Any]], plan_summary: Dict[str, Any]) -> None:
    if not init_db():
        raise RuntimeError(_DB_LAST_ERROR or "No se pudo inicializar PostgreSQL.")
    conn = get_db_connection()
    try:
        cur = conn.cursor()
        cur.execute(
            """
            INSERT INTO reportabilidad_semanas
            (semana_id, nombre_semana, curve_filename, curve_sheet, plan_scope, plan_columns_used, plan_summary_json)
            VALUES (%s, %s, %s, %s, %s, %s, %s)
            """,
            (
                semana_id,
                nombre_semana,
                curve_filename,
                plan_summary.get("curve_sheet", ""),
                plan_summary.get("plan_scope", ""),
                int(plan_summary.get("plan_columns_used") or 0),
                json_dumps(plan_summary),
            ),
        )
        insert_planning_rows(cur, semana_id, planned_rows)
        conn.commit()
        cur.close()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def insert_planning_rows(cur, semana_id: str, planned_rows: List[Dict[str, Any]]) -> None:
    for row in planned_rows:
        cur.execute(
            """
            INSERT INTO reportabilidad_planificacion
            (semana_id, id_solicitud, empresa, empresa_norm, numero_contrato, co_mel, dotacion_planificada, dotacion_dias_sumada)
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s)
            ON CONFLICT (semana_id, id_solicitud)
            DO UPDATE SET
                empresa = EXCLUDED.empresa,
                empresa_norm = EXCLUDED.empresa_norm,
                numero_contrato = EXCLUDED.numero_contrato,
                co_mel = EXCLUDED.co_mel,
                dotacion_planificada = EXCLUDED.dotacion_planificada,
                dotacion_dias_sumada = EXCLUDED.dotacion_dias_sumada
            """,
            (
                semana_id,
                row.get("ID", ""),
                row.get("EMPRESA", ""),
                row.get("EMPRESA_NORM", ""),
                row.get("NUMERO DE CONTRATO", ""),
                row.get("CO MEL", ""),
                row.get("DOTACION_PLANIFICADA", 0),
                row.get("DOTACION_DIAS_SUMADA", 0),
            ),
        )


def replace_week_curve(semana_id: str, curve_filename: str, planned_rows: List[Dict[str, Any]], plan_summary: Dict[str, Any]) -> None:
    if not init_db():
        raise RuntimeError(_DB_LAST_ERROR or "No se pudo inicializar PostgreSQL.")
    conn = get_db_connection()
    try:
        cur = conn.cursor()
        cur.execute("DELETE FROM reportabilidad_planificacion WHERE semana_id = %s", (semana_id,))
        cur.execute(
            """
            UPDATE reportabilidad_semanas
            SET updated_at = NOW(), curve_filename = %s, curve_sheet = %s, plan_scope = %s,
                plan_columns_used = %s, plan_summary_json = %s
            WHERE semana_id = %s
            """,
            (
                curve_filename,
                plan_summary.get("curve_sheet", ""),
                plan_summary.get("plan_scope", ""),
                int(plan_summary.get("plan_columns_used") or 0),
                json_dumps(plan_summary),
                semana_id,
            ),
        )
        insert_planning_rows(cur, semana_id, planned_rows)
        conn.commit()
        cur.close()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def save_reportability_upload(semana_id: str, filename: str, rows: List[Dict[str, str]], stats: Dict[str, Any]) -> str:
    if not init_db():
        raise RuntimeError(_DB_LAST_ERROR or "No se pudo inicializar PostgreSQL.")
    archivo_id = uuid.uuid4().hex[:12]
    conn = get_db_connection()
    try:
        cur = conn.cursor()
        cur.execute(
            """
            INSERT INTO reportabilidad_archivos
            (archivo_id, semana_id, filename, rows_imported, ids_reported, company_display, company_norm)
            VALUES (%s, %s, %s, %s, %s, %s, %s)
            """,
            (
                archivo_id,
                semana_id,
                filename,
                int(stats.get("rows_imported") or len(rows)),
                int(stats.get("ids_reported") or 0),
                stats.get("company_display", ""),
                stats.get("company_norm", ""),
            ),
        )
        for fila, row in enumerate(rows, start=1):
            cur.execute(
                """
                INSERT INTO reportabilidad_formato_v2
                (semana_id, archivo_id, fila, id_solicitud, modulo, rut, nombre_completo, empresa, numero_contrato, gerencia, sistema_turno, co_mel, genero, nombre_turno)
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                """,
                (
                    semana_id,
                    archivo_id,
                    fila,
                    row.get("ID", ""),
                    row.get("MODULO", ""),
                    row.get("RUT (CON GUION)", ""),
                    row.get("NOMBRE COMPLETO", ""),
                    row.get("EMPRESA", ""),
                    row.get("NUMERO DE CONTRATO", ""),
                    row.get("GERENCIA", ""),
                    row.get("SISTEMA DE TURNO", ""),
                    row.get("CO MEL", ""),
                    row.get("GENERO", ""),
                    row.get("NOMBRE DE TURNO", ""),
                ),
            )
        cur.execute("UPDATE reportabilidad_semanas SET updated_at = NOW() WHERE semana_id = %s", (semana_id,))
        conn.commit()
        cur.close()
        return archivo_id
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def get_week_base(semana_id: str) -> Optional[Dict[str, Any]]:
    if not init_db():
        return None
    rows = execute_db(
        """
        SELECT semana_id, created_at, updated_at, nombre_semana, curve_filename, curve_sheet, plan_scope, plan_columns_used, plan_summary_json
        FROM reportabilidad_semanas
        WHERE semana_id = %s
        """,
        (semana_id,),
        fetch=True,
    ) or []
    if not rows:
        return None
    row = rows[0]
    return {
        "semana_id": row[0],
        "created_at": row[1].strftime("%Y-%m-%d %H:%M:%S") if hasattr(row[1], "strftime") else clean_text(row[1]),
        "updated_at": row[2].strftime("%Y-%m-%d %H:%M:%S") if hasattr(row[2], "strftime") else clean_text(row[2]),
        "nombre_semana": row[3] or "",
        "curve_filename": row[4] or "",
        "curve_sheet": row[5] or "",
        "plan_scope": row[6] or "",
        "plan_columns_used": row[7] or 0,
        "plan_summary": parse_summary(row[8]),
    }


def load_week_state(semana_id: str, preview_only: bool = True) -> Optional[Dict[str, Any]]:
    base = get_week_base(semana_id)
    if not base:
        return None
    conn = get_db_connection()
    try:
        cur = conn.cursor()
        cur.execute(
            """
            SELECT id_solicitud, empresa, empresa_norm, numero_contrato, co_mel, dotacion_planificada, dotacion_dias_sumada
            FROM reportabilidad_planificacion
            WHERE semana_id = %s
            ORDER BY empresa, id_solicitud
            """,
            (semana_id,),
        )
        plan_rows = [
            {
                "ID": r[0] or "",
                "EMPRESA": r[1] or "",
                "EMPRESA_NORM": r[2] or "",
                "NUMERO DE CONTRATO": r[3] or "",
                "CO MEL": r[4] or "",
                "DOTACION_PLANIFICADA": r[5] or 0,
                "DOTACION_DIAS_SUMADA": r[6] or 0,
            }
            for r in cur.fetchall()
        ]
        cur.execute(
            """
            SELECT archivo_id, created_at, filename, rows_imported, ids_reported, company_display
            FROM reportabilidad_archivos
            WHERE semana_id = %s
            ORDER BY created_at DESC
            """,
            (semana_id,),
        )
        archivos = [
            {
                "archivo_id": r[0],
                "created_at": r[1].strftime("%Y-%m-%d %H:%M:%S") if hasattr(r[1], "strftime") else clean_text(r[1]),
                "filename": r[2] or "",
                "rows_imported": r[3] or 0,
                "ids_reported": r[4] or 0,
                "company_display": r[5] or "",
            }
            for r in cur.fetchall()
        ]
        cur.execute(
            """
            SELECT f.id, f.archivo_id, a.filename, f.id_solicitud, f.modulo, f.rut, f.nombre_completo, f.empresa, f.numero_contrato,
                   f.gerencia, f.sistema_turno, f.co_mel, f.genero, f.nombre_turno, a.created_at
            FROM reportabilidad_formato_v2 f
            JOIN reportabilidad_archivos a ON a.archivo_id = f.archivo_id
            WHERE f.semana_id = %s
            ORDER BY a.created_at ASC, f.id ASC
            """,
            (semana_id,),
        )
        report_rows_raw = cur.fetchall()
        cur.close()
    finally:
        conn.close()

    # Exportación completa: se conservan TODAS las personas/filas cargadas.
    # El ID se usa solo para calcular cobertura contra la curva, no para eliminar filas del formato final.
    transformed_all: List[Dict[str, Any]] = []
    reported_ids: set[str] = set()
    for r in report_rows_raw:
        row_id = normalize_id(r[3])
        if not row_id:
            continue
        reported_ids.add(row_id)
        transformed_all.append({
            "ARCHIVO_ID": r[1] or "",
            "ARCHIVO": r[2] or "",
            "ID": row_id,
            "MODULO": r[4] or "",
            "RUT (CON GUION)": r[5] or "",
            "NOMBRE COMPLETO": r[6] or "",
            "EMPRESA": r[7] or "",
            "NUMERO DE CONTRATO": r[8] or "",
            "GERENCIA": r[9] or "",
            "SISTEMA DE TURNO": r[10] or "",
            "CO MEL": r[11] or "",
            "GENERO": r[12] or "",
            "NOMBRE DE TURNO": r[13] or "",
        })

    planned_ids = {r["ID"] for r in plan_rows if r.get("ID")}
    plan_by_id = {r["ID"]: r for r in plan_rows if r.get("ID")}

    reported_companies: set[str] = set()
    for row in transformed_all:
        norm = normalize_company(row.get("EMPRESA"))
        if norm:
            reported_companies.add(norm)
    # Si un ID planificado fue reportado, la empresa de la curva también cuenta como reportada.
    for row_id in reported_ids:
        plan_row = plan_by_id.get(row_id)
        if plan_row and plan_row.get("EMPRESA_NORM"):
            reported_companies.add(plan_row["EMPRESA_NORM"])

    missing_ids = sorted(
        [r for r in plan_rows if r.get("ID") not in reported_ids],
        key=lambda r: (normalize_company(r.get("EMPRESA")), r.get("ID", "")),
    )
    missing_ids_out = [
        {
            "ID": r.get("ID", ""),
            "EMPRESA": r.get("EMPRESA", ""),
            "NUMERO DE CONTRATO": r.get("NUMERO DE CONTRATO", ""),
            "CO MEL": r.get("CO MEL", ""),
            "SPA/ CO": export_contact_value(r),
            "DOTACION_PLANIFICADA": clean_text(r.get("DOTACION_PLANIFICADA", "")),
        }
        for r in missing_ids
    ]

    company_bucket: Dict[str, Dict[str, Any]] = {}
    for row in plan_rows:
        norm = row.get("EMPRESA_NORM") or normalize_company(row.get("EMPRESA"))
        if not norm:
            continue
        bucket = company_bucket.setdefault(norm, {"EMPRESA": row.get("EMPRESA", ""), "ids_planificados": set(), "dotacion_planificada": 0.0, "contactos_spa_co": set()})
        bucket["ids_planificados"].add(row.get("ID"))
        contact = export_contact_value(row)
        if contact:
            bucket["contactos_spa_co"].add(contact)
        try:
            bucket["dotacion_planificada"] += float(row.get("DOTACION_PLANIFICADA") or 0)
        except Exception:
            pass

    missing_companies = []
    for norm, bucket in company_bucket.items():
        if norm not in reported_companies:
            dot = bucket["dotacion_planificada"]
            missing_companies.append(
                {
                    "EMPRESA": bucket["EMPRESA"],
                    "SPA/ CO": " | ".join(sorted(bucket.get("contactos_spa_co", set()))),
                    "ids_planificados": len(bucket["ids_planificados"]),
                    "dotacion_planificada": int(dot) if float(dot).is_integer() else round(dot, 2),
                }
            )
    missing_companies.sort(key=lambda r: normalize_company(r["EMPRESA"]))

    id_repairs = build_id_repair_suggestions(transformed_all, plan_rows, reported_ids)

    summary = {
        "curve_sheet": base["curve_sheet"],
        "plan_scope": base["plan_scope"],
        "planned_ids": len(planned_ids),
        "reported_ids": len(reported_ids),
        "matched_ids": len(planned_ids.intersection(reported_ids)),
        "missing_ids": len(planned_ids - reported_ids),
        "unmatched_reported_ids": len(reported_ids - planned_ids),
        "repair_suggestions": len(id_repairs),
        "planned_companies": len(company_bucket),
        "reported_companies": len(reported_companies),
        "missing_companies": len(missing_companies),
        "transformed_rows": len(transformed_all),
        "plan_columns_used": base["plan_columns_used"],
        "report_files": len(archivos),
        "curve_filename": base["curve_filename"],
        "nombre_semana": base["nombre_semana"],
        "updated_at": base["updated_at"],
    }
    return {
        **base,
        "job": semana_id,
        "summary": summary,
        "archivos": archivos,
        "preview": transformed_all[:20] if preview_only else transformed_all,
        "transformed_all": transformed_all,
        "missing_ids": missing_ids_out[:80] if preview_only else missing_ids_out,
        "missing_companies": missing_companies[:80] if preview_only else missing_companies,
        "id_repairs": id_repairs[:80] if preview_only else id_repairs,
    }


def list_weeks() -> List[Dict[str, Any]]:
    if not init_db():
        return []
    rows = execute_db(
        """
        SELECT semana_id
        FROM reportabilidad_semanas
        ORDER BY updated_at DESC
        LIMIT 150
        """,
        fetch=True,
    ) or []
    weeks = []
    for row in rows:
        state = load_week_state(row[0], preview_only=True)
        if state:
            weeks.append(state)
    return weeks


def delete_week_from_db(semana_id: str) -> None:
    if not init_db():
        raise RuntimeError(_DB_LAST_ERROR or "No se pudo inicializar PostgreSQL.")
    execute_db("DELETE FROM reportabilidad_semanas WHERE semana_id = %s", (semana_id,))


def delete_report_file_from_db(semana_id: str, archivo_id: str) -> None:
    if not init_db():
        raise RuntimeError(_DB_LAST_ERROR or "No se pudo inicializar PostgreSQL.")
    execute_db("DELETE FROM reportabilidad_archivos WHERE semana_id = %s AND archivo_id = %s", (semana_id, archivo_id))
    execute_db("UPDATE reportabilidad_semanas SET updated_at = NOW() WHERE semana_id = %s", (semana_id,))


def remove_local_artifacts(job: str) -> None:
    # En Render el disco local es efímero; de todos modos limpiamos cargas temporales
    # porque la información persistente queda guardada en PostgreSQL.
    for folder in (OUTPUT_DIR, UPLOAD_DIR):
        for path in folder.glob(f"*{job}*"):
            try:
                path.unlink()
            except Exception:
                pass


# -----------------------------
# Rutas Flask
# -----------------------------
@app.route("/", methods=["GET"])
def index():
    semanas = list_weeks() if DATABASE_URL else []
    return render_template("index.html", db_status=db_status(), semanas=semanas)


@app.route("/healthz", methods=["GET"])
def healthz():
    return {"status": "ok", "db": db_status()}


@app.route("/historial", methods=["GET"])
def historial():
    semanas = list_weeks() if DATABASE_URL else []
    return render_template("historial.html", semanas=semanas, db_status=db_status())


@app.route("/crear_semana", methods=["POST"])
def crear_semana():
    try:
        if not DATABASE_URL:
            flash("Debes configurar DATABASE_URL para guardar semanas acumulativas en PostgreSQL.", "error")
            return redirect(url_for("index"))
        curva = request.files.get("curva")
        reports = [f for f in request.files.getlist("reportabilidad") if f and f.filename]
        nombre_semana = clean_text(request.form.get("nombre_semana", ""))
        if not curva or not curva.filename:
            flash("Debes subir la curva de poblamiento para crear la semana.", "error")
            return redirect(url_for("index"))

        semana_id = uuid.uuid4().hex[:10]
        curve_original_name = secure_filename(curva.filename) or curva.filename
        curve_path = safe_filename("curva", curva.filename, semana_id)
        curva.save(curve_path)
        planned_rows, plan_summary = parse_curve_planning(curve_path)
        if not nombre_semana:
            nombre_semana = f"{plan_summary.get('plan_scope', 'Semana')} - {curve_original_name}"
        create_week_in_db(semana_id, nombre_semana, curve_original_name, planned_rows, plan_summary)

        loaded_reports = 0
        for idx, file_storage in enumerate(reports, start=1):
            original_name = secure_filename(file_storage.filename) or file_storage.filename
            report_path = safe_filename("report", file_storage.filename, semana_id, idx)
            file_storage.save(report_path)
            rows, stats = parse_reportability_file(report_path)
            save_reportability_upload(semana_id, original_name, rows, stats)
            loaded_reports += 1

        remove_local_artifacts(semana_id)
        if loaded_reports:
            flash(f"Semana creada y {loaded_reports} reportabilidad(es) agregada(s). Podrás seguir alimentando esta misma semana después.", "ok")
        else:
            flash("Semana creada con su curva. Ahora puedes ir cargando reportabilidades día a día.", "ok")
        return redirect(url_for("ver_semana", semana_id=semana_id))
    except Exception as exc:
        app.logger.exception("Error al crear semana")
        flash(f"Error al crear semana: {exc}", "error")
        return redirect(url_for("index"))


@app.route("/semana/<semana_id>", methods=["GET"])
def ver_semana(semana_id: str):
    meta = load_week_state(semana_id, preview_only=True) if DATABASE_URL else None
    if not meta:
        flash("No se encontró la semana solicitada.", "error")
        return redirect(url_for("historial"))
    return render_template("semana.html", meta=meta, target_columns=TARGET_COLUMNS, repair_columns=REPAIR_COLUMNS)


@app.route("/semana/<semana_id>/agregar", methods=["POST"])
def agregar_reportabilidad(semana_id: str):
    try:
        reports = [f for f in request.files.getlist("reportabilidad") if f and f.filename]
        if not reports:
            flash("Debes seleccionar al menos una reportabilidad para agregar.", "error")
            return redirect(url_for("ver_semana", semana_id=semana_id))
        if not get_week_base(semana_id):
            flash("La semana seleccionada no existe.", "error")
            return redirect(url_for("historial"))
        loaded_reports = 0
        for idx, file_storage in enumerate(reports, start=1):
            original_name = secure_filename(file_storage.filename) or file_storage.filename
            report_path = safe_filename("report", file_storage.filename, semana_id, idx)
            file_storage.save(report_path)
            rows, stats = parse_reportability_file(report_path)
            save_reportability_upload(semana_id, original_name, rows, stats)
            loaded_reports += 1
        remove_local_artifacts(semana_id)
        flash(f"Se agregaron {loaded_reports} reportabilidad(es) a la semana. El cruce quedó actualizado automáticamente.", "ok")
        return redirect(url_for("ver_semana", semana_id=semana_id))
    except Exception as exc:
        app.logger.exception("Error al agregar reportabilidad")
        flash(f"Error al agregar reportabilidad: {exc}", "error")
        return redirect(url_for("ver_semana", semana_id=semana_id))


@app.route("/semana/<semana_id>/reemplazar_curva", methods=["POST"])
def reemplazar_curva(semana_id: str):
    try:
        curva = request.files.get("curva")
        if not curva or not curva.filename:
            flash("Debes seleccionar una nueva curva.", "error")
            return redirect(url_for("ver_semana", semana_id=semana_id))
        if not get_week_base(semana_id):
            flash("La semana seleccionada no existe.", "error")
            return redirect(url_for("historial"))
        curve_original_name = secure_filename(curva.filename) or curva.filename
        curve_path = safe_filename("curva", curva.filename, semana_id)
        curva.save(curve_path)
        planned_rows, plan_summary = parse_curve_planning(curve_path)
        replace_week_curve(semana_id, curve_original_name, planned_rows, plan_summary)
        remove_local_artifacts(semana_id)
        flash("Curva reemplazada correctamente. Las reportabilidades ya cargadas se mantienen y el cruce fue recalculado.", "ok")
        return redirect(url_for("ver_semana", semana_id=semana_id))
    except Exception as exc:
        app.logger.exception("Error al reemplazar curva")
        flash(f"Error al reemplazar curva: {exc}", "error")
        return redirect(url_for("ver_semana", semana_id=semana_id))


@app.route("/semana/<semana_id>/eliminar_archivo/<archivo_id>", methods=["POST"])
def eliminar_archivo(semana_id: str, archivo_id: str):
    try:
        delete_report_file_from_db(semana_id, archivo_id)
        remove_local_artifacts(semana_id)
        flash("Reportabilidad eliminada de esta semana. El cruce quedó actualizado con los archivos restantes.", "ok")
    except Exception as exc:
        app.logger.exception("Error al eliminar archivo")
        flash(f"Error al eliminar reportabilidad: {exc}", "error")
    return redirect(url_for("ver_semana", semana_id=semana_id))


@app.route("/descargar/<semana_id>")
def descargar(semana_id: str):
    meta = load_week_state(semana_id, preview_only=False) if DATABASE_URL else None
    if not meta:
        flash("No se encontró la semana solicitada.", "error")
        return redirect(url_for("historial"))
    output_path = OUTPUT_DIR / f"reportabilidad_vca_{semana_id}.xlsx"
    summary = dict(meta["summary"])
    summary["archivos_reportabilidad"] = ", ".join([a["filename"] for a in meta.get("archivos", [])])
    write_output(output_path, meta["transformed_all"], meta["missing_ids"], meta["missing_companies"], summary, meta.get("id_repairs", []))
    return send_file(output_path, as_attachment=True, download_name=f"Reportabilidad_VCA_{meta['nombre_semana']}.xlsx")


@app.route("/semana/<semana_id>/eliminar", methods=["POST"])
def eliminar_semana(semana_id: str):
    try:
        delete_week_from_db(semana_id)
        remove_local_artifacts(semana_id)
        flash("Semana eliminada correctamente, incluyendo curva, reportabilidades y datos transformados asociados.", "ok")
    except Exception as exc:
        app.logger.exception("Error al eliminar semana")
        flash(f"Error al eliminar semana: {exc}", "error")
    return redirect(url_for("historial"))


# Compatibilidad con enlaces de versiones anteriores
@app.route("/resultado/<job>")
def resultado(job: str):
    return redirect(url_for("ver_semana", semana_id=job))


@app.route("/eliminar/<job>", methods=["POST"])
def eliminar(job: str):
    return redirect(url_for("eliminar_semana", semana_id=job), code=307)


if __name__ == "__main__":
    app.run(debug=True)
