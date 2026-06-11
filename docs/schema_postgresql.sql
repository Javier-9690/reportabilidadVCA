CREATE TABLE IF NOT EXISTS reportabilidad_procesos (
    job_id TEXT PRIMARY KEY,
    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    descripcion TEXT,
    curve_filename TEXT NOT NULL,
    report_filenames TEXT NOT NULL,
    curve_sheet TEXT,
    plan_scope TEXT,
    summary_json TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS reportabilidad_formato (
    id BIGSERIAL PRIMARY KEY,
    job_id TEXT NOT NULL REFERENCES reportabilidad_procesos(job_id) ON DELETE CASCADE,
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

CREATE TABLE IF NOT EXISTS reportabilidad_missing_ids (
    id BIGSERIAL PRIMARY KEY,
    job_id TEXT NOT NULL REFERENCES reportabilidad_procesos(job_id) ON DELETE CASCADE,
    id_solicitud TEXT,
    empresa TEXT,
    numero_contrato TEXT,
    dotacion_planificada TEXT
);

CREATE TABLE IF NOT EXISTS reportabilidad_missing_empresas (
    id BIGSERIAL PRIMARY KEY,
    job_id TEXT NOT NULL REFERENCES reportabilidad_procesos(job_id) ON DELETE CASCADE,
    empresa TEXT,
    ids_planificados INTEGER,
    dotacion_planificada TEXT
);

CREATE INDEX IF NOT EXISTS idx_reportabilidad_procesos_created ON reportabilidad_procesos(created_at DESC);
CREATE INDEX IF NOT EXISTS idx_reportabilidad_formato_job ON reportabilidad_formato(job_id, fila);
CREATE INDEX IF NOT EXISTS idx_reportabilidad_missing_ids_job ON reportabilidad_missing_ids(job_id);
CREATE INDEX IF NOT EXISTS idx_reportabilidad_missing_empresas_job ON reportabilidad_missing_empresas(job_id);
