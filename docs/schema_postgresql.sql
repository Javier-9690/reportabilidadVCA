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

ALTER TABLE reportabilidad_planificacion ADD COLUMN IF NOT EXISTS co_mel TEXT;

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

CREATE INDEX IF NOT EXISTS idx_rep_sem_updated ON reportabilidad_semanas(updated_at DESC);
CREATE INDEX IF NOT EXISTS idx_rep_plan_sem ON reportabilidad_planificacion(semana_id, id_solicitud);
CREATE INDEX IF NOT EXISTS idx_rep_arch_sem ON reportabilidad_archivos(semana_id, created_at DESC);
CREATE INDEX IF NOT EXISTS idx_rep_fmt_sem ON reportabilidad_formato_v2(semana_id, id_solicitud);
CREATE INDEX IF NOT EXISTS idx_rep_fmt_arch ON reportabilidad_formato_v2(archivo_id);
