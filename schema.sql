
CREATE TABLE IF NOT EXISTS evaluations (
    submission_id UUID PRIMARY KEY,
    submitted_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    student_name TEXT,
    registration TEXT,
    section TEXT,
    xlsx_filename TEXT,
    report_filename TEXT,
    code_filename TEXT,
    valid_days INTEGER,
    dci DOUBLE PRECISION,
    raw_score DOUBLE PRECISION,
    scaled_score DOUBLE PRECISION,
    indices JSONB,
    truthfulness JSONB,
    relationships JSONB,
    rubric_scores JSONB,
    suggestions JSONB,
    issues JSONB
);

CREATE INDEX IF NOT EXISTS idx_evaluations_registration
ON evaluations (registration);

CREATE INDEX IF NOT EXISTS idx_evaluations_submitted_at
ON evaluations (submitted_at);
