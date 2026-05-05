"""
db.py — Capa de acceso a datos (SQLite)
Facultad de Ingeniería · Universidad de La Sabana

Responsabilidades:
  - Crear y migrar el esquema de la base de datos
  - Importar datos desde un DataFrame (cargado desde Excel)
  - Exponer queries para la API REST
"""

import sqlite3
import pandas as pd
import duckdb
from pathlib import Path

DB_PATH = Path(__file__).parent.parent / "facultad.db"


# ---------------------------------------------------------------------------
# Esquema
# ---------------------------------------------------------------------------

SCHEMA = """
CREATE TABLE IF NOT EXISTS clases (
    id                      INTEGER PRIMARY KEY AUTOINCREMENT,
    ciclo_lectivo           TEXT    NOT NULL,
    nombre_profesor         TEXT    NOT NULL,
    documento_docente       TEXT,
    nombre_curso            TEXT    NOT NULL,
    componente              TEXT,
    componente_desc         TEXT,
    descripcion_materia     TEXT,
    hora_inicio             TEXT,
    hora_fin                TEXT,
    num_clase               TEXT,
    id_seccion_combinada    TEXT,
    fecha_inicio            TEXT,
    fecha_fin               TEXT,
    dia                     TEXT
);

CREATE INDEX IF NOT EXISTS idx_profesor    ON clases (nombre_profesor);
CREATE INDEX IF NOT EXISTS idx_ciclo       ON clases (ciclo_lectivo);
CREATE INDEX IF NOT EXISTS idx_curso       ON clases (nombre_curso);

CREATE TABLE IF NOT EXISTS semestres_cargados (
    ciclo_lectivo   TEXT PRIMARY KEY,
    filas_insertadas INTEGER,
    cargado_en      TEXT DEFAULT (datetime('now'))
);
"""


def get_conn() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db() -> None:
    """Crea las tablas si no existen."""
    with get_conn() as conn:
        conn.executescript(SCHEMA)
    print(f"Base de datos lista: {DB_PATH}")


# ---------------------------------------------------------------------------
# Importación desde Excel / DataFrame
# ---------------------------------------------------------------------------

COLUMNAS_REQUERIDAS = {
    "Ciclo Lectivo",
    "Nombre profesor",
    "Numero documento docente",
    "Nombre del curso",
    "Componente",
    "Componente Descripción",
    "Descripción Materia",
    "Hora Inicio",
    "Hora Final",
    "Nº Clase",
    "ID Sección Combinada",
    "F Inicial",
    "Fecha Final",
    "Día",
}


def validar_columnas(df: pd.DataFrame) -> list[str]:
    """Devuelve lista de columnas faltantes (vacía = OK)."""
    return sorted(COLUMNAS_REQUERIDAS - set(df.columns))


def cargar_dataframe(df: pd.DataFrame) -> dict:
    """
    Inserta filas del DataFrame en SQLite.
    - Omite semestres que ya existen (idempotente).
    - Devuelve resumen con semestres nuevos, omitidos y filas insertadas.
    """
    # Normalizar ID Sección Combinada
    df["ID Sección Combinada"] = (
        df["ID Sección Combinada"]
        .astype(str)
        .str.strip()
        .replace({"": None, "nan": None, "None": None})
    )

    semestres_en_excel = df["Ciclo Lectivo"].dropna().unique().tolist()

    with get_conn() as conn:
        ya_cargados = {
            r["ciclo_lectivo"]
            for r in conn.execute("SELECT ciclo_lectivo FROM semestres_cargados").fetchall()
        }

    semestres_nuevos = [s for s in semestres_en_excel if s not in ya_cargados]
    semestres_omitidos = [s for s in semestres_en_excel if s in ya_cargados]

    if not semestres_nuevos:
        return {
            "semestres_nuevos": [],
            "semestres_omitidos": semestres_omitidos,
            "filas_insertadas": 0,
            "mensaje": "Todos los semestres del archivo ya estaban cargados. No se insertó nada.",
        }

    df_nuevo = df[
        df["Ciclo Lectivo"].isin(semestres_nuevos) &
        df["Nombre profesor"].notna() &
        (df["Nombre profesor"].astype(str).str.strip() != "")
    ].copy()

    registros = [
        {
            "ciclo_lectivo":        row.get("Ciclo Lectivo"),
            "nombre_profesor":      row.get("Nombre profesor"),
            "documento_docente":    str(row.get("Numero documento docente", "") or ""),
            "nombre_curso":         row.get("Nombre del curso"),
            "componente":           row.get("Componente"),
            "componente_desc":      row.get("Componente Descripción"),
            "descripcion_materia":  row.get("Descripción Materia"),
            "hora_inicio":          row.get("Hora Inicio"),
            "hora_fin":             row.get("Hora Final"),
            "num_clase":            str(row.get("Nº Clase", "") or ""),
            "id_seccion_combinada": row.get("ID Sección Combinada"),
            "fecha_inicio":         str(row.get("F Inicial", "") or ""),
            "fecha_fin":            str(row.get("Fecha Final", "") or ""),
            "dia":                  str(row.get("Día", "") or "").strip(),
        }
        for _, row in df_nuevo.iterrows()
    ]

    with get_conn() as conn:
        conn.executemany(
            """INSERT INTO clases (
                ciclo_lectivo, nombre_profesor, documento_docente,
                nombre_curso, componente, componente_desc,
                descripcion_materia, hora_inicio, hora_fin,
                num_clase, id_seccion_combinada, fecha_inicio, fecha_fin, dia
            ) VALUES (
                :ciclo_lectivo, :nombre_profesor, :documento_docente,
                :nombre_curso, :componente, :componente_desc,
                :descripcion_materia, :hora_inicio, :hora_fin,
                :num_clase, :id_seccion_combinada, :fecha_inicio, :fecha_fin, :dia
            )""",
            registros,
        )
        for sem in semestres_nuevos:
            conn.execute(
                "INSERT OR REPLACE INTO semestres_cargados (ciclo_lectivo, filas_insertadas) VALUES (?, ?)",
                (sem, sum(1 for r in registros if r["ciclo_lectivo"] == sem)),
            )

    return {
        "semestres_nuevos": semestres_nuevos,
        "semestres_omitidos": semestres_omitidos,
        "filas_insertadas": len(registros),
        "mensaje": f"{len(semestres_nuevos)} semestre(s) cargado(s) con {len(registros)} filas.",
    }


# ---------------------------------------------------------------------------
# Queries para la API
# ---------------------------------------------------------------------------

QUERY_CONSOLIDADO = """
WITH horas_por_grupo AS (
    SELECT
        nombre_profesor,
        documento_docente,
        ciclo_lectivo,
        nombre_curso,
        descripcion_materia   AS departamento,
        componente_desc       AS componente,
        MIN(fecha_inicio)     AS fecha_inicio,
        MAX(fecha_fin)        AS fecha_fin,
        num_clase,
        COUNT(*)              AS horas_semana_grupo
    FROM df
    WHERE (id_seccion_combinada IS NULL OR id_seccion_combinada = '')
      {filtro_profesor}
      {filtro_semestre}
    GROUP BY
        nombre_profesor, documento_docente, ciclo_lectivo,
        nombre_curso, descripcion_materia, componente_desc, num_clase
),
consolidado AS (
    SELECT
        nombre_profesor,
        documento_docente,
        ciclo_lectivo         AS semestre,
        nombre_curso          AS asignatura,
        departamento,
        componente,
        MIN(fecha_inicio)     AS fecha_inicio,
        MAX(fecha_fin)        AS fecha_fin,
        SUM(horas_semana_grupo) * 16  AS sesiones
    FROM horas_por_grupo
    GROUP BY
        nombre_profesor, documento_docente, ciclo_lectivo,
        nombre_curso, departamento, componente
)
SELECT * FROM consolidado
ORDER BY nombre_profesor, semestre, asignatura
"""


def _cargar_df_desde_sqlite(semestres: list[str] | None = None) -> pd.DataFrame:
    """Lee clases desde SQLite a un DataFrame para procesarlo con DuckDB."""
    query = "SELECT * FROM clases"
    params = []
    if semestres:
        placeholders = ",".join("?" * len(semestres))
        query += f" WHERE ciclo_lectivo IN ({placeholders})"
        params = semestres
    with get_conn() as conn:
        return pd.read_sql_query(query, conn, params=params)


def consultar_consolidado(
    profesor: str | None = None,
    semestres: list[str] | None = None,
) -> list[dict]:
    """
    Devuelve el consolidado de horas aplicando las reglas de negocio.
    profesor  : fragmento de nombre (búsqueda LIKE, insensible a mayúsculas)
    semestres : lista de strings como ['PERIODO 2024-2', 'PERIODO 2025-1']
    """
    df = _cargar_df_desde_sqlite(semestres)
    if df.empty:
        return []

    filtro_profesor = ""
    if profesor:
        filtro_profesor = f"AND UPPER(nombre_profesor) LIKE UPPER('%{profesor}%')"

    filtro_semestre = ""
    if semestres:
        lista = ", ".join(f"'PERIODO {s}'" if not s.startswith("PERIODO") else f"'{s}'" for s in semestres)
        filtro_semestre = f"AND ciclo_lectivo IN ({lista})"

    query = QUERY_CONSOLIDADO.format(
        filtro_profesor=filtro_profesor,
        filtro_semestre=filtro_semestre,
    )

    result = duckdb.sql(query).df().fillna("—")
    return result.to_dict(orient="records")


def listar_profesores(busqueda: str | None = None) -> list[str]:
    query = "SELECT DISTINCT nombre_profesor FROM clases"
    params = []
    if busqueda:
        query += " WHERE UPPER(nombre_profesor) LIKE UPPER(?)"
        params = [f"%{busqueda}%"]
    query += " ORDER BY nombre_profesor"
    with get_conn() as conn:
        rows = conn.execute(query, params).fetchall()
    return [r["nombre_profesor"] for r in rows]


def listar_semestres() -> list[str]:
    with get_conn() as conn:
        rows = conn.execute(
            "SELECT ciclo_lectivo, filas_insertadas, cargado_en FROM semestres_cargados ORDER BY ciclo_lectivo"
        ).fetchall()
    return [dict(r) for r in rows]


def obtener_opciones(profesor: str | None = None, semestres: list[str] | None = None) -> dict:
    """Obtiene materias, departamentos y componentes únicos para los filtros."""
    query = "SELECT DISTINCT nombre_curso, descripcion_materia, componente_desc FROM clases WHERE 1=1"
    params = []

    if profesor:
        query += " AND UPPER(nombre_profesor) LIKE UPPER(?)"
        params.append(f"%{profesor}%")

    if semestres:
        placeholders = ",".join("?" * len(semestres))
        query += f" AND ciclo_lectivo IN ({placeholders})"
        params.extend(semestres)

    with get_conn() as conn:
        rows = conn.execute(query, params).fetchall()

    materias = sorted(list(set(r["nombre_curso"] for r in rows if r["nombre_curso"])))
    departamentos = sorted(list(set(r["descripcion_materia"] for r in rows if r["descripcion_materia"])))
    componentes = sorted(list(set(r["componente_desc"] for r in rows if r["componente_desc"])))

    return {
        "materias": materias,
        "departamentos": departamentos,
        "componentes": componentes
    }