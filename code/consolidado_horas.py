"""
Módulo B — Consolidado de horas por profesor
Facultad de Ingeniería · Universidad de La Sabana

Uso:
    python consolidado_horas.py --excel PREGRADO_CONSOLIDADO_COPIA.xlsx
    python consolidado_horas.py --excel PREGRADO_CONSOLIDADO_COPIA.xlsx --profesor "BRAVO BUITRAGO"
    python consolidado_horas.py --excel PREGRADO_CONSOLIDADO_COPIA.xlsx --profesor "BRAVO" --semestre "2024-2" "2025-1"
    python consolidado_horas.py --excel PREGRADO_CONSOLIDADO_COPIA.xlsx --listar-profesores
"""

import argparse
import sys
import pandas as pd
import duckdb


QUERY_CONSOLIDADO = """
WITH horas_por_grupo AS (
    -- Paso 1: contar horas semanales por grupo (Nº Clase) y materia
    -- Cada fila = 1 bloque de 1 hora (Hora Inicio → Hora Final)
    -- Se excluyen filas de sección combinada para no duplicar
    SELECT
        "Nombre profesor"                       AS profesor,
        "Numero documento docente"              AS documento,
        "Ciclo Lectivo"                         AS semestre,
        "Nombre del curso"                      AS asignatura,
        "Descripción Materia"                   AS departamento,
        "Componente"                            AS componente,
        "Componente Descripción"                AS componente_desc,
        MIN("F Inicial")                        AS fecha_inicio,
        MAX("Fecha Final")                      AS fecha_fin,
        "Nº Clase"                              AS num_clase,
        COUNT(*)                                AS horas_semana_grupo
    FROM df
    WHERE
        -- Excluir sección combinada (evita duplicados)
        ("ID Sección Combinada" IS NULL OR "ID Sección Combinada" = '')
        -- Filtro dinámico de profesor (se inyecta via Python)
        {filtro_profesor}
        -- Filtro dinámico de semestre
        {filtro_semestre}
    GROUP BY
        "Nombre profesor",
        "Numero documento docente",
        "Ciclo Lectivo",
        "Nombre del curso",
        "Descripción Materia",
        "Componente",
        "Componente Descripción",
        "Nº Clase"
),
consolidado AS (
    -- Paso 2: sumar todos los grupos de la misma materia y multiplicar × 16 semanas
    SELECT
        profesor,
        documento,
        semestre,
        asignatura,
        departamento,
        componente,
        componente_desc,
        MIN(fecha_inicio)                       AS fecha_inicio,
        MAX(fecha_fin)                          AS fecha_fin,
        SUM(horas_semana_grupo) * 16            AS total_sesiones
    FROM horas_por_grupo
    GROUP BY
        profesor,
        documento,
        semestre,
        asignatura,
        departamento,
        componente,
        componente_desc
)
SELECT
    profesor,
    documento,
    semestre,
    asignatura,
    departamento,
    componente_desc                             AS componente,
    fecha_inicio,
    fecha_fin,
    total_sesiones                              AS sesiones
FROM consolidado
ORDER BY profesor, semestre, asignatura
"""


def cargar_excel(ruta: str) -> pd.DataFrame:
    print(f"Cargando {ruta}...")
    df = pd.read_excel(ruta, dtype={
        "Numero documento docente": str,
        "Departamento":             str,
        "ID Sección Combinada":     str,
        "Nº Clase":                 str,
        "Id profesor":              str,
        "Sección Clase":            str,
        "ID Evento":                str,
    })
    # Normalizar: celdas vacías o solo espacios → None (para que el filtro IS NULL funcione)
    df["ID Sección Combinada"] = df["ID Sección Combinada"].str.strip().replace({"": None, "nan": None})
    print(f"  {len(df):,} filas · {len(df.columns)} columnas cargadas.")
    return df


def construir_filtros(profesor: str | None, semestres: list[str] | None) -> tuple[str, str]:
    filtro_profesor = ""
    if profesor:
        filtro_profesor = f"AND UPPER(\"Nombre profesor\") LIKE UPPER('%{profesor}%')"

    filtro_semestre = ""
    if semestres:
        lista = ", ".join(f"'PERIODO {s}'" for s in semestres)
        filtro_semestre = f"AND \"Ciclo Lectivo\" IN ({lista})"

    return filtro_profesor, filtro_semestre


def ejecutar_query(df: pd.DataFrame, filtro_profesor: str, filtro_semestre: str) -> pd.DataFrame:
    query = QUERY_CONSOLIDADO.format(
        filtro_profesor=filtro_profesor,
        filtro_semestre=filtro_semestre,
    )
    return duckdb.sql(query).df()


def imprimir_resultado(resultado: pd.DataFrame) -> None:
    if resultado.empty:
        print("\n⚠️  No se encontraron registros con los filtros aplicados.")
        return

    profesores = resultado["profesor"].unique()
    for prof in profesores:
        filas = resultado[resultado["profesor"] == prof]
        doc = filas["documento"].iloc[0]
        semestres = filas["semestre"].unique()

        print(f"\n{'='*70}")
        print(f"  PROFESOR : {prof}")
        print(f"  DOCUMENTO: {doc}")
        print(f"{'='*70}")

        for sem in semestres:
            filas_sem = filas[filas["semestre"] == sem]
            fecha_i = filas_sem["fecha_inicio"].iloc[0]
            fecha_f = filas_sem["fecha_fin"].iloc[0]
            print(f"\n  {sem}  |  {fecha_i} → {fecha_f}")
            print(f"  {'-'*66}")
            print(f"  {'ASIGNATURA':<35} {'SESIONES':>8}  {'DEPARTAMENTO'}")
            print(f"  {'-'*66}")
            for _, r in filas_sem.iterrows():
                print(f"  {r['asignatura']:<35} {int(r['sesiones']):>8}  {r['departamento']}")

        total = int(filas["sesiones"].sum())
        print(f"\n  {'TOTAL SESIONES':<35} {total:>8}")
        print(f"{'='*70}")


def exportar_excel(resultado: pd.DataFrame, ruta_salida: str) -> None:
    cols_export = {
        "semestre": "SEMESTRE",
        "asignatura": "ASIGNATURA",
        "sesiones": "No SESIONES",
        "departamento": "DEPARTAMENTO",
        "componente": "COMPONENTE",
        "fecha_inicio": "FECHA INICIAL",
        "fecha_fin": "FECHA FINAL",
        "profesor": "PROFESOR",
        "documento": "DOCUMENTO",
    }
    df_out = resultado.rename(columns=cols_export)[list(cols_export.values())]
    df_out.to_excel(ruta_salida, index=False)
    print(f"\n✅ Reporte exportado: {ruta_salida}")


def listar_profesores(df: pd.DataFrame) -> None:
    profesores = sorted(df["Nombre profesor"].dropna().unique())
    print(f"\n{len(profesores)} profesores encontrados:\n")
    for p in profesores:
        print(f"  {p}")


def main():
    parser = argparse.ArgumentParser(description="Consolidado de horas — Facultad de Ingeniería")
    parser.add_argument("--excel", required=True, help="Ruta al archivo Excel consolidado")
    parser.add_argument("--profesor", help="Nombre o fragmento del nombre del profesor")
    parser.add_argument("--semestre", nargs="+", help="Uno o varios semestres (ej: 2024-2 2025-1)")
    parser.add_argument("--output", help="Exportar resultado a Excel (ej: reporte.xlsx)")
    parser.add_argument("--listar-profesores", action="store_true", help="Listar todos los profesores en el archivo")
    args = parser.parse_args()

    df = cargar_excel(args.excel)

    if args.listar_profesores:
        listar_profesores(df)
        sys.exit(0)

    filtro_profesor, filtro_semestre = construir_filtros(args.profesor, args.semestre)
    resultado = ejecutar_query(df, filtro_profesor, filtro_semestre)

    imprimir_resultado(resultado)

    if args.output:
        exportar_excel(resultado, args.output)


if __name__ == "__main__":
    main()