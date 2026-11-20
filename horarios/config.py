# -*- coding: utf-8 -*-
"""
Constantes de dominio del motor de horarios.

Todo lo que es "configurable" (patrones de archivo, sedes, grupos de semanas,
turnos, rutas de recursos) vive aquí para no tener valores mágicos repartidos
por el código.
"""
import os
import re
from typing import Iterable

# ── Rutas de recursos (plantilla HTML y logo) ───────────────
PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ASSETS_DIR = os.path.join(PROJECT_ROOT, "assets")

TEMPLATE_FILE = os.path.join(ASSETS_DIR, "template_horario.html")
LOGO_FILE = os.path.join(ASSETS_DIR, "logotipo_3uteg.svg")

# ── Archivos de entrada ─────────────────────────────────────
# Excel consolidado de respaldo (solo si no hay reportes por sede).
RUTA_ENTRADA = "entrada_unificada.xlsx"
# Reportes individuales por sede.
PATRON_REPORTE_SEDE = "Reporte Horarios y Paquetes - Proceso *.xlsx"
# Reporte de Secciones: FUENTE PRINCIPAL. Una fila por sesión semanal (SESION ID)
# × currículo participante, con día, horario, salón, docente, semanas y paquetes.
# Si existe con las columnas mínimas, se usa como única fuente de datos.
PATRON_SECCIONES = "Resultados*Secciones*Proceso*.xlsx"
# Columnas mínimas para usar Secciones como fuente principal.
COLUMNAS_SECCIONES_REQUERIDAS = {"SEDE", "DIA", "HORA INICIO", "HORA FIN"}
# Renombrado Secciones → esquema interno (las demás columnas ya coinciden:
# SEDE, JORNADA, ASIGNATURA, NOMBRE, SALA, EDIFICIO, CURRICULO, NIVEL).
COLUMNAS_SECCIONES = {
    "LIGAS": "LIGA",
    "DOCENTE CODIGO": "DOCENTE_ID",
    "DOCENTE NOMBRE": "DOCENTE_NOMBRE",
    "PAQUETES": "PAQUETE",
    "SESION ID": "SESION_ID",
}

# Palabras que descartan un .xlsx como candidato a "Excel consolidado".
SKIP_XLSX_KEYWORDS = (
    "template", "logo", "mapeo", "reporte_packages", "reporte horarios y paquetes",
    "resultados", "secciones",
)

# Nombres posibles de la columna de sede en los distintos formatos de entrada.
COLUMNAS_SEDE = ("CAMPUS PLANTEL", "codigo_plantel", "SEDE")

# Renombrado de columnas de entrada → esquema interno homogéneo.
COLUMNAS_INTERNAS = {
    # Excel consolidado (inglés)
    "codigo_plantel": "SEDE",
    "codigo_modalidad": "MODALIDAD",
    "codigo_jornada": "JORNADA",
    "codigo_programa": "PROGRAMA",
    "codigo_curriculo": "CURRICULO",
    "course_code": "ASIGNATURA",
    "course_name": "NOMBRE",
    "group_label": "GROUP_LABEL",
    "link_code": "LIGA",
    "group_schedule": "GROUP_SCHEDULE",
    "instructor_code": "DOCENTE_ID",
    "instructor_name": "DOCENTE_NOMBRE",
    "sala": "SALA",
    "edificio": "EDIFICIO",
    "packages": "PAQUETE",
    "uteg_block": "UTEG_BLOCK",
    "utc_block": "UTEG_BLOCK",
    "packages_uteg_blocks_vacancies": "VACANTES_BLOQUE",
    "packages_utc_blocks_vacancies": "VACANTES_BLOQUE",
    "NIVEL": "NIVEL",
    # Reporte por sede (español)
    "CAMPUS PLANTEL": "SEDE",
    "MODALIDAD": "MODALIDAD",
    "JORNADA": "JORNADA",
    "CARRERA": "PROGRAMA",
    "CURRICULO": "CURRICULO",
    "ASIGNATURA": "ASIGNATURA",
    "NOMBRE ASIGNATURA": "NOMBRE",
    "GRUPO": "GROUP_LABEL",
    "LIGA": "LIGA",
    "HORARIOS DE LIGA": "GROUP_SCHEDULE",
    "ID DOCENTE": "DOCENTE_ID",
    "NOMBRE DOCENTE": "DOCENTE_NOMBRE",
    "SALA ASIGNADA": "SALA",
    "EDIFICIO": "EDIFICIO",
    "EDIFICIO ASIGNADO": "EDIFICIO",
    "SEMANAS": "SEMANAS",
    "PAQUETES DE LA LIGA": "PAQUETE",
    "BLOQUES UTEG": "UTEG_BLOCK",
    "BLOQUES UTC": "UTEG_BLOCK",
    "NIVELES BLOQUES": "NIVEL",
    "VACANTE DE PAQUETES": "VACANTES_BLOQUE",
}

# ── Días de la semana ───────────────────────────────────────
# Letra interna → nombre visible (el orden define las columnas de la grilla).
DIA_LETRA = {
    "L": "Lunes", "M": "Martes", "W": "Miércoles", "J": "Jueves",
    "V": "Viernes", "S": "Sábado", "D": "Domingo",
}
# El reporte de Secciones codifica el día como número (1=Lunes … 7=Domingo).
DIA_NUM_TO_LETRA = {1: "L", 2: "M", 3: "W", 4: "J", 5: "V", 6: "S", 7: "D"}
# Letra interna → abreviatura usada en el formato "Lu(07:00-09:00)".
DIA_LETRA_ABREV = {"L": "Lu", "M": "Ma", "W": "Mi", "J": "Ju", "V": "Vi", "S": "Sa", "D": "Do"}
# Abreviaturas que aparecen en "HORARIOS DE LIGA" (Lu, Ma, Mi, …) → letra interna.
DIA_PREFIX = {
    "L": "L", "M": "M", "W": "W", "J": "J", "V": "V", "S": "S", "D": "D", "F": "D",
    "LU": "L", "LUN": "L", "LUNES": "L",
    "MA": "M", "MAR": "M", "MARTES": "M",
    "MI": "W", "MIE": "W", "MIÉ": "W", "MIERCOLES": "W", "MIÉRCOLES": "W",
    "JU": "J", "JUE": "J", "JUEVES": "J",
    "VI": "V", "VIE": "V", "VIERNES": "V",
    "SA": "S", "SAB": "S", "SÁB": "S", "SABADO": "S", "SÁBADO": "S",
    "DO": "D", "DOM": "D", "DOMINGO": "D",
}

# ── Franja horaria de las grillas ───────────────────────────
HORA_INICIO = "07:00"
HORA_FIN = "22:00"
# Rango por turno (JORNADA). Hoy todos van de 07:00 a 22:00.
TURNOS = {
    turno: (HORA_INICIO, HORA_FIN)
    for turno in (
        "SABA", "DOMI", "SABAVESP", "VESPE", "NOCTJ", "INTE", "MATU", "MATUBACH",
        "MATUEXT", "MEDOMI", "MENOCTJ", "MESABA", "MESABAVESP", "NOCT", "VESP",
    )
}

# ── Sedes y grupos de semanas ───────────────────────────────
# Las sedes se detectan dinámicamente desde los archivos de entrada.
SEDES_CONOCIDAS: list[str] = []

# Semanas de materias de ciclo completo (respaldo por defecto: calendario de 14 semanas)
SEMANAS_SEMESTRALES = {"(1 - 14)"}

# ── Jerarquía de carpetas y periodos por tipo de horario ─────
# Valores por defecto (se recalculan dinámicamente con `construir_periodos()`
# a partir de los rangos reales presentes en la columna SEMANAS del Excel).
#   profesor/ -> módulos (cada uno integra materias del ciclo completo si existen)
#   sala/     -> módulos (cada uno integra ocupaciones del ciclo completo si existen)
#   grupo/    -> ciclo completo (si existe) + cada módulo

PERIODOS_PROFESOR = [
    {
        "id": "1_7",
        "nombre": "Semanas 1 – 7",
        "subtitulo": "Módulo 1 (Semanas 1–7)",
        "filtro": {"(1 - 7)"} | SEMANAS_SEMESTRALES,
        "vacias": True,
    },
    {
        "id": "8_14",
        "nombre": "Semanas 8 – 14",
        "subtitulo": "Módulo 2 (Semanas 8–14)",
        "filtro": {"(8 - 14)"} | SEMANAS_SEMESTRALES,
        "vacias": True,
    },
]

PERIODOS_SALA = [
    {
        "id": "1_7",
        "nombre": "Semanas 1 – 7",
        "subtitulo": "Módulo 1 (Semanas 1–7)",
        "filtro": {"(1 - 7)"} | SEMANAS_SEMESTRALES,
        "vacias": True,
    },
    {
        "id": "8_14",
        "nombre": "Semanas 8 – 14",
        "subtitulo": "Módulo 2 (Semanas 8–14)",
        "filtro": {"(8 - 14)"} | SEMANAS_SEMESTRALES,
        "vacias": True,
    },
]

PERIODOS_GRUPO = [
    {
        "id": "1_14",
        "nombre": "Semanas 1 – 14",
        "subtitulo": "Semanas 1–14",
        "filtro": {"(1 - 14)"},
        "vacias": True,
    },
    {
        "id": "1_7",
        "nombre": "Semanas 1 – 7",
        "subtitulo": "Semanas 1–7",
        "filtro": {"(1 - 7)"},
        "vacias": False,
    },
    {
        "id": "8_14",
        "nombre": "Semanas 8 – 14",
        "subtitulo": "Semanas 8–14",
        "filtro": {"(8 - 14)"},
        "vacias": False,
    },
]

# Opciones por defecto para el selector de semanas de la interfaz web (Paso 3)
GRUPOS_SEMANAS = [
    {"id": "1_14", "nombre": "Semanas 1 – 14"},
    {"id": "1_7", "nombre": "Semanas 1 – 7"},
    {"id": "8_14", "nombre": "Semanas 8 – 14"},
]

# Categorías principales de salida
CATEGORIAS_HORARIO = ("profesor", "sala", "grupo")

# Alias de compatibilidad hacia atrás
MODULOS_DOCENTES = PERIODOS_PROFESOR


# ── Construcción dinámica de periodos a partir de SEMANAS ───
_RE_RANGO_SEMANAS = re.compile(r"^\(?\s*(\d+)\s*[-–]\s*(\d+)\s*\)?$")
_RE_PERIODO_ID = re.compile(r"^(\d+)_(\d+)$")


def parsear_rango_semanas(valor: str) -> tuple[int, int] | None:
    """'(1 - 7)' → (1, 7); None si no corresponde a un rango numérico."""
    m = _RE_RANGO_SEMANAS.match(str(valor).strip())
    if not m:
        return None
    ini, fin = int(m.group(1)), int(m.group(2))
    return (ini, fin) if ini <= fin else (fin, ini)


def rango_desde_bits(bits: str) -> str | None:
    """Máscara de semanas → etiqueta SEMANAS.

    '11111110000000' → '(1 - 7)';  '00000001111111' → '(8 - 14)';
    rachas no contiguas se unen con ', ' ('1100011' → '(1 - 2), (6 - 7)').
    """
    rachas, ini = [], None
    for i, b in enumerate(str(bits).strip() + "0", start=1):
        if b == "1" and ini is None:
            ini = i
        elif b != "1" and ini is not None:
            rachas.append(f"({ini} - {i - 1})")
            ini = None
    return ", ".join(rachas) or None


def etiqueta_periodo(period_id: str) -> str:
    """'1_7' → 'Semanas 1–7', 'general' → 'Todas las semanas'."""
    if period_id == "general":
        return "Todas las semanas"
    m = _RE_PERIODO_ID.match(str(period_id).strip())
    if m:
        return f"Semanas {m.group(1)}–{m.group(2)}"
    return str(period_id)


def construir_periodos(semanas_values: Iterable | None) -> dict[str, list[dict]]:
    """Construye los periodos de `profesor`, `sala`, `grupo` y `grupos_semanas`
    dinámicamente a partir de los valores reales presentes en la columna `SEMANAS`.

    Soporta cualquier calendario académico (14 semanas: 1–7 / 8–14 / 1–14;
    16 semanas: 1–8 / 9–16 / 1–16; 12 semanas; periodo único; o sin columna SEMANAS).
    """
    valores_limpios: list[str] = []
    seen: set[str] = set()
    if semanas_values is not None:
        for v in semanas_values:
            if v is None:
                continue
            s = str(v).strip()
            if not s or s.lower() in ("nan", "none", "<na>"):
                continue
            if s not in seen:
                seen.add(s)
                valores_limpios.append(s)

    if not valores_limpios:
        unico = {
            "id": "general",
            "nombre": "Todas las semanas",
            "subtitulo": None,
            "filtro": set(),
            "vacias": True,
        }
        return {
            "profesor": [dict(unico)],
            "sala": [dict(unico)],
            "grupo": [dict(unico)],
            "grupos_semanas": [{"id": "general", "nombre": "Todas las semanas"}],
        }

    # Separar rangos numéricos "(ini - fin)" de posibles etiquetas textuales
    por_rango: dict[tuple[int, int], set[str]] = {}
    otros: list[str] = []
    for s in valores_limpios:
        rango = parsear_rango_semanas(s)
        if rango is not None:
            por_rango.setdefault(rango, set()).add(s)
        else:
            otros.append(s)

    # Agrupar rangos equivalentes que solo difieren en la última semana (ej. 1-15 y 1-16, o 9-15 y 9-16)
    rangos_unificados: dict[tuple[int, int], set[str]] = {}
    if por_rango:
        max_fin = max(fin for _, fin in por_rango)
        for (ini, fin), etiquetas in sorted(por_rango.items()):
            destino = (ini, max_fin) if (fin == max_fin - 1 and (ini, max_fin) in por_rango) else (ini, fin)
            rangos_unificados.setdefault(destino, set()).update(etiquetas)

    # Clasificar rangos en ciclo completo vs modulares
    completos: list[tuple[tuple[int, int], set[str]]] = []
    modulares: list[tuple[tuple[int, int], set[str]]] = []
    if rangos_unificados:
        min_ini = min(ini for ini, _ in rangos_unificados)
        max_fin = max(fin for _, fin in rangos_unificados)
        tiene_submodulos = any(
            (ini > min_ini or fin < max_fin) for (ini, fin) in rangos_unificados
        )
        for (ini, fin), filtro in sorted(rangos_unificados.items()):
            if tiene_submodulos and ini == min_ini and fin == max_fin:
                completos.append(((ini, fin), filtro))
            else:
                modulares.append(((ini, fin), filtro))

    filtro_completo: set[str] = set().union(*(f for _, f in completos)) if completos else set()
    periodos_profesor: list[dict] = []
    periodos_sala: list[dict] = []
    periodos_grupo: list[dict] = []

    # 1) Ciclo completo en `grupo/` (si existe)
    for (ini, fin), filtro in completos:
        periodos_grupo.append({
            "id": f"{ini}_{fin}",
            "nombre": f"Semanas {ini} – {fin}",
            "subtitulo": f"Semanas {ini}–{fin}",
            "filtro": set(filtro),
            "vacias": True,
        })

    # 2) Módulos (o rango único si no hay submódulos)
    es_multimodulo = len(modulares) > 1 or bool(completos)
    for idx, ((ini, fin), filtro) in enumerate(modulares, start=1):
        pid = f"{ini}_{fin}"
        nombre = f"Semanas {ini} – {fin}"
        subtitulo_mod = f"Módulo {idx} (Semanas {ini}–{fin})" if es_multimodulo else f"Semanas {ini}–{fin}"
        filtro_integrado = set(filtro) | filtro_completo

        periodos_profesor.append({
            "id": pid,
            "nombre": nombre,
            "subtitulo": subtitulo_mod,
            "filtro": set(filtro_integrado),
            "vacias": True,
        })
        periodos_sala.append({
            "id": pid,
            "nombre": nombre,
            "subtitulo": subtitulo_mod,
            "filtro": set(filtro_integrado),
            "vacias": True,
        })
        periodos_grupo.append({
            "id": pid,
            "nombre": nombre,
            "subtitulo": f"Semanas {ini}–{fin}",
            "filtro": set(filtro),
            "vacias": not es_multimodulo,
        })

    # 3) Etiquetas no numéricas (si aparecieran)
    for etiqueta in sorted(otros):
        pid = re.sub(r'[\\/*?"<>| ]', "_", etiqueta)
        entrada = {
            "id": pid,
            "nombre": etiqueta,
            "subtitulo": etiqueta,
            "filtro": {etiqueta},
            "vacias": False,
        }
        periodos_profesor.append(dict(entrada))
        periodos_sala.append(dict(entrada))
        periodos_grupo.append(dict(entrada))

    grupos_semanas = [{"id": p["id"], "nombre": p["nombre"]} for p in periodos_grupo]

    return {
        "profesor": periodos_profesor,
        "sala": periodos_sala,
        "grupo": periodos_grupo,
        "grupos_semanas": grupos_semanas,
    }



