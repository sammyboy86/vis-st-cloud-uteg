# -*- coding: utf-8 -*-
"""
Lectura de los Excel de entrada.

Fuentes soportadas (en orden de preferencia):
  1. Reporte Secciones  → "Resultados*Secciones*Proceso*.xlsx" (FUENTE ÚNICA:
     una fila por sesión con día, horario, salón, docente, semanas y paquetes;
     los bloques UTEG se reconstruyen a partir de los paquetes)
  2. Reportes por sede  → "Reporte Horarios y Paquetes - Proceso *.xlsx" (respaldo)
  3. Excel consolidado  → cualquier otro .xlsx de la carpeta (respaldo)

Todas las tablas se devuelven con el esquema interno de `config.COLUMNAS_INTERNAS`.
"""
import datetime as dt
import glob
import logging
import os
import re

import pandas as pd

from horarios.config import (
    COLUMNAS_INTERNAS,
    COLUMNAS_SECCIONES,
    COLUMNAS_SECCIONES_REQUERIDAS,
    COLUMNAS_SEDE,
    DIA_LETRA_ABREV,
    DIA_NUM_TO_LETRA,
    PATRON_REPORTE_SEDE,
    PATRON_SECCIONES,
    RUTA_ENTRADA,
    SKIP_XLSX_KEYWORDS,
    rango_desde_bits,
)
from horarios.utils import formato_sala, limpiar_texto, split_lista

# Tipo del mapa de salones: (SEDE, LIGA, letra_día) → "SALA - EDIFICIO"
SalonesMap = dict[tuple[str, str, str], str]


# ═════════════════════════════════════════════════════════════
# Búsqueda de archivos
# ═════════════════════════════════════════════════════════════

def _es_temporal_office(path: str) -> bool:
    return os.path.basename(path).startswith("~$")


def find_sede_report_excels(workdir=None) -> list[str]:
    """Reportes individuales por sede: 'Reporte Horarios y Paquetes - Proceso NNN.xlsx'.

    Si no hay coincidencias exactas con el patrón, acepta cualquier .xlsx con
    'reporte horarios' en el nombre.
    """
    wd = workdir or "."
    files = [f for f in glob.glob(os.path.join(wd, PATRON_REPORTE_SEDE)) if not _es_temporal_office(f)]
    if not files:
        files = [
            f for f in glob.glob(os.path.join(wd, "*.xlsx"))
            if not _es_temporal_office(f) and "reporte horarios" in os.path.basename(f).lower()
        ]
    return sorted(files)


def find_secciones_excels(workdir=None) -> list[str]:
    """Reporte(s) 'Resultados - Secciones - Proceso NNN.xlsx'.

    Trae, fila por fila, el salón real de cada sesión (LIGA + día) y su edificio,
    a diferencia del salón único por liga que trae 'Reporte Horarios y Paquetes'.
    Si no hay coincidencias con el patrón, acepta cualquier .xlsx con 'secciones'
    en el nombre.
    """
    wd = workdir or "."
    files = [f for f in glob.glob(os.path.join(wd, PATRON_SECCIONES)) if not _es_temporal_office(f)]
    if not files:
        files = [
            f for f in glob.glob(os.path.join(wd, "*.xlsx"))
            if not _es_temporal_office(f) and "secciones" in os.path.basename(f).lower()
        ]
    return sorted(files)


def find_input_excel(prefer: str = None, workdir=None) -> str:
    """Excel consolidado de respaldo: `prefer` si existe, si no el .xlsx más reciente."""
    wd = workdir or "."
    if prefer:
        prefer_abs = prefer if os.path.isabs(prefer) else os.path.join(wd, prefer)
        if os.path.exists(prefer_abs):
            return prefer_abs

    candidates = []
    for f in glob.glob(os.path.join(wd, "*.xlsx")):
        low = os.path.basename(f).lower()
        if _es_temporal_office(f) or any(k in low for k in SKIP_XLSX_KEYWORDS):
            continue
        candidates.append(f)

    if not candidates:
        raise FileNotFoundError("No se encontró ningún .xlsx de entrada en la carpeta actual.")
    return max(candidates, key=os.path.getmtime)


# ═════════════════════════════════════════════════════════════
# Detección de sedes
# ═════════════════════════════════════════════════════════════

def _columna_sede(columns) -> str | None:
    return next((c for c in columns if str(c).strip() in COLUMNAS_SEDE), None)


def _columna_semanas(columns) -> str | None:
    return next((c for c in columns if str(c).strip().upper() == "SEMANAS"), None)


def _rewind(archivo) -> None:
    if hasattr(archivo, "seek"):
        archivo.seek(0)


def detect_sedes(archivo) -> list[str]:
    """Códigos de sede presentes en un Excel (lee solo la columna de sede).

    `archivo` puede ser una ruta o un objeto tipo archivo (p. ej. un upload de
    Streamlit); en ese caso se rebobina para poder leerlo de nuevo después.
    """
    try:
        _rewind(archivo)
        header = pd.read_excel(archivo, engine="openpyxl", nrows=0)
        _rewind(archivo)
        col = _columna_sede(header.columns)
        if col is not None:
            df = pd.read_excel(archivo, engine="openpyxl", usecols=[col])
            _rewind(archivo)
            sedes = df[col].dropna().astype(str).str.strip()
            return sorted(s for s in sedes.unique().tolist() if s)
    except Exception:
        _rewind(archivo)
    return []


def detect_semanas(archivo) -> list[str]:
    """Valores únicos de semanas en un Excel (lee solo las columnas necesarias).

    Usa la columna `SEMANAS` (Reporte Horarios y Paquetes) o, si no existe, la
    máscara `S1…Sn` / `SEMANAS CODIGO` del reporte de Secciones.
    """
    try:
        _rewind(archivo)
        header = pd.read_excel(archivo, engine="openpyxl", nrows=0)
        _rewind(archivo)
        col = _columna_semanas(header.columns)
        if col is not None:
            df = pd.read_excel(archivo, engine="openpyxl", usecols=[col])
            _rewind(archivo)
            vals = df[col].dropna().astype(str).str.strip()
        else:
            cols = [c for c in header.columns if _es_columna_semana(c) or str(c).strip() == "SEMANAS CODIGO"]
            if not cols:
                return []
            df = pd.read_excel(archivo, engine="openpyxl", usecols=cols)
            _rewind(archivo)
            df.columns = [str(c).strip() for c in df.columns]
            vals = _semanas_secciones(df).dropna().astype(str)
        return sorted(v for v in vals.unique().tolist() if v and v.lower() not in ("nan", "none", "<na>"))
    except Exception:
        _rewind(archivo)
    return []


def _sedes_en_archivo(path: str) -> set[str]:
    """Lee solo la columna de sede para decidir si el archivo entra al filtro."""
    header = pd.read_excel(path, engine="openpyxl", nrows=0)
    col = _columna_sede(header.columns)
    if col is None:
        return set()
    series = pd.read_excel(path, usecols=[col], engine="openpyxl")[col]
    return {s for s in series.dropna().astype(str).str.strip().unique() if s}


# ═════════════════════════════════════════════════════════════
# Carga de datos principales
# ═════════════════════════════════════════════════════════════

def load_data(path: str) -> pd.DataFrame:
    """Lee un Excel y renombra sus columnas al esquema interno."""
    df = pd.read_excel(path, engine="openpyxl")
    df.columns = [str(c).strip() for c in df.columns]
    df = df.rename(columns=COLUMNAS_INTERNAS)
    if "SEDE" in df.columns:
        mask = df["SEDE"].notna()
        df.loc[mask, "SEDE"] = df.loc[mask, "SEDE"].astype(str).str.strip()
    return df


def _filtrar_sedes(df: pd.DataFrame, sedes: set[str]) -> pd.DataFrame:
    sedes_norm = {str(s).strip() for s in sedes}
    return df[df["SEDE"].astype(str).str.strip().isin(sedes_norm)].copy()


def load_data_unificado(path_entrada=None, sedes_filtro=None, workdir=None) -> pd.DataFrame:
    """Carga la fuente principal de datos.

    1. Reporte de Secciones (si existe con las columnas mínimas) → fuente única.
    2. Reportes por sede "Horarios y Paquetes" concatenados.
    3. Excel consolidado.

    sedes_filtro: set/list de códigos de plantel; solo se cargan archivos de esas sedes.
    """
    if sedes_filtro is not None:
        sedes_filtro = {str(s) for s in sedes_filtro}

    if not path_entrada and find_secciones_principales(workdir=workdir):
        return load_data_secciones(sedes_filtro=sedes_filtro, workdir=workdir)

    sede_files = find_sede_report_excels(workdir=workdir)
    if sede_files:
        return _cargar_reportes_por_sede(sede_files, sedes_filtro)
    return _cargar_consolidado(path_entrada, sedes_filtro, workdir)


def _cargar_reportes_por_sede(sede_files: list[str], sedes_filtro: set[str] | None) -> pd.DataFrame:
    dfs = []
    for f in sede_files:
        sedes_archivo = _sedes_en_archivo(f)
        if not sedes_archivo:
            logging.warning("Archivo sin sede identificable, se omite: %s", f)
            continue
        if sedes_filtro and not sedes_archivo.intersection(sedes_filtro):
            logging.info("Omitido %s: sede(s) %s fuera del filtro", f, ", ".join(sorted(sedes_archivo)))
            continue

        df = load_data(f)
        if df.empty:
            logging.warning("Archivo vacío, se omite: %s", f)
            continue
        if sedes_filtro:
            df = _filtrar_sedes(df, sedes_filtro)
            if df.empty:
                continue

        sedes = sorted(df["SEDE"].dropna().astype(str).unique())
        logging.info("Cargado %s: %d filas, sede(s): %s", f, len(df), ", ".join(sedes))
        dfs.append(df)

    if not dfs:
        if sedes_filtro:
            raise ValueError("No se encontraron reportes por sede para: " + ", ".join(sorted(sedes_filtro)))
        raise ValueError("Todos los reportes por sede están vacíos.")
    return pd.concat(dfs, ignore_index=True)


def _cargar_consolidado(path_entrada, sedes_filtro: set[str] | None, workdir) -> pd.DataFrame:
    path = path_entrada or find_input_excel(RUTA_ENTRADA, workdir=workdir)
    logging.info("Sin reportes por sede; usando consolidado: %s", path)
    df = load_data(path)
    if sedes_filtro:
        df = _filtrar_sedes(df, sedes_filtro)
        if df.empty:
            raise ValueError("El consolidado no tiene filas para las sedes: " + ", ".join(sorted(sedes_filtro)))
    return df


# ═════════════════════════════════════════════════════════════
# Reporte de Secciones como fuente única
# ═════════════════════════════════════════════════════════════

def find_secciones_principales(workdir=None) -> list[str]:
    """Reportes de Secciones con las columnas mínimas para ser fuente principal."""
    out = []
    for f in find_secciones_excels(workdir=workdir):
        try:
            header = pd.read_excel(f, engine="openpyxl", nrows=0)
        except Exception as exc:
            logging.warning("No se pudo leer %s: %s", f, exc)
            continue
        cols = {str(c).strip() for c in header.columns}
        if COLUMNAS_SECCIONES_REQUERIDAS <= cols:
            out.append(f)
        else:
            logging.info(
                "Secciones %s sin columnas %s; solo se usará para salones.",
                f, COLUMNAS_SECCIONES_REQUERIDAS - cols,
            )
    return out


def es_base_secciones(df: pd.DataFrame) -> bool:
    """True si el DataFrame proviene de `load_data_secciones` (salón ya resuelto por sesión)."""
    return "SESION_ID" in df.columns


def load_data_secciones(sedes_filtro=None, workdir=None) -> pd.DataFrame:
    """Construye la tabla interna completa solo a partir del reporte de Secciones.

    Cada fila de Secciones es una sesión semanal (SESION ID) × currículo
    participante. Se conserva esa granularidad y se agregan:

      GROUP_SCHEDULE  'Lu(08:00-11:00)'  (DIA + HORA INICIO + HORA FIN)
      SEMANAS         '(1 - 7)'          (máscara S1…Sn / SEMANAS CODIGO)
      UTEG_BLOCK      'AME-MATU-LIDGSEP26C-N2-P1, …'  (un bloque por paquete
                      de la población CURRICULO + NIVEL en su sede y jornada)

    El salón queda en la propia fila (SALA/EDIFICIO), por lo que no hace falta
    el mapa (SEDE, LIGA, día) → salón.
    """
    archivos = find_secciones_principales(workdir=workdir)
    if not archivos:
        raise FileNotFoundError(
            f"No se encontró un reporte de Secciones (patrón: {PATRON_SECCIONES}) con las columnas "
            f"{sorted(COLUMNAS_SECCIONES_REQUERIDAS)}."
        )

    dfs = []
    for f in archivos:
        df = pd.read_excel(f, engine="openpyxl", sheet_name=0)
        df.columns = [str(c).strip() for c in df.columns]
        if sedes_filtro:
            df = _filtrar_sedes(df, sedes_filtro)
        if df.empty:
            continue
        df = _normalizar_secciones(df, f)
        logging.info(
            "Cargado Secciones %s: %d filas, %d sesiones, sede(s): %s",
            f, len(df), df["SESION_ID"].nunique() if "SESION_ID" in df else len(df),
            ", ".join(sorted(df["SEDE"].unique())),
        )
        dfs.append(df)

    if not dfs:
        if sedes_filtro:
            raise ValueError("El reporte de Secciones no tiene filas para: " + ", ".join(sorted(sedes_filtro)))
        raise ValueError("El reporte de Secciones está vacío.")
    return pd.concat(dfs, ignore_index=True)


def _normalizar_secciones(df: pd.DataFrame, origen: str = "") -> pd.DataFrame:
    letras = df["DIA"].map(_letra_dia)
    ini = df["HORA INICIO"].map(hora_hhmm)
    fin = df["HORA FIN"].map(hora_hhmm)
    validas = letras.notna() & ini.notna() & fin.notna()
    if not validas.all():
        logging.warning(
            "Secciones %s: %d filas sin día u horario válido se omiten.", origen, int((~validas).sum())
        )

    semanas = _semanas_secciones(df)
    df = df.rename(columns=COLUMNAS_SECCIONES)
    df["SEDE"] = df["SEDE"].map(limpiar_texto)
    df["GROUP_SCHEDULE"] = letras.map(DIA_LETRA_ABREV) + "(" + ini + "-" + fin + ")"
    df["SEMANAS"] = semanas
    if "SESION_ID" not in df.columns:
        df["SESION_ID"] = range(len(df))

    paquetes = df["PAQUETE"] if "PAQUETE" in df.columns else pd.Series("", index=df.index)
    pkg_listas = paquetes.map(lambda v: [limpiar_texto(p) for p in split_lista(limpiar_texto(v))])
    df["PAQUETE"] = pkg_listas.map(", ".join)
    df["UTEG_BLOCK"] = [
        ", ".join(_id_bloque(r, p) for p in pkgs)
        for (_, r), pkgs in zip(df.iterrows(), pkg_listas)
    ]
    return df[validas].copy()


def _id_bloque(row: pd.Series, paquete: str) -> str:
    """Identificador legible del bloque: SEDE-JORNADA-CURRICULO-N{NIVEL}-P{paquete}.

    Termina en número para que `bloques.py` (base + número → letra) siga funcionando.
    """
    partes = [limpiar_texto(row.get(c)) for c in ("SEDE", "JORNADA", "CURRICULO")]
    nivel = limpiar_texto(row.get("NIVEL"))
    return "-".join([p for p in partes if p] + [f"N{nivel}" if nivel else "N0", f"P{paquete}"])


def _es_columna_semana(col) -> bool:
    return re.fullmatch(r"S\d+", str(col).strip()) is not None


def _semanas_secciones(df: pd.DataFrame) -> pd.Series:
    """Etiqueta SEMANAS por fila a partir de S1…Sn o, en su defecto, de SEMANAS CODIGO."""
    cols = sorted((c for c in df.columns if _es_columna_semana(c)), key=lambda c: int(str(c).strip()[1:]))
    if cols:
        bits = df[cols].apply(lambda s: s.map(lambda v: "1" if limpiar_texto(v) == "1" else "0"))
        bits = bits.agg("".join, axis=1)
    elif "SEMANAS CODIGO" in df.columns:
        bits = df["SEMANAS CODIGO"].map(limpiar_texto)
        # Si Excel guardó la máscara como número se pierden los ceros a la izquierda.
        largo = int(bits.str.len().max() or 0)
        bits = bits.str.zfill(largo)
    else:
        return pd.Series(pd.NA, index=df.index, dtype="object")
    return bits.map(rango_desde_bits)


def hora_hhmm(valor) -> str | None:
    """Normaliza una hora de Excel a 'HH:MM'.

    Acepta fracción de día (0.3333 → 08:00), datetime/time, Timedelta,
    'HH:MM[:SS]' y enteros tipo HHMM (800, 1730).
    """
    if isinstance(valor, (dt.datetime, dt.time)):
        return f"{valor:%H:%M}"
    if valor is None or pd.isna(valor):
        return None
    if isinstance(valor, pd.Timedelta):
        minutos = round(valor.total_seconds() / 60)
    elif isinstance(valor, (int, float)):
        f = float(valor)
        if 0 <= f < 1:
            minutos = round(f * 24 * 60)
        elif f.is_integer() and 0 <= f < 2400 and int(f) % 100 < 60:
            minutos = int(f) // 100 * 60 + int(f) % 100
        else:
            return None
    else:
        s = str(valor).strip()
        m = re.match(r"^(\d{1,2}):(\d{2})", s)
        if m:
            minutos = int(m.group(1)) * 60 + int(m.group(2))
        else:
            try:
                return hora_hhmm(float(s))
            except ValueError:
                return None
    return f"{minutos // 60:02d}:{minutos % 60:02d}"


# ═════════════════════════════════════════════════════════════
# Salón real por (SEDE, LIGA, día) — reporte de Secciones
# ═════════════════════════════════════════════════════════════

def salones_map_para(df: pd.DataFrame, sedes_filtro=None, workdir=None) -> "SalonesMap":
    """Mapa de salones solo si la base NO es Secciones (allí el salón ya va por fila)."""
    if es_base_secciones(df):
        return {}
    return load_salones_por_liga_dia(sedes_filtro=sedes_filtro, workdir=workdir)


def load_salones_por_liga_dia(sedes_filtro=None, workdir=None) -> SalonesMap:
    """Construye el mapa (SEDE, LIGA, día) → salón real (+ edificio si existe).

    Es el complemento que le falta a "Reporte Horarios y Paquetes" para poder
    mostrar un salón distinto por día cuando la liga cambia de salón entre sesiones.
    """
    archivos = find_secciones_excels(workdir=workdir)
    if not archivos:
        logging.warning(
            "No se encontró el reporte de Secciones (patrón: %s). Se usará el salón único "
            "de la liga para todos sus días (comportamiento anterior, puede ser incorrecto "
            "si la liga cambia de salón entre sesiones).",
            PATRON_SECCIONES,
        )
        return {}

    sedes_f = {str(s) for s in sedes_filtro} if sedes_filtro else None
    mapa: SalonesMap = {}
    conflictos = set()
    total_filas = 0

    for f in archivos:
        df = pd.read_excel(f, engine="openpyxl")
        df.columns = [str(c).strip() for c in df.columns]
        faltantes = {"SEDE", "LIGAS", "DIA", "SALA"} - set(df.columns)
        if faltantes:
            logging.warning("Secciones %s: faltan columnas %s; se omite este archivo.", f, faltantes)
            continue
        if sedes_f:
            sedes_norm = {str(s).strip() for s in sedes_f}
            df = df[df["SEDE"].astype(str).str.strip().isin(sedes_norm)]

        col_edificio = next(
            (c for c in df.columns if str(c).strip().upper() in ("EDIFICIO", "EDIFICIO ASIGNADO")),
            None,
        )

        for _, r in df.iterrows():
            total_filas += 1
            letra = _letra_dia(r.get("DIA"))
            if pd.isna(r.get("SALA")) or letra is None:
                continue

            edificio = r.get(col_edificio) if col_edificio else None
            texto_sala = formato_sala(r.get("SALA"), edificio)
            if not texto_sala:
                continue

            sede_val = str(r.get("SEDE")).strip()
            for liga in split_lista(r.get("LIGAS")):
                key = (sede_val, liga, letra)
                if key in mapa and mapa[key] != texto_sala:
                    conflictos.add(key)
                mapa[key] = texto_sala  # si hay filas duplicadas que difieren, gana la última

    if conflictos:
        logging.warning(
            "Secciones: %d combinaciones (SEDE, LIGA, día) tienen salones/edificios distintos entre "
            "filas repetidas; se usó el último valor leído. Ejemplos: %s",
            len(conflictos), sorted(conflictos)[:5],
        )
    logging.info(
        "Salones y edificios por día cargados desde Secciones: %d combinaciones (SEDE, LIGA, día) "
        "a partir de %d filas.",
        len(mapa), total_filas,
    )
    return mapa


def _letra_dia(dia_num) -> str | None:
    """1 → 'L', 2 → 'M', … ; None si el valor no es un día válido."""
    if pd.isna(dia_num):
        return None
    try:
        return DIA_NUM_TO_LETRA.get(int(dia_num))
    except (TypeError, ValueError):
        return None
