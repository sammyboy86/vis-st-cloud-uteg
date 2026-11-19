# -*- coding: utf-8 -*-
"""
Utilidades genéricas de texto y valores de celdas.
"""
import base64
import mimetypes
import pathlib
import re

import pandas as pd

_CARACTERES_INVALIDOS = re.compile(r'[\\/*?"<>| ]')


def slugify(texto: str) -> str:
    """Convierte un texto en un nombre de archivo/carpeta seguro."""
    return _CARACTERES_INVALIDOS.sub("_", str(texto))


def limpiar_texto(val) -> str:
    """Normaliza el valor de una celda a str: NaN/None → '', 12.0 → '12'."""
    if val is None or pd.isna(val):
        return ""
    if isinstance(val, float) and val.is_integer():
        return str(int(val))
    s = str(val).strip()
    if s.lower() in ("nan", "none", "null"):
        return ""
    if s.endswith(".0"):
        try:
            return str(int(float(s)))
        except ValueError:
            pass
    return s


def formato_sala(sala, edificio=None) -> str:
    """Texto visible del salón: 'SALA - EDIFICIO', o el que exista de los dos."""
    sala_str = limpiar_texto(sala)
    edif_str = limpiar_texto(edificio)
    if sala_str and edif_str:
        return f"{sala_str} - {edif_str}"
    return sala_str or edif_str


def split_lista(cell) -> list[str]:
    """Separa una celda 'a, b, c' en ['a', 'b', 'c'] (NaN → [])."""
    if pd.isna(cell):
        return []
    return [str(x).strip() for x in str(cell).split(",") if str(x).strip()]


def valor_predominante(serie: pd.Series, default: str):
    """Valor más frecuente de la serie, o `default` si está toda vacía."""
    if serie.isna().all():
        return default
    return serie.mode().iat[0]


def logo_data_uri(logo_path: str | pathlib.Path) -> str:
    """Devuelve un data URI base64 con el MIME correcto (SVG, PNG, etc.)."""
    path = pathlib.Path(logo_path)
    if not path.exists():
        return ""
    ext = path.suffix.lower()
    if ext == ".svg":
        mime = "image/svg+xml"
    else:
        mime = mimetypes.guess_type(path.name)[0] or "image/png"
    b64 = base64.b64encode(path.read_bytes()).decode()
    return f"data:{mime};base64,{b64}"


def asegurar_columna_bloque(df: pd.DataFrame) -> pd.DataFrame:
    """Renombra 'uteg_block'/'utc_block' (cualquier capitalización) a 'UTEG_BLOCK' si hace falta."""
    if "UTEG_BLOCK" in df.columns:
        return df
    for c in df.columns:
        if str(c).strip().lower() in ("uteg_block", "utc_block"):
            return df.rename(columns={c: "UTEG_BLOCK"})
    return df
