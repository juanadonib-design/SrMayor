"""Lógica de lectura y extracción de códigos revisados desde archivos Excel.

- Lee .xlsx y .xls (todas las hojas).
- Detecta automáticamente la fila de encabezados (donde estén "Revisada" y "Código").
- Extrae los códigos de las filas donde Revisada = Verdadero.
- Permite filtrar por Número Libramiento.
"""
import io
import re
import unicodedata

import numpy as np
import pandas as pd

COL_HOJA = "Hoja"
COL_CODIGO = "Código"
COL_LIBRAMIENTO = "Número Libramiento"

_TRUE_VALUES = {"verdadero", "true"}


def norm(value) -> str:
    """Minúsculas, sin acentos y sin espacios repetidos."""
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass
    s = unicodedata.normalize("NFKD", str(value))
    s = "".join(c for c in s if not unicodedata.combining(c))
    return re.sub(r"\s+", " ", s).strip().lower()


def clean_value(value) -> str:
    """Convierte una celda a texto limpio (123.0 -> '123')."""
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass
    if isinstance(value, (float, np.floating)) and float(value).is_integer():
        return str(int(value))
    s = str(value).strip()
    return "" if s.lower() in {"nan", "none", "nat"} else s


def is_true(value) -> bool:
    if isinstance(value, (bool, np.bool_)):
        return bool(value)
    return norm(value) in _TRUE_VALUES


def _find(headers, exact=(), startswith=(), contains=()):
    for i, h in enumerate(headers):
        if h in exact:
            return i
    for i, h in enumerate(headers):
        if any(h.startswith(p) for p in startswith):
            return i
    for i, h in enumerate(headers):
        if any(c in h for c in contains):
            return i
    return None


def _find_header_row(raw: pd.DataFrame, max_rows: int = 40):
    for i in range(min(max_rows, len(raw))):
        headers = [norm(v) for v in raw.iloc[i].tolist()]
        i_rev = _find(headers, exact=("revisada",), startswith=("revisad",))
        i_cod = _find(headers, exact=("codigo",), startswith=("codigo",))
        if i_rev is not None and i_cod is not None:
            return i, headers, i_rev, i_cod
    return None


def _read_all_sheets(data: bytes, filename: str):
    ext = filename.lower().rsplit(".", 1)[-1]
    engines = ["xlrd", "openpyxl"] if ext == "xls" else ["openpyxl", "xlrd"]
    last_error = None
    for engine in engines:
        try:
            return pd.read_excel(
                io.BytesIO(data), sheet_name=None, header=None, engine=engine, dtype=object
            )
        except Exception as exc:  # noqa: BLE001 - probamos el otro motor
            last_error = exc
    raise ValueError(f"No se pudo leer el archivo Excel: {last_error}")


def extract_reviewed(data: bytes, filename: str):
    """Devuelve (DataFrame, hay_libramiento, notas).

    El DataFrame tiene las columnas: Hoja, Código, Número Libramiento.
    """
    sheets = _read_all_sheets(data, filename)
    frames, notes, has_lib = [], [], False

    for name, raw in sheets.items():
        found = _find_header_row(raw)
        if found is None:
            notes.append(f"Hoja «{name}»: no tiene columnas «Revisada» y «Código».")
            continue
        h, headers, i_rev, i_cod = found
        i_lib = _find(headers, exact=("numero libramiento",), contains=("libramiento",))
        has_lib = has_lib or i_lib is not None

        body = raw.iloc[h + 1 :]
        sel = body[body.iloc[:, i_rev].map(is_true)]
        out = pd.DataFrame(
            {
                COL_HOJA: name,
                COL_CODIGO: sel.iloc[:, i_cod].map(clean_value),
                COL_LIBRAMIENTO: sel.iloc[:, i_lib].map(clean_value) if i_lib is not None else "",
            }
        )
        out = out[out[COL_CODIGO] != ""]
        frames.append(out)

    if not frames:
        raise ValueError(
            "No se encontró ninguna hoja con las columnas «Revisada» y «Código». "
            + " ".join(notes)
        )
    return pd.concat(frames, ignore_index=True), has_lib, notes


def filter_by_libramiento(df: pd.DataFrame, text: str) -> pd.DataFrame:
    """Filtra por uno o varios Números de Libramiento (separados por coma, espacio o ;)."""
    tokens = {norm(t) for t in re.split(r"[,\s;]+", text.strip()) if t}
    if not tokens:
        return df.iloc[0:0]
    return df[df[COL_LIBRAMIENTO].map(lambda v: norm(v) in tokens)]
