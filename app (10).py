import html
import io
import json
import re
import unicodedata
from io import BytesIO

import numpy as np
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter

# ======================================================
# FUNCIÓN: BLOQUEAR LETRAS (SOLO NÚMEROS)
# ======================================================
def solo_numeros(key):
    valor = st.session_state.get(key, "")
    st.session_state[key] = re.sub(r"\D", "", valor)

# ======================================================
# CONFIGURACIÓN
# ======================================================
st.set_page_config(
    page_title="DRCC DATA UNIFY",
    page_icon="📊",
    layout="wide"
)

# ======================================================
# ESTILOS
# ======================================================
st.markdown("""
<style>
.main-title { color:#1E3A8A; font-size:42px; font-weight:bold; margin-bottom:0; }
.sub-title { color:#333; font-size:20px; font-weight:600; margin-top:5px; }
</style>
""", unsafe_allow_html=True)

# ======================================================
# ENCABEZADO
# ======================================================
st.markdown('<p class="main-title">DRCC DATA UNIFY</p>', unsafe_allow_html=True)
st.markdown('<p class="sub-title">Creado por Juan Adonai Brito, | Idea: Chabellys Encarnacion</p>', unsafe_allow_html=True)
st.markdown(
    '<p style="color:#555; font-size:16px;">'
    'Ahorra tiempo al unificar estructuras programáticas y libramientos en SIGEF.'
    '</p>',
    unsafe_allow_html=True
)
st.divider()

# ======================================================
# BASE DE DATOS: NUEVOS FUNCIONARIOS DESIGNADOS
# (Decretos 551-26 al 558-26)
# ======================================================
FUNCIONARIOS_NUEVOS = [
    {
        "aliases": ["INTRANT", "INSTITUTO NACIONAL DE TRANSITO Y TRANSPORTE TERRESTRE"],
        "funcionarios": [
            {"nombre": "Juan Manuel Méndez García", "cargo": "Director Ejecutivo", "decreto": "551-26"},
        ],
    },
    {
        "aliases": ["COE", "CENTRO DE OPERACIONES DE EMERGENCIA"],
        "funcionarios": [
            {"nombre": "Erdwin Robert Olivares Luciano", "cargo": "Director", "decreto": "551-26"},
        ],
    },
    {
        "aliases": ["DEFENSA CIVIL"],
        "funcionarios": [
            {"nombre": "Carlos Manuel Paulino Cárdenas", "cargo": "Director Ejecutivo", "decreto": "551-26"},
        ],
    },
    {
        "aliases": ["JAC", "JUNTA DE AVIACION CIVIL"],
        "funcionarios": [
            {"nombre": "Julio Peña Guzmán", "cargo": "Presidente", "decreto": "552-26"},
        ],
    },
    {
        "aliases": ["DEPARTAMENTO AEROPORTUARIO"],
        "funcionarios": [
            {"nombre": "Mérido de Jesús Torres Espinal", "cargo": "Director Ejecutivo", "decreto": "552-26"},
        ],
    },
    {
        "aliases": ["MERCADOM", "MERCADOS DOMINICANOS DE ABASTO AGROPECUARIO"],
        "funcionarios": [
            {"nombre": "Mariana Tavarez de Santos", "cargo": "Directora General", "decreto": "552-26"},
        ],
    },
    {
        "aliases": ["INAIPI", "INSTITUTO NACIONAL DE ATENCION INTEGRAL A LA PRIMERA INFANCIA"],
        "funcionarios": [
            {"nombre": "Kenia Xiomara Guante Valdez", "cargo": "Directora", "decreto": "553-26"},
        ],
    },
    {
        "aliases": ["INSTITUTO NACIONAL DE MIGRACION"],
        "funcionarios": [
            {"nombre": "José Benedicto Hernández Tejada", "cargo": "Director Ejecutivo", "decreto": "553-26"},
        ],
    },
    {
        "aliases": ["MIREX", "MINISTERIO DE RELACIONES EXTERIORES", "RELACIONES EXTERIORES"],
        "funcionarios": [
            {"nombre": "Víctor O. Disonó Haza", "cargo": "Ministro", "decreto": "554-26"},
        ],
    },
    {
        "aliases": ["MINISTERIO DE VIVIENDA", "MINISTERIO DE LA VIVIENDA, HABITAT Y EDIFICACIONES (MIVHED)", "MIVHED"],
        "funcionarios": [
            {"nombre": "Jean Luis Rodríguez", "cargo": "Ministro", "decreto": "554-26"},
        ],
    },
    {
        "aliases": ["AUTORIDAD PORTUARIA DOMINICANA", "AUTORIDAD PORTUARIA", "APORDOM"],
        "funcionarios": [
            {"nombre": "Francisco Alejandro Campos Alvarez", "cargo": "Director Ejecutivo", "decreto": "554-26"},
            {"nombre": "Julia Luz Marina Múñiz Suberví", "cargo": "Presidente Consejo de Administración", "decreto": "554-26"},
        ],
    },
    {
        "aliases": ["DIRECCION DE DESARROLLO PROVINCIAL"],
        "funcionarios": [
            {"nombre": "José Dolores Andújar Ramírez", "cargo": "Director", "decreto": "555-26"},
        ],
    },
    {
        "aliases": ["UTEPDA", "UNIDAD TECNICA EJECUTORA DE PROYECTOS DE DESARROLLO AGROFORESTAL"],
        "funcionarios": [
            {"nombre": "Nidio Encarnación Santiago", "cargo": "Director Ejecutivo", "decreto": "555-26"},
        ],
    },
    {
        "aliases": ["CORAASAN", "CORPORACION DEL ACUEDUCTO Y ALCANTARILLADO DE SANTIAGO"],
        "funcionarios": [
            {"nombre": "César Andrés Pichardo Fermín", "cargo": "Presidente", "decreto": "555-26"},
            {"nombre": "Bernardo Antonio Inoa Pichardo", "cargo": "Miembro Consejo de Directores", "decreto": "555-26"},
        ],
    },
    {
        "aliases": ["CUED", "CONSEJO UNIFICADO DE LAS EMPRESAS DISTRIBUIDORAS DE ELECTRICIDAD"],
        "funcionarios": [
            {"nombre": "Joel Adrián Santos Echavarría", "cargo": "Presidente / Coordinador Gabinete de Energía", "decreto": "556-26"},
        ],
    },
    {
        "aliases": ["FUERZA AEREA", "FARD", "FUERZA AEREA DE LA REPUBLICA DOMINICANA"],
        "funcionarios": [
            {"nombre": "Enmanuel Marcelino Souffront Tamayo", "cargo": "Comandante General", "decreto": "557-26"},
            {"nombre": "Manuel José Brito Estepan", "cargo": "Subcomandante General", "decreto": "557-26"},
            {"nombre": "Dionisio De La Rosa Hernández", "cargo": "Inspector General", "decreto": "557-26"},
        ],
    },
    {
        "aliases": [ "FUERZAS ARMADAS" , "MINISTERIO DE DEFENSA"],
        "funcionarios": [
            {"nombre": "Jorge Iván Camino Pérez", "cargo": "Inspector General de las Fuerzas Armadas", "decreto": "557-26"},
            {"nombre": "Delio Buenaventura Colón Rosario", "cargo": "Viceministro de Defensa", "decreto": "557-26"},
        ],
    },
    {
        "aliases": ["EJERCITO DE LA REPUBLICA DOMINICANA", "EJERCITO", "ERD"],
        "funcionarios": [
            {"nombre": "Jimmy Arias Grullón", "cargo": "Comandante General", "decreto": "557-26"},
            {"nombre": "José Manuel Duran Infante", "cargo": "Subcomandante General", "decreto": "557-26"},
            {"nombre": "Raúl Esteban Mora Hernández", "cargo": "Inspector General", "decreto": "557-26"},
            {"nombre": "Rafael Raimundo Ramírez Tejeda", "cargo": "Comandante Regimiento Guardia Presidencial", "decreto": "557-26"},
        ],
    },
    {
        "aliases": ["CUSEP", "CUERPO DE SEGURIDAD PRESIDENCIAL"],
        "funcionarios": [
            {"nombre": "Guillermo Caro Cruz", "cargo": "Jefe", "decreto": "557-26"},
        ],
    },
    {
        "aliases": ["POLICIA NACIONAL"],
        "funcionarios": [
            {"nombre": "Ernesto Rafael Rodríguez García", "cargo": "Director General", "decreto": "558-26"},
            {"nombre": "Martín Miguel Tapia Sánchez", "cargo": "Inspector General", "decreto": "574-26"},
        ],
    },
    {
        "aliases": ["SUPERINTENDENCIA DE BANCOS"],
        "funcionarios": [
            {"nombre": "Víctor Livio Enmanuel Cedeño Brea", "cargo": "Superintendente de Bancos", "decreto": "606-26"},
        ],
    },
    {
        "aliases": ["HOSPITAL MATERNO DR. REYNALDO ALMANZAR"],
        "funcionarios": [
            {"nombre": "Dr. Jorge Vilorio", "cargo": "Director Ejecutivo", "decreto": "1"},
        ],
    },
    {
        "aliases": ["HOSPITAL MUNICIPAL LILIAN FERNÁNDEZ", "LILIAN"],
        "funcionarios": [
            {"nombre": "Juan Carlos Gómez", "cargo": "Director Ejecutivo", "decreto": "1"},
        ],
    },
    {
        "aliases": ["HOSPITAL MUNICIPAL DR. JORGE A. MARTINEZ ", "ICO", "JORGE ARMANDO", "HOSPITAL MUNICIPAL DE TAMBORIL", "JORGE ARMANDO MARTINEZ"],
        "funcionarios": [
            {"nombre": "José Luis Gómez Rodríguez", "cargo": "Director Ejecutivo", "decreto": "1"},
        ],
    },
]

# ======================================================
# HELPERS: NORMALIZACIÓN Y DETECCIÓN DE FUNCIONARIOS
# ======================================================
def normalizar(texto):
    if texto is None:
        return ""
    texto = str(texto).upper().strip()
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    texto = re.sub(r"[^A-Z0-9\s]", " ", texto)
    texto = re.sub(r"\s+", " ", texto).strip()
    return texto

def buscar_funcionario_nuevo(texto_institucion):
    """Devuelve (nombres_funcionarios, decretos) si la institución coincide, si no (None, None)."""
    texto_norm = normalizar(texto_institucion)
    if not texto_norm:
        return None, None

    for entrada in FUNCIONARIOS_NUEVOS:
        for alias in entrada["aliases"]:
            alias_norm = normalizar(alias)
            patron = r"\b" + re.escape(alias_norm) + r"\b"
            if re.search(patron, texto_norm) or re.search(
                r"\b" + re.escape(texto_norm) + r"\b", alias_norm
            ):
                nombres = "; ".join(
                    f"{f['nombre']} ({f['cargo']})" for f in entrada["funcionarios"]
                )
                decretos = ", ".join(
                    sorted(set(f["decreto"] for f in entrada["funcionarios"]))
                )
                return nombres, decretos
    return None, None

def detectar_columna(cols, claves):
    for col in cols:
        col_texto = str(col).strip().lower()
        if any(k in col_texto for k in claves):
            return col
    return None

# ============================================================================
# EXTRACTOR DE CÓDIGOS REVISADOS — LÓGICA
# ============================================================================
COL_HOJA = "Hoja"
COL_CODIGO = "Código"
COL_LIBRAMIENTO = "Número Libramiento"
COL_DOC = "Documento"

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


def _col_index(ref: str) -> int:
    """'AB12' -> 27 (índice de columna base 0)."""
    n = 0
    for ch in ref:
        if not ch.isalpha():
            break
        n = n * 26 + (ord(ch.upper()) - 64)
    return n - 1


def _read_xlsx_stdlib(data: bytes) -> dict:
    """Lector de .xlsx sin dependencias (zipfile + xml). Devuelve {hoja: DataFrame}."""
    import posixpath
    import xml.etree.ElementTree as ET
    import zipfile

    ns = {
        "m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
        "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        "pr": "http://schemas.openxmlformats.org/package/2006/relationships",
    }
    rid_attr = "{%s}id" % ns["r"]

    def text_of(el) -> str:
        return "".join(t.text or "" for t in el.iter("{%s}t" % ns["m"]))

    with zipfile.ZipFile(io.BytesIO(data)) as z:
        names = set(z.namelist())

        shared = []
        if "xl/sharedStrings.xml" in names:
            root = ET.fromstring(z.read("xl/sharedStrings.xml"))
            shared = [text_of(si) for si in root.findall("m:si", ns)]

        rels = {}
        if "xl/_rels/workbook.xml.rels" in names:
            for rel in ET.fromstring(z.read("xl/_rels/workbook.xml.rels")).findall("pr:Relationship", ns):
                target = rel.get("Target", "")
                rels[rel.get("Id")] = (
                    target.lstrip("/") if target.startswith("/") else posixpath.normpath("xl/" + target)
                )

        wb = ET.fromstring(z.read("xl/workbook.xml"))
        out = {}
        for i, sh in enumerate(wb.findall("m:sheets/m:sheet", ns), start=1):
            name = sh.get("name") or f"Hoja{i}"
            path = rels.get(sh.get(rid_attr)) or f"xl/worksheets/sheet{i}.xml"
            if path not in names:
                continue
            rows = {}
            for row in ET.fromstring(z.read(path)).iter("{%s}row" % ns["m"]):
                for c in row.findall("m:c", ns):
                    ref, t = c.get("r", ""), c.get("t")
                    v = c.find("m:v", ns)
                    if t == "inlineStr":
                        val = text_of(c)
                    elif v is None or v.text is None:
                        continue
                    elif t == "s":
                        val = shared[int(v.text)]
                    elif t == "b":
                        val = v.text.strip() == "1"
                    elif t in ("str", "e"):
                        val = v.text
                    else:
                        try:
                            val = float(v.text)
                        except ValueError:
                            val = v.text
                    if ref:
                        r_idx = int("".join(ch for ch in ref if ch.isdigit())) - 1
                        rows.setdefault(r_idx, {})[_col_index(ref)] = val
            if not rows:
                out[name] = pd.DataFrame()
                continue
            n_rows, n_cols = max(rows) + 1, max(max(r) for r in rows.values()) + 1
            grid = [[None] * n_cols for _ in range(n_rows)]
            for r_idx, cols in rows.items():
                for c_idx, val in cols.items():
                    grid[r_idx][c_idx] = val
            out[name] = pd.DataFrame(grid, dtype=object)
        return out


_OLE_SIG = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"


def _ole_stream(data: bytes, wanted=("workbook", "book")) -> bytes:
    """Extrae el stream 'Workbook' de un contenedor OLE2 (.xls) sin librerías externas."""
    import struct

    if data[:8] != _OLE_SIG:
        raise ValueError("El archivo no es un .xls válido.")
    ssz = 1 << struct.unpack_from("<H", data, 0x1E)[0]
    mssz = 1 << struct.unpack_from("<H", data, 0x20)[0]
    n_fat, dir_start = struct.unpack_from("<II", data, 0x2C)
    mini_cutoff = struct.unpack_from("<I", data, 0x38)[0]
    minifat_start, _n_minifat, difat_start, n_difat = struct.unpack_from("<IIII", data, 0x3C)
    FREE = 0xFFFFFFFA  # valores >= FREE son marcas especiales (fin de cadena, libre, FAT...)

    def sector(i):
        off = (i + 1) * ssz
        return data[off : off + ssz]

    def words(buf):
        return struct.unpack("<%dI" % (len(buf) // 4), buf[: len(buf) // 4 * 4])

    difat = list(struct.unpack_from("<109I", data, 0x4C))
    s = difat_start
    for _ in range(n_difat):
        if s >= FREE:
            break
        w = words(sector(s))
        difat.extend(w[:-1])
        s = w[-1]
    fat = []
    for fs in [x for x in difat if x < FREE][:n_fat]:
        fat.extend(words(sector(fs)))

    def chain(start, table):
        out, guard = [], 0
        while start < FREE and start < len(table) and guard <= len(table):
            out.append(start)
            start = table[start]
            guard += 1
        return out

    def read_chain(start):
        return b"".join(sector(i) for i in chain(start, fat))

    dir_data = read_chain(dir_start)
    entries = []
    for off in range(0, len(dir_data) - 127, 128):
        e = dir_data[off : off + 128]
        nlen = struct.unpack_from("<H", e, 0x40)[0]
        name = e[: max(nlen - 2, 0)].decode("utf-16-le", "ignore")
        start = struct.unpack_from("<I", e, 0x74)[0]
        size = struct.unpack_from("<I", e, 0x78)[0]
        entries.append((name, e[0x42], start, size))

    root = next((e for e in entries if e[1] == 5), None)
    target = next((e for e in entries if e[1] == 2 and e[0].lower() in wanted), None)
    if target is None:
        raise ValueError("El archivo no contiene un libro de Excel (.xls) legible.")
    _, _, start, size = target
    if size >= mini_cutoff or root is None:
        return read_chain(start)[:size]

    minifat = []
    for s in chain(minifat_start, fat):
        minifat.extend(words(sector(s)))
    container = read_chain(root[2])
    parts = [container[i * mssz : (i + 1) * mssz] for i in chain(start, minifat)]
    return b"".join(parts)[:size]


class _Cur:
    """Cursor sobre varios bloques (registro + CONTINUE) para leer cadenas BIFF8."""

    def __init__(self, chunks, pos=0):
        self.chunks, self.ci, self.pos = chunks, 0, pos

    def avail(self):
        return len(self.chunks[self.ci]) - self.pos

    def advance(self):
        self.ci += 1
        self.pos = 0
        if self.ci >= len(self.chunks):
            raise EOFError

    def take(self, n):
        out = b""
        while n > 0:
            if self.avail() <= 0:
                self.advance()
            k = min(n, self.avail())
            out += self.chunks[self.ci][self.pos : self.pos + k]
            self.pos += k
            n -= k
        return out


def _biff8_string(cur: _Cur, long_len: bool = True) -> str:
    import struct

    if long_len:
        cch, flags = struct.unpack("<HB", cur.take(3))
    else:
        cch, flags = struct.unpack("<BB", cur.take(2))
    runs = struct.unpack("<H", cur.take(2))[0] if flags & 8 else 0
    ext = struct.unpack("<I", cur.take(4))[0] if flags & 4 else 0
    is16, need, parts = flags & 1, cch, []
    while need > 0:
        width = 2 if is16 else 1
        if cur.avail() < width:  # la cadena continúa en un bloque CONTINUE (con byte de opciones)
            cur.advance()
            is16 = cur.chunks[cur.ci][0] & 1
            cur.pos = 1
            width = 2 if is16 else 1
        k = min(need, cur.avail() // width)
        raw = cur.chunks[cur.ci][cur.pos : cur.pos + k * width]
        cur.pos += k * width
        parts.append(raw.decode("utf-16-le", "replace") if is16 else raw.decode("latin-1"))
        need -= k
    cur.take(runs * 4 + ext)
    return "".join(parts)


def _rk_value(v: int) -> float:
    import struct

    if v & 2:
        n = ((v - (1 << 32)) if v & 0x80000000 else v) >> 2
        n = float(n)
    else:
        n = struct.unpack("<d", struct.pack("<Q", (v & 0xFFFFFFFC) << 32))[0]
    return n / 100 if v & 1 else n


def _read_xls_stdlib(data: bytes) -> dict:
    """Lector de .xls (BIFF5/BIFF8) sin dependencias. Devuelve {hoja: DataFrame}."""
    import struct

    buf = _ole_stream(data)
    n = len(buf)

    def records(pos):
        while pos + 4 <= n:
            rid, ln = struct.unpack_from("<HH", buf, pos)
            yield rid, buf[pos + 4 : pos + 4 + ln]
            pos += 4 + ln

    # ---- Sección global: versión, hojas y tabla de cadenas compartidas
    biff8, sheets, sst, recs = True, [], [], []
    depth = 0
    for rid, p in records(0):
        if rid == 0x0809:
            depth += 1
            if depth == 1 and len(p) >= 2:
                biff8 = struct.unpack_from("<H", p, 0)[0] >= 0x0600
        elif rid == 0x000A:
            depth -= 1
            if depth <= 0:
                break
        elif rid == 0x002F:
            raise ValueError("El archivo .xls está protegido con contraseña.")
        recs.append((rid, p))

    i = 0
    while i < len(recs):
        rid, p = recs[i]
        if rid == 0x0085 and len(p) >= 8:
            off, _vis, typ = struct.unpack_from("<IBB", p, 0)
            if biff8:
                name = _biff8_string(_Cur([p], 6), long_len=False)
            else:
                name = p[7 : 7 + p[6]].decode("cp1252", "replace")
            if typ == 0:
                sheets.append((name, off))
        elif rid == 0x00FC:
            chunks = [p]
            while i + 1 < len(recs) and recs[i + 1][0] == 0x003C:
                i += 1
                chunks.append(recs[i][1])
            try:
                cur = _Cur(chunks, 8)
                for _ in range(struct.unpack_from("<I", p, 4)[0]):
                    sst.append(_biff8_string(cur))
            except (EOFError, struct.error):
                pass
        i += 1

    def text(p, start):
        if biff8:
            return _biff8_string(_Cur([p], start))
        ln = struct.unpack_from("<H", p, start)[0]
        return p[start + 2 : start + 2 + ln].decode("cp1252", "replace")

    # ---- Hojas
    out = {}
    for sname, soff in sheets:
        cells, pending, depth = {}, None, 0
        for rid, p in records(soff):
            try:
                if rid == 0x0809:
                    depth += 1
                elif rid == 0x000A:
                    depth -= 1
                    if depth <= 0:
                        break
                elif depth != 1:
                    continue
                elif rid == 0x00FD:  # LABELSST
                    r, c, _x, k = struct.unpack_from("<HHHI", p, 0)
                    if k < len(sst):
                        cells[(r, c)] = sst[k]
                elif rid == 0x0204:  # LABEL
                    r, c = struct.unpack_from("<HH", p, 0)
                    cells[(r, c)] = text(p, 6)
                elif rid == 0x0203:  # NUMBER
                    r, c, _x, v = struct.unpack_from("<HHHd", p, 0)
                    cells[(r, c)] = v
                elif rid == 0x027E:  # RK
                    r, c, _x, v = struct.unpack_from("<HHHI", p, 0)
                    cells[(r, c)] = _rk_value(v)
                elif rid == 0x00BD:  # MULRK
                    r, c1 = struct.unpack_from("<HH", p, 0)
                    for k in range((len(p) - 6) // 6):
                        _x, v = struct.unpack_from("<HI", p, 4 + 6 * k)
                        cells[(r, c1 + k)] = _rk_value(v)
                elif rid == 0x0205:  # BOOLERR
                    r, c, _x, v, is_err = struct.unpack_from("<HHHBB", p, 0)
                    if not is_err:
                        cells[(r, c)] = bool(v)
                elif rid == 0x0006:  # FORMULA
                    r, c = struct.unpack_from("<HH", p, 0)
                    res = p[6:14]
                    if res[6:8] == b"\xff\xff":
                        if res[0] == 0:
                            pending = (r, c)
                        elif res[0] == 1:
                            cells[(r, c)] = bool(res[2])
                    else:
                        cells[(r, c)] = struct.unpack("<d", res)[0]
                elif rid == 0x0207 and pending is not None:  # STRING (resultado de fórmula)
                    cells[pending] = text(p, 0)
                    pending = None
            except (struct.error, EOFError, IndexError):
                continue
        if not cells:
            out[sname] = pd.DataFrame()
            continue
        n_rows = max(r for r, _ in cells) + 1
        n_cols = max(c for _, c in cells) + 1
        grid = [[None] * n_cols for _ in range(n_rows)]
        for (r, c), v in cells.items():
            grid[r][c] = v
        out[sname] = pd.DataFrame(grid, dtype=object)
    if not out:
        raise ValueError("El archivo .xls no tiene hojas de cálculo legibles.")
    return out


def _decode_text(data: bytes) -> str:
    for enc in ("utf-8-sig", "utf-16", "cp1252"):
        try:
            return data.decode(enc)
        except UnicodeError:
            continue
    return data.decode("latin-1", "replace")


def _read_html_tables(data: bytes) -> dict:
    """Muchos sistemas exportan HTML con extensión .xls: se leen sus tablas."""
    from html.parser import HTMLParser

    class _P(HTMLParser):
        def __init__(self):
            super().__init__(convert_charrefs=True)
            self.tables, self.stack, self.row, self.cell = [], [], None, None

        def handle_starttag(self, tag, attrs):
            if tag == "table":
                self.stack.append([])
            elif tag == "tr" and self.stack:
                self.row = []
            elif tag in ("td", "th") and self.row is not None:
                self.cell = []
            elif tag == "br" and self.cell is not None:
                self.cell.append(" ")

        def handle_endtag(self, tag):
            if tag in ("td", "th") and self.cell is not None and self.row is not None:
                self.row.append(" ".join("".join(self.cell).split()))
                self.cell = None
            elif tag == "tr" and self.row is not None and self.stack:
                self.stack[-1].append(self.row)
                self.row = None
            elif tag == "table" and self.stack:
                self.tables.append(self.stack.pop())

        def handle_data(self, d):
            if self.cell is not None:
                self.cell.append(d)

    p = _P()
    p.feed(_decode_text(data))
    out = {}
    for k, rows in enumerate((t for t in p.tables if t), start=1):
        width = max(len(r) for r in rows)
        out[f"Tabla{k}"] = pd.DataFrame([r + [None] * (width - len(r)) for r in rows], dtype=object)
    return out


def _read_delimited(data: bytes) -> dict:
    import csv

    txt = _decode_text(data)
    first = txt.splitlines()[0] if txt.strip() else ""
    delim = max("\t;,|", key=first.count)
    rows = list(csv.reader(io.StringIO(txt), delimiter=delim))
    if not rows:
        return {}
    width = max(len(r) for r in rows)
    return {"Hoja1": pd.DataFrame([r + [None] * (width - len(r)) for r in rows], dtype=object)}


def _read_all_sheets(data: bytes, filename: str):
    head = data[:8]

    # .xlsx (ZIP), aunque el nombre diga .xls
    if head[:2] == b"PK":
        try:
            return pd.read_excel(io.BytesIO(data), sheet_name=None, header=None, engine="openpyxl", dtype=object)
        except Exception:  # noqa: BLE001 - sin openpyxl o archivo raro: lector propio
            pass
        try:
            return _read_xlsx_stdlib(data)
        except Exception as exc:  # noqa: BLE001
            raise ValueError(f"No se pudo leer el archivo Excel: {exc}")

    # .xls clásico (OLE2)
    if head == _OLE_SIG:
        try:
            return pd.read_excel(io.BytesIO(data), sheet_name=None, header=None, engine="xlrd", dtype=object)
        except Exception:  # noqa: BLE001 - sin xlrd o archivo raro: lector propio
            pass
        try:
            return _read_xls_stdlib(data)
        except Exception as exc:  # noqa: BLE001
            raise ValueError(f"No se pudo leer el archivo .xls: {exc}")

    # .xls que en realidad es HTML o texto delimitado (exportaciones de sistemas)
    sheets = _read_html_tables(data) if b"<" in data[:2000] else {}
    if not sheets:
        sheets = _read_delimited(data)
    if not sheets:
        raise ValueError("No se reconoce el formato del archivo (se esperaba .xlsx o .xls).")
    return sheets


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


def split_new_records(df_new: pd.DataFrame, refs: list):
    """Separa los registros de df_new que NO están (nuevos) / SÍ están (repetidos) en los documentos de referencia.

    Dos registros son el mismo cuando comparten el mismo Código (sin importar mayúsculas, acentos ni espacios).
    """
    ref_codes = {norm(c) for ref in refs for c in ref[COL_CODIGO]}
    repetido = df_new[COL_CODIGO].map(lambda c: norm(c) in ref_codes)
    return df_new[~repetido], df_new[repetido]


@st.cache_data(show_spinner="Leyendo Excel…")
def cargar_archivo(data: bytes, fname: str):
    return extract_reviewed(data, fname)


def copy_list_component(records: list):
    """Un registro debajo del otro, cada uno con su botón de copiado.

    records = [(código, detalle, texto_a_copiar)]
    """
    rows = "".join(
        f'<div class="row"><div class="info"><span class="code">{html.escape(code)}</span>'
        + (f'<span class="lib">{html.escape(detail)}</span>' if detail else "")
        + f'</div><button class="btn" data-v="{html.escape(copy_text, quote=True)}">Copiar</button></div>'
        for code, detail, copy_text in records
    )
    all_text = json.dumps("\n".join(ct for _, _, ct in records)).replace("<", "\\u003c")
    page = f"""
    <style>
      body {{ margin:0; font-family: "Source Sans Pro", system-ui, sans-serif; }}
      .top {{ display:flex; justify-content:flex-end; margin-bottom:8px; }}
      .list {{ max-height:440px; overflow-y:auto; padding-right:4px; }}
      .row {{ display:flex; align-items:center; justify-content:space-between; gap:12px;
              background:#f6f8fa; color:#1f2328; border:1px solid #d0d7de; border-radius:8px;
              padding:8px 12px; margin-bottom:6px; }}
      .info {{ display:flex; flex-direction:column; min-width:0; }}
      .code {{ font-family: ui-monospace, Menlo, Consolas, monospace; font-weight:600; word-break:break-all; }}
      .lib {{ font-size:12px; color:#656d76; word-break:break-word; }}
      .btn {{ cursor:pointer; border:1px solid #d0d7de; background:#fff; color:#1f2328; border-radius:6px;
              padding:5px 12px; font-size:13px; white-space:nowrap; }}
      .btn:hover {{ border-color:#ff4b4b; color:#ff4b4b; }}
      .btn.ok {{ background:#1a7f37; border-color:#1a7f37; color:#fff; }}
    </style>
    <div class="top"><button class="btn" id="copyall">Copiar todos</button></div>
    <div class="list">{rows}</div>
    <script>
      function copyText(text, btn) {{
        const done = () => {{
          const old = btn.textContent; btn.textContent = '¡Copiado!'; btn.classList.add('ok');
          setTimeout(() => {{ btn.textContent = old; btn.classList.remove('ok'); }}, 1200);
        }};
        const fallback = () => {{
          const t = document.createElement('textarea'); t.value = text;
          t.style.position = 'fixed'; t.style.opacity = '0'; document.body.appendChild(t);
          t.select(); try {{ document.execCommand('copy'); done(); }} catch (e) {{}}
          document.body.removeChild(t);
        }};
        if (navigator.clipboard && window.isSecureContext) {{
          navigator.clipboard.writeText(text).then(done).catch(fallback);
        }} else {{ fallback(); }}
      }}
      document.querySelectorAll('button[data-v]').forEach(b =>
        b.addEventListener('click', () => copyText(b.dataset.v, b)));
      document.getElementById('copyall').addEventListener('click', e => copyText({all_text}, e.target));
    </script>
    """
    height = min(len(records) * 62 + 60, 520)
    if hasattr(st, "iframe"):  # Streamlit reciente (components.html está en desuso)
        st.iframe(page, height=height)
    else:
        components.html(page, height=height, scrolling=False)


def render_extractor():
    """Interfaz del modo 'Extractor de códigos revisados'.

    Usa return (no st.stop) para que el resto de la app, como el pie de página, siga mostrándose.
    """
    st.subheader("✅ Extractor de códigos revisados")
    st.caption(
        "Sube uno o varios Excel (.xlsx / .xls). Se busca la columna **Revisada**, se toman las filas con "
        "**Verdadero** y se extraen sus **Códigos**."
    )

    uploaded = st.file_uploader(
        "Archivos Excel",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
        key="ext_files",
        help="El orden en que los subas define el Documento 1, 2, 3… (el 1.º es el más antiguo).",
    )
    if not uploaded:
        return

    # ---- Lectura de cada documento (numerados en el orden de carga)
    docs, has_lib_any = [], False
    for i, f in enumerate(uploaded, start=1):
        label = f"{i}. {f.name}"
        try:
            d, has_lib, notes = cargar_archivo(f.getvalue(), f.name)
        except Exception as exc:  # noqa: BLE001
            st.error(f"«{f.name}»: {exc}")
            continue
        for n in notes:
            st.info(f"{f.name}: {n}")
        has_lib_any = has_lib_any or has_lib
        docs.append((label, d.assign(**{COL_DOC: label})))

    if not docs:
        return

    multi = len(docs) > 1
    st.metric("Registros con Revisada = Verdadero", sum(len(d) for _, d in docs))
    if multi:
        for label, d in docs:
            st.caption(f"**{label}** — {len(d)} registros")
    st.divider()

    # ---- Registros nuevos respecto a otros documentos
    solo_nuevos = st.checkbox(
        "Identificar registros nuevos (exportar solo los que no se repiten del documento anterior)",
        key="ext_solo_nuevos",
        disabled=not multi,
        help="Compara por Código. Los registros que ya estaban en el documento de referencia se excluyen.",
    )
    if not multi:
        st.caption("Sube al menos 2 documentos para habilitar esta opción.")

    result = pd.concat([d for _, d in docs], ignore_index=True)

    if solo_nuevos:
        labels = [label for label, _ in docs]
        by_label = dict(docs)
        nuevo_label = st.selectbox(
            "Documento nuevo (el que se revisa)", labels, index=len(labels) - 1, key="ext_nuevo"
        )
        otros = [l for l in labels if l != nuevo_label]
        ref_labels = st.multiselect(
            "Comparar contra (documentos anteriores)", otros, default=otros, key="ext_refs"
        )
        if not ref_labels:
            st.warning("Selecciona al menos un documento de referencia.")
            return
        nuevos, repetidos = split_new_records(by_label[nuevo_label], [by_label[l] for l in ref_labels])
        m1, m2 = st.columns(2)
        m1.metric("Nuevos", len(nuevos))
        m2.metric("Repetidos (ya existían)", len(repetidos))
        result = nuevos
        if result.empty:
            st.warning("No hay registros nuevos: todos los códigos del documento ya estaban en los de referencia.")
            return

    st.divider()

    # ---- Filtro por Número Libramiento
    usar_filtro = st.checkbox("Aplicar Filtro de búsqueda para exportar.", key="ext_usar_filtro")
    if usar_filtro:
        if not has_lib_any:
            st.error("Los archivos no tienen una columna de «Libramiento» para filtrar.")
            return
        libramiento_filtro = st.text_input(
            "Número Libramiento(*)",
            placeholder="Ej.: 1234 (puedes escribir varios separados por coma)",
            key="ext_libramiento",
        )
        if not libramiento_filtro.strip():
            st.warning("Escribe el Número Libramiento(*) para exportar los códigos relacionados.")
            return
        result = filter_by_libramiento(result, libramiento_filtro)

    if result.empty:
        st.warning("No hay códigos para exportar con los criterios indicados.")
        return

    def _detalle(row) -> str:
        partes = []
        if row[COL_LIBRAMIENTO]:
            partes.append(f"Libramiento: {row[COL_LIBRAMIENTO]}")
        if multi:
            partes.append(f"Documento: {row[COL_DOC]}")
        return " · ".join(partes)

    def _texto_copiar(row) -> str:
        """Código de trámite + Número de Libramiento (separados por tabulador: al pegar en Excel quedan en 2 columnas)."""
        return f"{row[COL_CODIGO]}\t{row[COL_LIBRAMIENTO]}" if row[COL_LIBRAMIENTO] else row[COL_CODIGO]

    records = [(row[COL_CODIGO], _detalle(row), _texto_copiar(row)) for _, row in result.iterrows()]

    st.subheader(f"Exportación ({len(records)} código{'s' if len(records) != 1 else ''})")
    copy_list_component(records)

    d1, d2 = st.columns(2)
    d1.download_button(
        "⬇️ Descargar TXT",
        "\n".join(result[COL_CODIGO]),
        file_name="codigos.txt",
        mime="text/plain",
        use_container_width=True,
    )
    d2.download_button(
        "⬇️ Descargar CSV",
        result[[COL_DOC, COL_HOJA, COL_CODIGO, COL_LIBRAMIENTO]].to_csv(index=False).encode("utf-8-sig"),
        file_name="codigos.csv",
        mime="text/csv",
        use_container_width=True,
    )

# ======================================================
# SELECCIÓN DE MODO
# ======================================================
modo = st.radio(
    "🧭 Selecciona el modo de trabajo",
    [
        "🔁 Modo múltiple (Excel)",
        "🧩 Modo manual (uno por uno)",
        "✅ Extractor de códigos revisados",
    ],
    horizontal=True
)

st.divider()

# ======================================================
# MODO MÚLTIPLE
# ======================================================
if modo.startswith("🔁"):

    col1, col2 = st.columns([1, 2], gap="large")

    # Inicializar variables
    df = None
    override = False

    with col1:
        st.info("### 📂 Cargar archivo Excel")
        uploaded_file = st.file_uploader("Subir archivo (.xlsx / .xls)", type=["xlsx", "xls"])

        if uploaded_file:
            try:
                # Leer archivo completo sin asumir encabezado
                uploaded_file.seek(0)
                df_raw = pd.read_excel(uploaded_file, header=None, dtype=str).fillna("")

                # Función más precisa para detectar encabezado real
                def detectar_header(df_temp):
                    for i in range(min(6, len(df_temp))):
                        fila = df_temp.iloc[i].astype(str).str.lower()
                        if any("estructura" in c for c in fila) and any("libramiento" in c for c in fila):
                            return i
                    return None

                header_row = detectar_header(df_raw)

                if header_row is not None:
                    uploaded_file.seek(0)
                    df = pd.read_excel(uploaded_file, header=header_row, dtype=str).fillna("")
                else:
                    df = df_raw.copy()
                    df.columns = [f"Columna_{i+1}" for i in range(len(df.columns))]

                st.success("✅ Archivo cargado correctamente")

                override = st.checkbox("✏️ Manual")

            except Exception as e:
                st.error(f"Error al leer el archivo: {e}")

    with col2:
        if df is None:
            st.warning("Esperando archivo para procesar...")
        else:
            try:
                columnas = list(df.columns)
                col_estructura = None
                col_libramiento = None

                # Lógica: Si es manual (override) mostramos tabla y selectores.
                # Si es automático, NO mostramos tabla, solo procesamos.
                if override:
                    st.info("Modo manual activado.")
                    st.subheader("👀 Vista previa de los datos")
                    st.dataframe(df.head(20), use_container_width=True)

                    col_estructura = st.selectbox(
                        "Selecciona la columna de Estructura Programática",
                        columnas
                    )
                    col_libramiento = st.selectbox(
                        "Selecciona la columna de Número de Libramiento",
                        columnas
                    )
                else:
                    # Detección automática (sin vista previa)
                    col_estructura = detectar_columna(df.columns, ["estructura", "programática", "programatica"])
                    col_libramiento = detectar_columna(df.columns, ["libramiento", "número", "numero"])

                # Procesamiento
                if not col_estructura or not col_libramiento:
                    st.error("❌ No se pudieron identificar las columnas automáticamente. Activa la casilla manual.")
                else:
                    # Si estamos en automático, mostramos qué columnas eligió el sistema
                    if not override:
                        st.caption(f"✅ Columnas detectadas: **{col_estructura}** y **{col_libramiento}**")

                    def transformar(fila):
                        v1 = str(fila[col_estructura]) if pd.notna(fila[col_estructura]) else ""
                        v2 = str(fila[col_libramiento]) if pd.notna(fila[col_libramiento]) else ""

                        v1 = v1.strip().split('.')[0]
                        v2 = v2.strip().split('.')[0]

                        v1 = re.sub(r"\D", "", v1).zfill(12)

                        if v1 == "000000000000" or not v2:
                            return ""
                        return f"{v1[:4]}.{v1[4:6]}.{v1[8:]}.{v2}"

                    resultados = df.apply(transformar, axis=1)
                    validos = resultados[resultados != ""]

                    if not validos.empty:
                        resultado_final = ";".join(validos)
                        st.success("✔️ Datos unificados correctamente")
                        st.metric("📊 Registros unificados", len(validos))

                        # Usamos st.code para mantener el botón de copiar
                        st.code(resultado_final, language=None)
                    else:
                        st.warning("⚠️ No se encontraron datos válidos.")

                    # ======================================================
                    # BOTÓN: GENERAR REPORTE DE NUEVOS FUNCIONARIOS
                    # ======================================================
                    st.divider()
                    st.subheader("📋 Reporte de instituciones con nuevos funcionarios")

                    col_institucion = detectar_columna(df.columns, ["institucion", "institución"])
                    col_entidad = detectar_columna(df.columns, ["entidad contratante", "entidad"])
                    col_razon_social = detectar_columna(df.columns, ["razon social", "razón social"])
                    col_tipo = detectar_columna(df.columns, ["tipo"])
                    col_monto = detectar_columna(df.columns, ["monto neto", "monto"])
                    col_odc = detectar_columna(df.columns, ["odc", "certificacion de contrato", "certificación de contrato"])
                    col_moneda = detectar_columna(df.columns, ["moneda"])

                    # Para cada fila se usa: Institución si tiene valor, si no
                    # Entidad Contratante, si no Razón Social.
                    def obtener_institucion(fila):
                        for col in (col_institucion, col_entidad, col_razon_social):
                            if col and pd.notna(fila[col]) and str(fila[col]).strip():
                                return str(fila[col]).strip()
                        return ""

                    if not (col_institucion or col_entidad or col_razon_social):
                        st.warning(
                            "⚠️ No se encontró una columna de Institución, Entidad Contratante o Razón Social "
                            "en el archivo. No es posible generar el reporte."
                        )
                    else:
                        df_institucion_vista = df.apply(
                            lambda fila: pd.Series({
                                "Institución": fila[col_institucion] if col_institucion else "",
                                "Entidad Contratante": fila[col_entidad] if col_entidad else "",
                                "Valor usado": obtener_institucion(fila),
                            }),
                            axis=1,
                        )

                        filas_reporte = []
                        for _, fila in df.iterrows():
                            texto_institucion = obtener_institucion(fila)
                            nombres_func, decretos = buscar_funcionario_nuevo(texto_institucion)

                            filas_reporte.append({
                                "Nuevo Funcionario": nombres_func or "",
                                "Decreto": decretos or "",
                                "Nombre de la Institución": texto_institucion,
                                "Tipo": fila[col_tipo] if col_tipo else "",
                                "Estructura Programática": fila[col_estructura] if col_estructura else "",
                                "Monto Neto": fila[col_monto] if col_monto else "",
                                "Número Libramiento(*)": fila[col_libramiento] if col_libramiento else "",
                                "#ODC/Certificación de Contrato(*)": fila[col_odc] if col_odc else "",
                                "Moneda": fila[col_moneda] if col_moneda else "",
                                "Razón Social": fila[col_razon_social] if col_razon_social else "",
                                "Expediente Trabajado": "",
                                "_match": bool(nombres_func),
                            })

                        df_reporte_completo = pd.DataFrame(filas_reporte)
                        total_coincidencias = int(df_reporte_completo["_match"].sum())

                        # Solo se conservan las filas con institución de nuevo funcionario
                        df_reporte = df_reporte_completo[df_reporte_completo["_match"]].reset_index(drop=True)

                        if total_coincidencias:
                            st.success(
                                f"✔️ Reporte generado — {total_coincidencias} registro(s) "
                                f"con institución de nuevo funcionario, resaltado en amarillo."
                            )
                        else:
                            st.info(
                                "El reporte se generó, pero ninguna institución del archivo "
                                "coincide con la lista de nuevos funcionarios."
                            )

                        # --------------------------------------------
                        # CONSTRUIR EXCEL CON RESALTADO AMARILLO
                        # --------------------------------------------
                        df_salida = df_reporte.drop(columns=["_match"])

                        buffer = BytesIO()
                        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                            df_salida.to_excel(writer, index=False, sheet_name="Reporte")
                            ws = writer.sheets["Reporte"]

                            amarillo = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
                            header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
                            header_font = Font(color="FFFFFF", bold=True)

                            n_cols = len(df_salida.columns)

                            # Encabezado
                            for c in range(1, n_cols + 1):
                                celda = ws.cell(row=1, column=c)
                                celda.fill = header_fill
                                celda.font = header_font
                                celda.alignment = Alignment(horizontal="center", vertical="center")

                            # Resaltar filas con coincidencia
                            for i, coincide in enumerate(df_reporte["_match"], start=2):
                                if coincide:
                                    for c in range(1, n_cols + 1):
                                        ws.cell(row=i, column=c).fill = amarillo

                            # Ancho de columnas automático
                            for c in range(1, n_cols + 1):
                                letra = get_column_letter(c)
                                max_len = max(
                                    [len(str(df_salida.iloc[r, c - 1])) for r in range(len(df_salida))]
                                    + [len(str(df_salida.columns[c - 1]))]
                                )
                                ws.column_dimensions[letra].width = min(max_len + 3, 45)

                            ws.freeze_panes = "A2"

                        buffer.seek(0)

                        st.dataframe(
                            df_salida.style.apply(
                                lambda row: [
                                    "background-color: #FFF9B0" if df_reporte.loc[row.name, "_match"] else ""
                                    for _ in row
                                ],
                                axis=1,
                            ),
                            use_container_width=True,
                        )

                        st.download_button(
                            label="⬇️ Descargar Reporte en Excel",
                            data=buffer,
                            file_name="Reporte_Nuevos_Funcionarios.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        )

            except Exception as e:
                st.error(f"Error en unificación: {e}")

# ======================================================
# MODO MANUAL (AUTOMÁTICO + BLOQUEO DE LETRAS)
# ======================================================
if modo.startswith("🧩"):

    st.subheader("🧩 Unificación manual")
    st.caption("Ideal cuando el volumen de trabajo es bajo")

    col1, col2 = st.columns(2)

    with col1:
        st.text_input(
            "Estructura Programática (12 dígitos)",
            placeholder="Ej: 010203040506",
            key="estructura",
            on_change=solo_numeros,
            args=("estructura",)
        )

    with col2:
        st.text_input(
            "Número de Libramiento (1 o 5 dígitos)",
            placeholder="Ej: 1234 o 12345",
            key="libramiento",
            on_change=solo_numeros,
            args=("libramiento",)
        )

    estructura = st.session_state.get("estructura", "")
    libramiento = st.session_state.get("libramiento", "")

    # 🔄 VALIDACIÓN + UNIFICACIÓN AUTOMÁTICA
    if estructura and libramiento:

        errores = False

        if len(estructura) != 12:
            st.error("❌ La Estructura Programática debe tener exactamente 12 dígitos")
            errores = True

        if not (1 <= len(libramiento) <= 5):
            st.error("❌ El Número de Libramiento debe tener entre 1 y 5 dígitos")
            errores = True

        if not errores:
            resultado = (
                f"{estructura[:4]}."
                f"{estructura[4:6]}."
                f"{estructura[8:]}."
                f"{libramiento}"
            )

            st.success("✔️ Unificación automática exitosa")
            st.code(resultado, language=None)

# ======================================================
# MODO EXTRACTOR DE CÓDIGOS REVISADOS
# ======================================================
if modo.startswith("✅"):
    render_extractor()

st.divider()
st.caption("DRCC DATA UNIFY - Herramienta diseñada para agilizar el proceso de firma en SIGEF")
