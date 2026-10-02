"""Extractor de Códigos Revisados (Streamlit) — archivo único.
 
- Lee .xlsx y .xls (todas las hojas).
- Detecta automáticamente la fila de encabezados (donde estén "Revisada" y "Código").
- Extrae los códigos de las filas donde Revisada = Verdadero.
- Permite filtrar por Número Libramiento y copiar cada registro.
 
Ejecutar:  streamlit run app.py
"""
import html
import io
import json
import re
import unicodedata
 
import numpy as np
import pandas as pd
import requests
import streamlit as st
import streamlit.components.v1 as components
 
# ============================================================================
# LÓGICA DE EXTRACCIÓN
# ============================================================================
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
 
 
def _try_pip_install(package: str) -> bool:
    """Intenta instalar un paquete con pip (una vez, con tiempo límite). True si quedó importable."""
    import importlib
    import subprocess
    import sys
 
    for extra in ([], ["--break-system-packages"], ["--user"]):
        try:
            r = subprocess.run(
                [sys.executable, "-m", "pip", "install", "--quiet", *extra, package],
                capture_output=True,
                timeout=60,
            )
        except Exception:  # noqa: BLE001
            return False
        if r.returncode == 0:
            importlib.invalidate_caches()
            return True
    return False
 
 
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
 
 
def _read_all_sheets(data: bytes, filename: str):
    es_xlsx = data[:2] == b"PK"  # los .xlsx son ZIP, aunque el nombre diga .xls
 
    if es_xlsx:
        try:
            return pd.read_excel(io.BytesIO(data), sheet_name=None, header=None, engine="openpyxl", dtype=object)
        except ImportError:
            pass  # sin openpyxl: usamos el lector propio
        except Exception:  # noqa: BLE001
            pass
        try:
            return _read_xlsx_stdlib(data)
        except Exception as exc:  # noqa: BLE001
            raise ValueError(f"No se pudo leer el archivo Excel: {exc}")
 
    # .xls clásico: xlrd es una dependencia obligatoria del proyecto.
    # Está declarada en requirements.txt, por lo que debe instalarse durante el despliegue.
    try:
        return pd.read_excel(
            io.BytesIO(data),
            sheet_name=None,
            header=None,
            engine="xlrd",
            dtype=object,
        )
    except ImportError as exc:
        raise RuntimeError(
            "La dependencia obligatoria «xlrd» no está instalada. "
            "Verifica que requirements.txt incluya «xlrd>=2.0.1» "
            "y vuelve a desplegar la aplicación."
        ) from exc
    except Exception as exc:  # noqa: BLE001
        raise ValueError(f"No se pudo leer el archivo Excel .xls: {exc}")
 
 
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
 
 
# ============================================================================
# INTERFAZ STREAMLIT
# ============================================================================
st.set_page_config(page_title="Extractor de Códigos Revisados", page_icon="✅", layout="centered")
 
 
# ----------------------------------------------------------------------------
# Utilidades
# ----------------------------------------------------------------------------
def fetch_from_github(owner: str, repo: str, path: str, branch: str, token: str | None) -> bytes:
    """Descarga un archivo de un repositorio de GitHub (público o privado con token)."""
    url = f"https://api.github.com/repos/{owner}/{repo}/contents/{path.lstrip('/')}"
    headers = {"Accept": "application/vnd.github.raw+json"}
    if token:
        headers["Authorization"] = f"Bearer {token}"
    resp = requests.get(url, headers=headers, params={"ref": branch}, timeout=30)
    if resp.status_code == 404:
        raise FileNotFoundError("No se encontró el archivo (revisa repositorio, ruta, rama o token).")
    resp.raise_for_status()
    return resp.content
 
 
def get_secret_token() -> str | None:
    try:
        return st.secrets.get("GITHUB_TOKEN")
    except Exception:  # no hay secrets.toml
        return None
 
 
def copy_list_component(records: list[tuple[str, str]]):
    """Un registro debajo del otro, cada uno con su botón de copiado."""
    rows = "".join(
        f'<div class="row"><div class="info"><span class="code">{html.escape(code)}</span>'
        + (f'<span class="lib">Libramiento: {html.escape(lib)}</span>' if lib else "")
        + f'</div><button class="btn" data-v="{html.escape(code, quote=True)}">Copiar</button></div>'
        for code, lib in records
    )
    all_text = json.dumps("\n".join(code for code, _ in records)).replace("</", "<\\/")
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
      .lib {{ font-size:12px; color:#656d76; }}
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
    height = min(len(records) * 56 + 60, 520)
    components.html(page, height=height, scrolling=False)
 
 
# ----------------------------------------------------------------------------
# Interfaz
# ----------------------------------------------------------------------------
st.title("✅ Extractor de Códigos Revisados")
st.caption(
    "Sube un Excel (.xlsx / .xls). Ambos formatos son compatibles. Se busca la columna **Revisada**, se toman las filas con "
    "**Verdadero** y se extraen sus **Códigos**."
)
 
origen = st.radio("Origen del archivo", ["Subir archivo", "Desde GitHub"], horizontal=True)
 
data, fname = None, None
 
if origen == "Subir archivo":
    uploaded = st.file_uploader("Archivo Excel", type=["xlsx", "xls"])
    if uploaded is not None:
        data, fname = uploaded.getvalue(), uploaded.name
else:
    c1, c2 = st.columns(2)
    owner = c1.text_input("Usuario / organización")
    repo = c2.text_input("Repositorio")
    c3, c4 = st.columns([3, 1])
    path = c3.text_input("Ruta del archivo", placeholder="datos/reporte.xlsx")
    branch = c4.text_input("Rama", value="main")
    token = st.text_input(
        "Token de GitHub (solo repos privados)",
        type="password",
        value="",
        help="También puedes guardarlo como GITHUB_TOKEN en los secrets de Streamlit.",
    ) or get_secret_token()
    if st.button("Cargar desde GitHub", type="primary"):
        if not (owner and repo and path):
            st.warning("Completa usuario, repositorio y ruta.")
        else:
            try:
                with st.spinner("Descargando…"):
                    st.session_state["gh_data"] = (
                        fetch_from_github(owner, repo, path, branch or "main", token),
                        path.rsplit("/", 1)[-1],
                    )
            except Exception as exc:  # noqa: BLE001
                st.session_state.pop("gh_data", None)
                st.error(f"No se pudo descargar: {exc}")
    if "gh_data" in st.session_state:
        data, fname = st.session_state["gh_data"]
        st.success(f"Archivo cargado: {fname}")
 
if data is None:
    st.stop()
 
if not fname.lower().endswith((".xlsx", ".xls")):
    st.error("El archivo debe ser .xlsx o .xls.")
    st.stop()
 
try:
    df, has_lib, notes = st.cache_data(show_spinner="Leyendo Excel…")(extract_reviewed)(data, fname)
except Exception as exc:  # noqa: BLE001
    st.error(str(exc))
    st.stop()
 
for n in notes:
    st.info(n)
 
st.metric("Registros con Revisada = Verdadero", len(df))
st.divider()
 
usar_filtro = st.checkbox("Aplicar Filtro de búsqueda para exportar.")
 
result = df
if usar_filtro:
    if not has_lib:
        st.error("El archivo no tiene una columna de «Libramiento» para filtrar.")
        st.stop()
    libramiento = st.text_input(
        "Número Libramiento(*)",
        placeholder="Ej.: 1234 (puedes escribir varios separados por coma)",
    )
    if not libramiento.strip():
        st.warning("Escribe el Número Libramiento(*) para exportar los códigos relacionados.")
        st.stop()
    result = filter_by_libramiento(df, libramiento)
 
if result.empty:
    st.warning("No hay códigos para exportar con los criterios indicados.")
    st.stop()
 
records = list(zip(result[COL_CODIGO], result[COL_LIBRAMIENTO]))
 
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
    result.to_csv(index=False).encode("utf-8-sig"),
    file_name="codigos.csv",
    mime="text/csv",
    use_container_width=True,
)
 
