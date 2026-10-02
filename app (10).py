import html
import json

import requests
import streamlit as st
import streamlit.components.v1 as components

from extractor import (
    COL_CODIGO,
    COL_LIBRAMIENTO,
    extract_reviewed,
    filter_by_libramiento,
)

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
    "Sube un Excel (.xlsx / .xls). Se busca la columna **Revisada**, se toman las filas con "
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
