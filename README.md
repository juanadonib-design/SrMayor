# Extractor de Códigos Revisados (Streamlit)

Sube uno o varios Excel (`.xlsx` / `.xls`). La app busca la columna **Revisada**, toma las filas con
**Verdadero** y extrae sus **Códigos**. Cada código se muestra en una lista (uno debajo del otro) con su
botón de copiado, más "Copiar todos" y descarga en TXT/CSV.

## Cómo funciona

1. **Subir archivos:** puedes subir varios a la vez. El orden de carga define el Documento 1, 2, 3…
2. **Identificar registros nuevos:** con 2 o más archivos, activa el check para exportar solo los códigos del
   documento nuevo que **no** estaban en el/los documento(s) de referencia (por defecto, el último contra los
   anteriores). La comparación es por **Código** (sin importar mayúsculas, acentos ni espacios sobrantes).
   Se muestran los totales de *Nuevos* y *Repetidos*.
3. **Filtro por Número Libramiento(\*):** marca el check de filtro y escribe uno o varios números (separados
   por coma). Se aplica sobre lo que se vaya a exportar (incluido el resultado de "registros nuevos").
4. La exportación aparece automáticamente.

## Ejecutar en local

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Desplegar desde GitHub (Streamlit Community Cloud)

1. Sube `app.py`, `requirements.txt` y este README a un repositorio (en la raíz).
2. En <https://share.streamlit.io> crea una app apuntando al repo, rama y `app.py`.
