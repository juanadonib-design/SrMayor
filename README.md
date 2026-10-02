Extractor de Códigos Revisados (Streamlit)
Sube un Excel (.xlsx / .xls), la app busca la columna Revisada, toma las filas con Verdadero y extrae sus Códigos. Cada código se muestra en una lista (uno debajo del otro) con su botón de copiado, más "Copiar todos" y descarga en TXT/CSV.

Cómo funciona
La app detecta sola la fila de encabezados (donde aparezcan Revisada y Código), en todas las hojas.
Marca el check "Aplicar Filtro de búsqueda para exportar." para habilitar el campo Número Libramiento(*) y exportar solo los códigos de ese libramiento (puedes escribir varios separados por coma).
La exportación aparece automáticamente.
Ejecutar en local
pip install -r requirements.txt
streamlit run app.py
Conexión con GitHub
Desplegar desde GitHub (Streamlit Community Cloud)

Sube app.py, extractor.py, requirements.txt y este README a un repositorio.
En https://share.streamlit.io crea una app nueva apuntando al repo, rama y app.py.
Cargar el Excel desde un repositorio de GitHub

En la app elige Desde GitHub e indica usuario, repositorio, ruta y rama. Para repos privados agrega un token (permiso de lectura de contenidos) en los Secrets de Streamlit:

GITHUB_TOKEN = "ghp_xxx"
