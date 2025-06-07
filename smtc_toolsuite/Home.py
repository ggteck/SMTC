import os
import subprocess
from pathlib import Path
import streamlit as st

# path to your repo
REPO_PATH = Path(__file__).parent

scripts_catalog = st.Page(
    "Indice/scripts_catalog.py", title="Indice", icon=":material/dashboard:", default=True
)
manufacturing_plan = st.Page(
    "Herramientas/manufacturing_plan.py", title="Manufacturing plan", icon=":material/dashboard:"
)
local_shipments = st.Page(
    "Herramientas/local_shipments_l.py", title="Seguimiento a embarques", icon=":material/dashboard:"
)
doc = st.Page(
    "Documentacion/doc.py", title="Ejemplo de documento", icon=":material/dashboard:"
)

pg = st.navigation(
    {
        "Indice": [scripts_catalog],
        "Herrramientas": [manufacturing_plan, local_shipments],
        "Documentacion": [doc],
    }
)

pg.run()
if st.button("Actualizar repositorio"):
    res = subprocess.run(
        ["git", "pull"],
        cwd=str(REPO_PATH),
        capture_output=True,
        text=True
    )
    if res.returncode == 0:
        st.success(f"✅ Actualizado:\n{res.stdout}")
    else:
        st.error(f"❌ Error:\n{res.stderr}")