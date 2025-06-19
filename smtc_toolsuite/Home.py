import os
import subprocess
from pathlib import Path
import streamlit as st

scripts_catalog = st.Page(
    "Indice/scripts_catalog.py", title="Indice", icon=":material/list:", default=True
)
manufacturing_plan = st.Page(
    "Herramientas/manufacturing_plan.py", title="Manufacturing plan", icon=":material/edit_calendar:"
)
local_shipments = st.Page(
    "Herramientas/local_shipments_l.py", title="Seguimiento a embarques", icon=":material/local_shipping:"
)

carga_demanda = st.Page(
    "Herramientas/carga_demanda_l.py", title="Cargar demanda", icon=":material/app_registration:"
)

doc = st.Page(
    "Documentacion/plan_de_manufactura_doc.py", title="Plan de manufactura", icon=":material/edit_calendar:"
)

pg = st.navigation(
    {
        "Indice": [scripts_catalog],
        "Herrramientas": [manufacturing_plan, local_shipments, carga_demanda],
        "Documentacion": [doc],
    }
)

pg.run()

