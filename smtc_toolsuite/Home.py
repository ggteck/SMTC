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

clear_to_build = st.Page(
    "Herramientas/clear_to_build.py", title="Clear to build", icon=":material/done_all:"
)

doc = st.Page(
    "Documentacion/plan_de_manufactura_doc.py", title="Plan de manufactura", icon=":material/edit_calendar:"
)
email_management = st.Page(
    "Open_Order/email_management.py", title="Email management", icon=":material/mail:"
)

pg = st.navigation(
    {
        "Indice": [scripts_catalog],
        "Herrramientas": [manufacturing_plan, local_shipments, carga_demanda, clear_to_build],
        "Open Order": [email_management],
        "Documentacion": [doc],
    }
)

pg.run()

