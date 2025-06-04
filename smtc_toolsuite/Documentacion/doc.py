import streamlit as st
import streamlit.components.v1 as components

# Load your static HTML file
with open(r"E:\silentgil\SMTC\Documentacion\carga_demanda.html", "r", encoding="utf-8") as f:
    html_content = f.read()

# Display it inside the Streamlit app
components.html(html_content, height=600, scrolling=True)