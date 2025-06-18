import streamlit as st
import streamlit.components.v1 as components
from pathlib import Path


REPO_PATH = Path(__file__).resolve().parent.parent.parent / "Documentacion\\manufacturing_plan.html"

# Load your static HTML file
with open(REPO_PATH , "r", encoding="utf-8") as f:
    html_content = f.read()

# Display it inside the Streamlit app
components.html(html_content, height=800, width=1000,  scrolling=True)
