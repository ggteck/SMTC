import subprocess
from pathlib import Path
import streamlit as st

# path to your repo
REPO_PATH = Path(__file__).parent

st.title("Central Hub")

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