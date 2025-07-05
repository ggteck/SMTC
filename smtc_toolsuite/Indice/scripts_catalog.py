import streamlit as st
import subprocess
from pathlib import Path

# path to your repo
REPO_PATH = Path(__file__).parent.parent

GIT_EXE = REPO_PATH.parent.parent / "mingit" / "cmd" / "git.exe"

if st.button("Actualizar repositorio"):
    res = subprocess.run(
        [str(GIT_EXE), "pull"],
        cwd=str(REPO_PATH),
        capture_output=True,
        text=True
    )
    if res.returncode == 0:
        st.success(f"✅ Actualizado:\n{res.stdout}")
    else:
        st.error(f"❌ Error:\n{res.stderr}")