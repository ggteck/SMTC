import streamlit as st
import subprocess
from pathlib import Path

# path to your repo
REPO_PATH = Path(__file__).parent.parent

GIT_EXE = REPO_PATH.parent.parent / "mingit" / "cmd" / "git.exe"

if st.button("Actualizar repositorio"):
    commands = [
        [str(GIT_EXE), "fetch", "origin"],
        [str(GIT_EXE), "reset", "--hard", "origin/main"],
        [str(GIT_EXE), "clean", "-fd"],
        [str(GIT_EXE), "pull"],
    ]
    outputs = []
    failed_res = None

    for command in commands:
        res = subprocess.run(
            command,
            cwd=str(REPO_PATH),
            capture_output=True,
            text=True
        )
        outputs.append(f"$ {' '.join(command)}\n{res.stdout}{res.stderr}")
        if res.returncode != 0:
            failed_res = res
            break

    output_text = "\n".join(outputs)
    if failed_res is None:
        st.success(f"✅ Actualizado:\n{output_text}")
    else:
        st.error(f"❌ Error:\n{output_text}")
