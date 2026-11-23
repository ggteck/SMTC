import streamlit as st
import os
import shutil
import subprocess
from pathlib import Path

# path to your repo
REPO_PATH = Path(__file__).parent.parent

GIT_EXE = REPO_PATH.parent.parent / "mingit" / "cmd" / "git.exe"


def find_git():
    if GIT_EXE.is_file():
        return str(GIT_EXE)

    installed_git = shutil.which("git")
    if installed_git:
        return installed_git

    for variable, suffix in (
        ("ProgramFiles", ("Git", "cmd", "git.exe")),
        ("ProgramFiles(x86)", ("Git", "cmd", "git.exe")),
        ("LOCALAPPDATA", ("Programs", "Git", "cmd", "git.exe")),
    ):
        root = os.environ.get(variable)
        if root:
            candidate = Path(root).joinpath(*suffix)
            if candidate.is_file():
                return str(candidate)
    return None


if st.button("Actualizar repositorio"):
    git_exe = find_git()
    if git_exe is None:
        st.error("Git no está instalado o no se encontró. Instale Git for Windows y vuelva a intentar la actualización.")
        st.stop()

    commands = [
        [git_exe, "fetch", "origin"],
        [git_exe, "reset", "--hard", "origin/main"],
        [git_exe, "clean", "-fd"],
        [git_exe, "pull"],
    ]
    outputs = []
    failed_res = None

    for command in commands:
        try:
            res = subprocess.run(
                command,
                cwd=str(REPO_PATH),
                capture_output=True,
                text=True
            )
        except FileNotFoundError:
            st.error("Git no está instalado o no está disponible. Instale Git for Windows y vuelva a intentar la actualización.")
            st.stop()
        except OSError as exc:
            st.error(f"No se pudo ejecutar Git: {exc}")
            st.stop()
        outputs.append(f"$ {' '.join(command)}\n{res.stdout}{res.stderr}")
        if res.returncode != 0:
            failed_res = res
            break

    output_text = "\n".join(outputs)
    if failed_res is None:
        st.success(f"✅ Actualizado:\n{output_text}")
    else:
        st.error(f"❌ Error:\n{output_text}")
