@echo off
echo Checking if JupyterLab is installed...

REM Check if JupyterLab is available
jupyter-lab --version >nul 2>&1

REM If the last command (jupyter-lab --version) failed, install JupyterLab
if errorlevel 1 (
    echo JupyterLab is not installed. Installing now...
    pip install jupyterlab XlsxWriter PyGithub pandas openpyxl ipywidgets fqdn tqdm python-dotenv
    echo JupyterLab has been installed.
)

echo Starting JupyterLab in the specified directory...

REM Change the directory to your specified folder
cd /D "%~1"

REM Start JupyterLab
call jupyter-lab

echo JupyterLab has started.
pause
