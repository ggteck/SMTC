@echo off
REM Check if Streamlit is installed by trying to import it
python -c "import streamlit" >nul 2>&1
if errorlevel 1 (
    echo Streamlit not found, installing Streamlit...
    pip install streamlit XlsxWriter PyGithub pandas openpyxl ipywidgets fqdn tqdm python-dotenv
)

streamlit run "local_shipments_l.py"
    