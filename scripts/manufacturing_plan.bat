@echo off
REM Check if Streamlit is installed by trying to import it
python -c "import streamlit" >nul 2>&1
if errorlevel 1 (
    echo Streamlit not found, installing Streamlit...
    python -m pip install streamlit
)

streamlit run "manufacturing_plan.py"
