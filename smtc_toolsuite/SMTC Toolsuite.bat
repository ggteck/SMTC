@echo off
@REM REM ─────────────────────────────────────────────────────────────────────────────
@REM REM  1) VARIABLES: adjust these if needed
@REM REM ─────────────────────────────────────────────────────────────────────────────
@REM SETLOCAL EnableDelayedExpansion

@REM REM Name of your SSH key files inside this folder:
@REM set PRIV_KEY=id_git_smtc
@REM set PUB_KEY=id_git_smtc.pub

@REM REM The SSH URL of your Git repo:
@REM set GIT_SSH_URL=git@ggteck.net:/home/git/repos/SMTC.git

@REM REM Where to clone (relative to this script):
@REM set CLONE_DIR=repo

@REM REM Python packages you need:
@REM set PIP_PKGS=streamlit XlsxWriter PyGithub pandas openpyxl ipywidgets fqdn tqdm python-dotenv

@REM REM ─────────────────────────────────────────────────────────────────────────────
@REM REM  2) CHECK & INSTALL GIT (via winget). If winget doesn’t exist, prompt user.
@REM REM ─────────────────────────────────────────────────────────────────────────────
@REM where git >nul 2>&1
@REM if errorlevel 1 (
@REM     echo.
@REM     echo ───────────────────────────────────────────────────────────────────
@REM     echo Git not found on this machine.
@REM     echo Attempting to install via winget...
@REM     echo ───────────────────────────────────────────────────────────────────
@REM     winget --silent install --id Git.Git >"%TEMP%\winget_git_log.txt" 2>&1
@REM     if errorlevel 1 (
@REM         echo.
@REM         echo ERROR: Could not install Git automatically. 
@REM         echo Please install “Git for Windows” manually from https://git-scm.com/download/win
@REM         echo and then re-run this script.
@REM         pause
@REM         exit /b 1
@REM     ) else (
@REM         echo Git installed successfully.
@REM     )
@REM ) else (
@REM     echo Git is already installed.
@REM )

@REM REM ─────────────────────────────────────────────────────────────────────────────
@REM REM  3) SET UP %USERPROFILE%\.ssh\, COPY KEYS, FIX PERMISSIONS
@REM REM ─────────────────────────────────────────────────────────────────────────────
@REM set SSH_DIR=%USERPROFILE%\.ssh
@REM if not exist "%SSH_DIR%" (
@REM     mkdir "%SSH_DIR%"
@REM )

@REM REM Copy the key files (overwrite if they exist)
@REM copy /Y "%~dp0%PRIV_KEY%" "%SSH_DIR%\%PRIV_KEY%" >nul
@REM copy /Y "%~dp0%PUB_KEY%" "%SSH_DIR%\%PUB_KEY%"     >nul

@REM REM Restrict permissions on the private key so only the user can read it:
@REM icacls "%SSH_DIR%\%PRIV_KEY%" /inheritance:r /grant:r "%USERNAME%:R" >nul
@REM REM (Optional) also restrict the .ssh folder itself:
@REM icacls "%SSH_DIR%" /inheritance:r /grant:r "%USERNAME%:F" >nul

@REM REM Ensure ssh-agent is running (Windows OpenSSH)
@REM echo Starting ssh-agent...
@REM powershell -Command "Start-Service ssh-agent" >nul 2>&1

@REM REM Add the private key to the agent (so Git can use it without passphrase prompt)
@REM echo Adding SSH key to agent...
@REM ssh-add "%SSH_DIR%\%PRIV_KEY%" >nul 2>&1

@REM REM ─────────────────────────────────────────────────────────────────────────────
@REM REM  4) CLONE OR PULL INTO THE CURRENT DIRECTORY (where this .bat resides)
@REM REM ─────────────────────────────────────────────────────────────────────────────
@REM REM %~dp0 expands to the folder containing this batch file (e.g. C:\MyApp\)
@REM if not exist "%~dp0\.git" (
@REM     echo.
@REM     echo Cloning repository into current folder...
@REM     git clone "%GIT_SSH_URL%" "%~dp0"
@REM     if errorlevel 1 (
@REM         echo ERROR: “git clone” failed.
@REM         echo   - Check that "%PRIV_KEY%" is correct and “ssh-add” succeeded.
@REM         echo   - Ensure your SSH URL is correct.
@REM         pause
@REM         exit /b 1
@REM     )
@REM ) else (
@REM     echo.
@REM     echo Repository already exists; pulling latest changes...
@REM     git -C "%~dp0" pull
@REM     if errorlevel 1 (
@REM         echo ERROR: “git pull” failed.
@REM         pause
@REM         exit /b 1
@REM     )
@REM )

@REM REM ─────────────────────────────────────────────────────────────────────────────
@REM REM  5) INSTALL PYTHON DEPENDENCIES
@REM REM ─────────────────────────────────────────────────────────────────────────────
@REM echo.
@REM echo Installing Python packages (if missing)...
@REM pip list --disable-pip-version-check >"%TEMP%\pip_list.txt"
@REM for %%P in (%PIP_PKGS%) do (
@REM     findstr /B /I "%%P " "%TEMP%\pip_list.txt" >nul
@REM     if errorlevel 1 (
@REM         echo   • Installing %%P...
@REM         pip install %%P
@REM         if errorlevel 1 (
@REM             echo     ERROR: Could not install %%P. Please install it manually.
@REM             pause
@REM             exit /b 1
@REM         )
@REM     ) else (
@REM         echo   • %%P already installed.
@REM     )
@REM )
@REM del "%TEMP%\pip_list.txt"

@REM REM ─────────────────────────────────────────────────────────────────────────────
@REM REM  6) LAUNCH STREAMLIT FROM THE CURRENT DIRECTORY
@REM REM ─────────────────────────────────────────────────────────────────────────────
@REM echo.
@REM echo Starting Streamlit app...
@REM cd "%~dp0"
streamlit run "Home.py"

@REM ENDLOCAL

