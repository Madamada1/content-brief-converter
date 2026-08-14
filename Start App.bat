@echo off
setlocal
cd /d "%~dp0"

if not exist venv (
    echo Setting up the app for the first time - this can take a minute...
    py -3 -m venv venv
    if errorlevel 1 (
        echo.
        echo Could not find Python. Please install Python from https://www.python.org/downloads/
        echo and make sure to tick "Add python.exe to PATH" during install, then run this file again.
        pause
        exit /b 1
    )
)

call venv\Scripts\activate.bat
pip install -r requirements.txt --quiet --disable-pip-version-check

start "Content Brief Converter - leave this window open while using the app" cmd /k venv\Scripts\python.exe app.py

timeout /t 3 /nobreak >nul
start "" http://127.0.0.1:5050

endlocal
