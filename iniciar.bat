@echo off
cd /d "%~dp0"
if exist "src\dist\comparador.exe" (
    start "" "src\dist\comparador.exe"
) else (
    python src/comparador.py
)