@echo off
setlocal
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" goto create_project_env
goto install_build_dependencies

:create_project_env
where py >nul 2>nul
if errorlevel 1 (
    where python >nul 2>nul
    if errorlevel 1 goto no_python
    python -m venv ".venv"
) else (
    py -3 -m venv ".venv"
)
if errorlevel 1 goto setup_error

:install_build_dependencies
echo Instalando dependencias de empaquetado...
".venv\Scripts\python.exe" -m pip install --disable-pip-version-check -r "requirements-build.txt"
if errorlevel 1 goto dependencies_error

echo Creando fluji.exe en la carpeta del proyecto...
".venv\Scripts\python.exe" -m PyInstaller --noconfirm --clean --onefile --windowed --name fluji --distpath "." --workpath "build\work" --specpath "build" --add-data "%CD%\data\raw\cotizaciones.xlsx;data\raw" "src\comparador.py"
if errorlevel 1 goto build_error

echo Empaquetado completado: fluji.exe
exit /b 0

:no_python
echo ERROR: No se encontro Python. Instale Python y vuelva a ejecutar empaquetar.bat.
pause
exit /b 1

:setup_error
echo ERROR: No se pudo crear el entorno virtual de Python.
pause
exit /b 1

:dependencies_error
echo ERROR: No se pudieron instalar las dependencias de empaquetado.
pause
exit /b 1

:build_error
echo ERROR: No se pudo crear fluji.exe.
pause
exit /b 1
