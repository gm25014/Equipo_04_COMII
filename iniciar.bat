@echo off
setlocal
cd /d "%~dp0"

if exist "fluji.exe" (
    start "" "fluji.exe"
    exit /b 0
)

where py >nul 2>nul
if errorlevel 1 (
    where python >nul 2>nul
    if errorlevel 1 goto no_python
    set "PYTHON_CMD=python"
) else (
    set "PYTHON_CMD=py -3"
)

if exist ".venv\Scripts\python.exe" goto use_project_env

%PYTHON_CMD% -c "import pandas, openpyxl, reportlab" >nul 2>nul
if errorlevel 1 goto create_project_env

echo Dependencias detectadas en Python. Abriendo la interfaz grafica...
%PYTHON_CMD% "src\comparador.py"
if errorlevel 1 goto execution_error
exit /b 0

:create_project_env
echo Creando entorno virtual e instalando dependencias...
%PYTHON_CMD% -m venv ".venv"
if errorlevel 1 goto setup_error

:use_project_env
echo Verificando dependencias...
".venv\Scripts\python.exe" -m pip install --disable-pip-version-check -r "requirements.txt"
if errorlevel 1 goto dependencies_error

echo Abriendo la interfaz grafica...
".venv\Scripts\python.exe" "src\comparador.py"
if errorlevel 1 goto execution_error
exit /b 0

:no_python
echo ERROR: No se encontro Python. Instale Python y vuelva a ejecutar iniciar.bat.
pause
exit /b 1

:setup_error
echo ERROR: No se pudo crear el entorno virtual de Python.
pause
exit /b 1

:dependencies_error
echo ERROR: No se pudieron instalar las dependencias definidas en requirements.txt.
pause
exit /b 1

:execution_error
echo ERROR: El comparador no pudo completar el procesamiento.
pause
exit /b 1
