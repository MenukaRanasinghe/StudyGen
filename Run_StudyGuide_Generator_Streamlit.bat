@echo off
setlocal enabledelayedexpansion
cd /d "%~dp0"

echo.
echo === Study Guide Generator (Streamlit) ===
echo Folder: %cd%
echo.

REM ---- Choose Python ----
set "PYEXE="
if exist ".venv\Scripts\python.exe" (
  set "PYEXE=.venv\Scripts\python.exe"
) else (
  where py >nul 2>nul
  if !errorlevel! EQU 0 (
    set "PYEXE=py -3"
  ) else (
    where python >nul 2>nul
    if !errorlevel! EQU 0 (
      set "PYEXE=python"
    )
  )
)

if "%PYEXE%"=="" (
  echo ERROR: Python was not found.
  echo Install Python 3, or create a venv in this folder named ".venv".
  echo.
  pause
  exit /b 1
)

echo Using: %PYEXE%

REM ---- Install deps (first run convenience) ----
if not exist ".venv" (
  echo.
  echo Creating venv...
  py -3 -m venv .venv
  set "PYEXE=.venv\Scripts\python.exe"
)

echo.
echo Installing requirements...
%PYEXE% -m pip install -U pip
%PYEXE% -m pip install -r requirements.txt

echo.
echo Starting Streamlit...
echo Open your browser at: http://localhost:8501

echo.
%PYEXE% -m streamlit run app.py

pause
