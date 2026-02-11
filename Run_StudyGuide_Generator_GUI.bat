@echo off
setlocal enabledelayedexpansion
cd /d "%~dp0"

REM ===== Settings =====
set "SCRIPT=generate_study_guide.py"

echo.
echo === Study Guide Generator (GUI) ===
echo Folder: %cd%
echo Script: %SCRIPT%
echo.

REM ===== Basic checks =====
if not exist "%SCRIPT%" (
  echo ERROR: Cannot find "%SCRIPT%" in this folder.
  echo Put the .bat and the .py in the SAME folder, or update SCRIPT in this .bat file.
  echo.
  dir /b *.py
  echo.
  pause
  exit /b 1
)

REM ===== Choose Python =====
set "PYEXE="

if exist ".venv\\Scripts\\python.exe" (
  set "PYEXE=.venv\\Scripts\\python.exe"
) else (
  REM Try the Python launcher (preferred on Windows)
  where py >nul 2>nul
  if !errorlevel! EQU 0 (
    set "PYEXE=py -3"
  ) else (
    REM Try plain python
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
  echo Quick setup:
  echo   py -3 -m venv .venv
  echo   .venv\\Scripts\\python.exe -m pip install -U pip
  echo   .venv\\Scripts\\python.exe -m pip install openai python-docx
  echo.
  pause
  exit /b 1
)

echo Using: %PYEXE%
echo.

REM ===== Run =====
%PYEXE% "%SCRIPT%" --gui

echo.
echo (If the window closed immediately, run this .bat from Command Prompt to see the error.)
pause
