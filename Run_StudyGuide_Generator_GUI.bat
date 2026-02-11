@echo off
cd /d "%~dp0"
".venv\Scripts\python.exe" generate_study_guide.py --gui
pause
