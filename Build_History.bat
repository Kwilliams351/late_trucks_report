@echo off
rem Backfill: drag one or more old exports (or a folder of them) onto this file.
rem Merges them into History\late_trucks_history.csv (one row per TLM, newest copy wins) and rebuilds
rem late_trucks_daily.csv and "Trucks History Summary.xlsx". No daily report is built.
setlocal
cd /d "%~dp0"
if "%~1"=="" (
  echo Drag the export files or a folder of exports onto this file.
  pause
  exit /b 1
)
set "PY=py -3"
where py >nul 2>nul || set "PY=python"
%PY% late_trucks.py %* --history-only
if errorlevel 1 (
  echo.
  echo Build failed - see message above.
  pause
  exit /b 1
)
start "" "%~dp0History"
