@echo off
rem Drag the Tableau export CSV (custom SQL data source, "Download > Data", or SQL result) onto this file.
rem Builds Output\Trucks Report <date>.xlsx (formula-linked: Dashboard, Worklist, Report Card, Data ...)
rem + late_trucks_feed.json/.js, and adds every truck to History\ (history CSV + Trucks History Summary.xlsx).
rem Save the previous Trucks Report first so your Worklist entries carry forward.
rem Optional: add --arr-basis team below to score arrivals against the Report Card Schedule tab.
rem HUB = folder holding Coordinator Hub.html; the feed is copied there so the Day-of board shows real late trucks.
setlocal
cd /d "%~dp0"
set "HUB=%~dp0hub"
if "%~1"=="" (
  echo Drag the Trucks Report export CSV onto this file.
  pause
  exit /b 1
)
set "PY=py -3"
where py >nul 2>nul || set "PY=python"
%PY% late_trucks.py "%~1" --days 7 --hub "%HUB%" %2 %3 %4 %5 %6 %7 %8 %9
if errorlevel 1 (
  echo.
  echo Build failed - see message above.
  pause
  exit /b 1
)
start "" "%~dp0Output"
