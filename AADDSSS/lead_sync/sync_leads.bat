@echo off
setlocal

rem Go to ADS root
cd /d "%~dp0.."

rem Run lead sync script. Prefer py, fallback to python.
where py >nul 2>nul
if errorlevel 1 goto USE_PYTHON
py -3 "%~dp0sync_leads.py" %*
goto AFTER_RUN

:USE_PYTHON
python "%~dp0sync_leads.py" %*

:AFTER_RUN
set EXITCODE=%ERRORLEVEL%
echo.
if "%EXITCODE%"=="0" goto SUCCESS
echo Sync failed. Please check ADS log folder.
pause
exit /b %EXITCODE%

:SUCCESS
echo Sync finished. Output: ADS\lead_sync\leads.xlsx
pause
exit /b 0
