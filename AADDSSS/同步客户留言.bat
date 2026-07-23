@echo off
setlocal
cd /d "%~dp0"

where py >nul 2>nul
if errorlevel 1 goto USE_PYTHON
py -3 "%~dp0lead_sync\sync_leads_gui.py"
goto END

:USE_PYTHON
python "%~dp0lead_sync\sync_leads_gui.py"

:END
exit /b %ERRORLEVEL%
