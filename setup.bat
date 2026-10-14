@echo off
rem Office Agent setup (Windows 10/11). Run once in the project root.
setlocal
cd /d "%~dp0"

where python >nul 2>nul || (echo [ERROR] Python 3.10+ not found in PATH & exit /b 1)
where claude >nul 2>nul || (echo [ERROR] Claude Code CLI not found in PATH & exit /b 1)

if not exist .venv python -m venv .venv
.venv\Scripts\python -m pip install --upgrade pip
.venv\Scripts\python -m pip install -r requirements.txt || (echo [ERROR] pip install failed. Use your internal PyPI mirror (pip config) if needed. & exit /b 1)

if not exist input mkdir input
if not exist output mkdir output

claude mcp remove office --scope project >nul 2>nul
claude mcp add --transport stdio --scope project --env OFFICE_AGENT_WORKDIR=%CD% --env PYTHONPATH=%CD% --env OFFICE_AGENT_READONLY_DIRS=input office -- "%CD%\.venv\Scripts\python.exe" -m office_agent.server
if errorlevel 1 (echo [ERROR] claude mcp add failed & exit /b 1)

echo.
echo Setup done. Next:
echo   1) .venv\Scripts\python tests\smoke_test.py   (verifies Office COM)
echo   2) claude   (then run /mcp to confirm "office" is connected)
endlocal
