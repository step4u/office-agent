@echo off
rem Headless run: run_agent.bat "instruction text"
setlocal
cd /d "%~dp0"
if "%~1"=="" (
  echo Usage: run_agent.bat "instruction"
  exit /b 1
)
claude -p "%~1" --allowedTools "mcp__office" --permission-mode acceptEdits
endlocal
