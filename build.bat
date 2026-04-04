@echo off
REM ═══════════════════════════════════════════════════════════════
REM  Build: Combines modular planner-dev files into a single
REM  standalone HTML file. No admin rights required.
REM ═══════════════════════════════════════════════════════════════
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0build.ps1"
if %errorlevel% neq 0 (
  echo [BUILD] ERROR: Build failed!
  pause
  exit /b 1
)
pause
