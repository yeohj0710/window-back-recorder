@echo off
setlocal
cd /d "%~dp0"
for %%I in ("%~dp0*.exe") do (
  start "" "%%~fI"
  exit /b
)
