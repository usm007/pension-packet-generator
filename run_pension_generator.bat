@echo off
cd /d "%~dp0"
python pension_packet_generator.py
if errorlevel 1 (
  echo.
  echo The program exited with an error. Review the message above.
)
echo.
pause
