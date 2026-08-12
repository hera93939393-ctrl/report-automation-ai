@echo off
cd /d "%~dp0"
echo Running... (see run_log.txt if it fails)
"C:\Users\lakka\AppData\Local\Programs\Python\Python313\python.exe" report_tool.py > run_log.txt 2>&1
echo.
echo ---- run_log.txt ----
type run_log.txt
echo ---------------------
pause
