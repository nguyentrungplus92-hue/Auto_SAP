@echo off
echo ============================================
echo   SAP Auto Tasks - Starting...
echo ============================================

REM Start SAP GUI and auto login
echo [1/3] Starting SAP GUI...
start "" "C:\Program Files (x86)\SAP\FrontEnd\SAPgui\sapshcut.exe" -system=V2Q -client=250 -user=ITS-2015030 -pw=Trung999
timeout /t 20 /nobreak

REM Start Web Server
echo [2/3] Starting Web Server...
start "SAP_Auto_Web" cmd /k "cd /d C:\Project\Auto_SAP\Auto_SAP\sap_auto && C:\Project\Auto_SAP\Auto_SAP\venv\Scripts\python.exe manage.py runserver 0.0.0.0:8010"
timeout /t 5 /nobreak

REM Start Scanner
echo [3/3] Starting Scanner...
start "SAP_Auto_Scanner" cmd /k "cd /d C:\Project\Auto_SAP\Auto_SAP\sap_auto && C:\Project\Auto_SAP\Auto_SAP\venv\Scripts\python.exe manage.py scan_tasks"

echo ============================================
echo   All services started!
echo ============================================