@echo off
echo ================================================
echo   SAP Auto Tasks - Starting...
echo ================================================
echo.

REM === Duong dan project ===
SET BASE_DIR=D:\Auto_SAP
SET PROJECT_DIR=D:\Auto_SAP\sap_auto

cd /d %PROJECT_DIR%

if not exist "manage.py" (
    echo [LOI] Khong tim thay manage.py tai: %PROJECT_DIR%
    pause
    exit /b
)

REM === Kich hoat venv ===
call %BASE_DIR%\venv\Scripts\activate

echo [OK] Project: %PROJECT_DIR%
echo [OK] Venv activated
echo.

REM === Chay Scanner (cua so rieng) ===
echo [1/2] Starting Scanner...
start "SAP-Scanner" cmd /k "cd /d %PROJECT_DIR% && call %BASE_DIR%\venv\Scripts\activate && python manage.py scan_tasks"

timeout /t 3 /nobreak >nul

REM === Chay Web Server ===
echo [2/2] Starting Web Server...
echo.
echo     Dashboard:  http://localhost:8000
echo     Admin:      http://localhost:8000/admin
echo.
python manage.py runserver 0.0.0.0:8000

pause
