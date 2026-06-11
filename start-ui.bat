@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo ========================================
echo   SKU Price Compare - Giao dien web
echo ========================================
echo.

where npm >nul 2>&1
if errorlevel 1 (
  echo [LOI] Khong tim thay npm / Node.js.
  echo        Cai Node.js: https://nodejs.org
  echo        Hoac chay trong terminal: npm run ui
  echo.
  pause
  exit /b 1
)

netstat -ano | findstr /R /C:":3847 " | findstr LISTENING >nul 2>&1
if not errorlevel 1 (
  echo Tim thay server cu tren port 3847 — dang tat de khoi dong lai...
  for /f "tokens=5" %%a in ('netstat -ano ^| findstr /R /C:":3847 " ^| findstr LISTENING') do (
    taskkill /F /PID %%a >nul 2>&1
  )
  timeout /t 2 /nobreak >nul
)

echo Dang khoi dong server tai http://localhost:3847
echo Giu cua so nay mo khi dang su dung. Nhan Ctrl+C de tat.
echo.

call npm run ui

echo.
if errorlevel 1 (
  echo [LOI] Server khong khoi dong duoc.
  echo  - Kiem tra file .env.dev va .env.production
  echo  - Port 3847 bi chiem: tat terminal cu hoac restart may
  echo  - Thu chay: npm run ui
) else (
  echo Server da tat.
)
echo.
pause
