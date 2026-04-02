@echo off
setlocal

set "ROOT_DIR=%~dp0"
cd /d "%ROOT_DIR%frontend" || goto :cd_error

set "FRONTEND_PORT=%~1"
if "%FRONTEND_PORT%"=="" set "FRONTEND_PORT=5173"

set "BACKEND_PORT=%~2"
if "%BACKEND_PORT%"=="" (
  if exist "%ROOT_DIR%.backend_port" (
    for /f "usebackq delims=" %%p in ("%ROOT_DIR%.backend_port") do set "BACKEND_PORT=%%p"
  )
)
if "%BACKEND_PORT%"=="" set "BACKEND_PORT=8001"

set "VITE_API_BASE=http://127.0.0.1:%BACKEND_PORT%/api"
set "VITE_FILE_BASE=http://127.0.0.1:%BACKEND_PORT%"

set "NEED_INSTALL=0"
if not exist "node_modules" set "NEED_INSTALL=1"
if not exist "node_modules\.bin\vite.cmd" set "NEED_INSTALL=1"

if "%NEED_INSTALL%"=="1" (
  echo [INFO] frontend dependencies missing or incomplete, installing...
  call npm.cmd install
  if errorlevel 1 (
    echo [ERROR] npm install failed.
    exit /b 1
  )
)

echo [INFO] Working dir: %CD%
echo [INFO] Starting frontend at http://127.0.0.1:%FRONTEND_PORT%
echo [INFO] API base: %VITE_API_BASE%

call npm.cmd run dev -- --host 127.0.0.1 --port %FRONTEND_PORT%
exit /b %ERRORLEVEL%

:cd_error
echo [ERROR] Failed to enter frontend directory.
exit /b 1
