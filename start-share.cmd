@echo off
title Stock Review
setlocal EnableExtensions
cd /d "%~dp0"

set "SRC=%~dp0"
if "%SRC:~-1%"=="\" set "SRC=%SRC:~0,-1%"
set "DEST=%LOCALAPPDATA%\BuildsafeStockReview"
set "URL=http://localhost:5174/?code=SS-DS24"

echo.
echo Stock Review
echo ------------
echo Copying to this computer so it can start...
echo %DEST%
echo.

if not exist "%DEST%" mkdir "%DEST%"
robocopy "%SRC%" "%DEST%" /E /XO /R:1 /W:1 /NFL /NDL /NJH /NJS /NC /NS /XF "START STOCK REVIEW.cmd" "start-share.cmd" "Open Demand SS-DS24.url" "READ ME - how to open.txt"
if errorlevel 8 (
  echo Copy failed. Ask Bo to check the Z: folder.
  pause
  exit /b 1
)

set "NODE=%DEST%\runtime\node.exe"
if not exist "%NODE%" (
  echo node.exe is missing in runtime\. Ask Bo to recopy the folder.
  pause
  exit /b 1
)

cd /d "%DEST%"

netstat -ano | findstr ":5174" | findstr "LISTENING" >nul
if not errorlevel 1 (
  echo Already running. Opening your browser...
  start "" "http://localhost:5174/?code=SS-DS24"
  echo Leave the other Stock Review window open.
  pause
  exit /b 0
)

echo Starting. Your browser will open in a few seconds.
echo Leave this window open until you are finished.
echo.
start "" cmd /c "timeout /t 5 /nobreak >nul & start http://localhost:5174/?code=SS-DS24"
"%NODE%" server.js
echo.
echo Stock Review has stopped.
pause
