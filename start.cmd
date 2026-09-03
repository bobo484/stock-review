@echo off
cd /d "%~dp0"
set NODE=C:\Users\boc\AppData\Local\Programs\node\node.exe
if not exist "%NODE%" set NODE=node
"%NODE%" server.js
pause
