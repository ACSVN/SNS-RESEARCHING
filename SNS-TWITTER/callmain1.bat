cd %CD%

@echo off
del x.txt /q /f >nul 2>nul
start "" /w dxdiag /t x

timeout 5

start callmain2.bat
exit

