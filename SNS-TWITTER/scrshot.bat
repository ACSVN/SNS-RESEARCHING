set curdir=%~dp0

start https://mobile.twitter.com/%2%
del %curdir%\images\Cap.png
del %curdir%\images\top.png
cd %curdir%
timeout /t 20 /nobreak >nul
powershell -NoProfile -ExecutionPolicy Unrestricted .\scrshotPS.ps1
exit 
