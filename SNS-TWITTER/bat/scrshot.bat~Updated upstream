set THIS_PATH=%~dp0
set PROJECT_PATH=%THIS_PATH:bat\=%

start https://mobile.twitter.com/%2%
del %PROJECT_PATH%images\Cap.png
del %PROJECT_PATH%images\top.png
cd %THIS_PATH%
timeout /t 20 /nobreak >nul
powershell -NoProfile -ExecutionPolicy Unrestricted .\scrshotPS.ps1
exit 
