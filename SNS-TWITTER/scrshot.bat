set curdir=%~dp0

start https://mobile.twitter.com/%2%
del C:\Rproject\result\Cap.png
del C:\Rproject\result\top.png
cd %curdir%\SNS-Tw
timeout /t 20 /nobreak >nul
powershell -NoProfile -ExecutionPolicy Unrestricted .\scrshotPS.ps1
exit 
