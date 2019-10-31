@echo off
taskkill /F /IM EXCEL.EXE

set THIS_PATH=%~dp0


rem ===== Run MainProject created by R =====
"C:\Program Files\R\R-3.6.1\bin\rscript.exe" %THIS_PATH%\sources\MainProgram.R %1 %2
rem timeout /t 30 /nobreak >nul


rem ===== Take Screenshot of Twitter-Account =====
cd %THIS_PATH%
start scrshot.bat %1 %2
timeout /t 35 /nobreak >nul
taskkill /F /IM chrome.exe /T


rem ===== Insert Twitter-Account TopImage to "output/client_name/Twitter-PPT.pptx" =====
"C:\Program Files\R\R-3.6.1\bin\rscript.exe" %THIS_PATH%\insertimg1.r %1 %2


exit
