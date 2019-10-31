@echo off

set curdir=%~dp0

rem !!!call MainProgram!!!
rem "C:\Program Files\R\R-3.6.1\bin\rscript.exe" %curdir%\sources\MainProgram.R %1 %2
rem timeout /t 30 /nobreak >nul



rem Caution Sawa san!!!!! (from Nose)

rem !!!take screenshot!!!
rem cd %curdir%
rem start scrshot.bat %1 %2
rem timeout /t 35 /nobreak >nul
rem taskkill /F /IM chrome.exe /T



rem !!!insert Twitter top image to output/Twitter-PPT.pptx!!!
"C:\Program Files\R\R-3.6.1\bin\rscript.exe" %curdir%\insertimg.r %1 %2

rem !!!take screenshot!!!
cd %curdir%
start scrshot.bat %1 %2
timeout /t 35 /nobreak >nul
taskkill /F /IM chrome.exe /T

pause
exit
