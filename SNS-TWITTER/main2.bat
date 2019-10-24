@echo off

set THIS_PATH=%~dp0
set FILE_NAME_EXCEL=Twitter-Excel.xlsm
set FILE_NAME_PPT=Twitter-PPT.pptx

rem ===== Copy EXCEL & PPT file from "input" to "output" =====
rem copy %THIS_PATH%\input\%FILE_NAME_EXCEL% %THIS_PATH%\output\%FILE_NAME_EXCEL%
copy %THIS_PATH%\input\%FILE_NAME_PPT% %THIS_PATH%\output\%FILE_NAME_PPT%

rem ===== Take Screenshot of Twitter-Account =====
cd %THIS_PATH%
start scrshot.bat %1 %2
timeout /t 35 /nobreak >nul
taskkill /F /IM chrome.exe /T

rem ===== Insert Twitter-Account TopImage to "output/Twitter-PPT.pptx" =====
"C:\Program Files\R\R-3.6.1\bin\rscript.exe" %THIS_PATH%\insertimg2.r %1 %2

rem ===== Run MainProject created by R =====
"C:\Program Files\R\R-3.6.1\bin\rscript.exe" %THIS_PATH%\sources\MainProgram.R %1 %2
rem timeout /t 30 /nobreak >nul

exit
