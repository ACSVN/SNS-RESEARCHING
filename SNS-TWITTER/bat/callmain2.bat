setlocal enableDelayedExpansion
for /f "tokens=2 delims=:" %%a in ('find "Current Mode:" x.txt') do @SET test1=%%a
echo %test1%

start %CD%\testmain.bat %test1%

endlocal
@echo del x.txt /q /f >nul 2>nul

exit

