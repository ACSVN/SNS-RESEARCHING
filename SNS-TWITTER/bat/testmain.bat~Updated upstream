rem cd %CD%

SET EXCEL_EXE=C:\Program Files\Microsoft Office\Office14\EXCEL.EXE
SET EXCEL_FILE1=C:\SNS-TWITTER\vba\call.xlsm
SET EXCEL_FILE2=C:\SNS-TWITTER\vba\call2.xlsm

set str=%1
echo %str%

echo %str:~0,4%
echo %str:~7,11%

if %str:~0,4% == 1920 (
START "%EXCEL_EXE%" "%EXCEL_FILE1%"
rem start bat1.bat %str:~0,4%
exit
) else (
START "%EXCEL_EXE%" "%EXCEL_FILE2%"
rem start bat2.bat %str:~0,4%
exit
)