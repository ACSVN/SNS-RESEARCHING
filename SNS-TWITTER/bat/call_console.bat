@echo off


rem ===== 一つ前のディレクトリ名を取得 =====
pushd %~dp0..
set PROJECT_DIR=%CD%
popd


rem ===== batフォルダ内に移動し「x.txt」を削除 =====
cd %CD%
del x.txt /Q /F >NUL 2>NUL


rem ===== 画像データを含む「x.txt」を作成 =====
echo Getting your screen information. Please wait...
echo:
start "" /W dxdiag /T x


rem ===== 「x.txt」からスクリーン幅を取得 =====
setlocal enableDelayedExpansion
for /F "tokens=2 delims=:" %%a in ('find "Current Mode:" x.txt') do @set SCREEN_INFO=%%a
set SCREEN_WIDTH=%SCREEN_INFO:~1,5%
echo The width of your screen is %SCREEN_WIDTH%.
echo:


rem ===== 不要になったbatフォルダ内の「x.txt」を削除 =====
del x.txt /Q /F >NUL 2>NUL


rem ===== Excelコンソールへ =====
echo Moving for the console...
echo:
call :func_excel_console %SCREEN_WIDTH% %PROJECT_DIR%


endlocal
exit



rem ===== Excelコンソールの分岐 =====
:func_excel_console

set EXCEL_EXE=C:\Program Files\Microsoft Office\Office14\EXCEL.EXE

if %1 == 1920 (
	set EXCEL_FILENAME=%2\console\call.xlsm
) else (
	set EXCEL_FILENAME=%2\console\call2.xlsm
)

rem START "%EXCEL_EXE%" "%EXCEL_FILENAME"
echo %EXCEL_FILENAME%

pause
exit /B

