@echo off


rem ===== Excel (call.xlsm, call2.xlsm)からの引数 =====
rem %1 ... 会社名 (Powerpointスライド1枚目用)
rem %2 ... Twitterアカウント名 (セーブフォルダ名)
rem %3 ... SCREEN_SWITCH (Width:1920=1, Width:Others=2, Error=0)


rem ===== 一つ前のディレクトリ名を取得 =====
pushd %~dp0..
set PROJECT_DIR=%CD%
popd


rem ===== Excelコンソールを閉じる =====
taskkill /F /IM EXCEL.EXE


rem ===== 「MainProgram.R」でTwitterデータの取得 =====
call :func_mainprogram_r %1 %2 %PROJECT_DIR%


rem ===== Twitterアカウントのスクリーンショットを撮る =====
call :func_take_capture %2 %PROJECT_DIR%


rem ===== Twitterアカウント画像の貼り付け =====
call :func_paste_img_twitter %1 %2 %3 %PROJECT_DIR%


rem ===== 最後にoutputフォルダを開いて終了 =====
start %PROJECT_DIR%\output

exit





rem ==================== SUBROUTINE ====================

:func_mainprogram_r

rem %1 ... 会社名
rem %2 ... Twitterアカウント名
rem %3 ... PROJECT_DIR

echo Starting to get Twitter Contents.
echo:
echo ******************************
echo CompanyName : %1
echo AccountName : %2
echo ******************************
echo:

"C:\Program Files\R\R-3.6.1\bin\Rscript.exe" %3\sources\MainProgram.R %1 %2
echo Running %3\sources\MainProgram.R. Please wait...
echo:

exit /B





:func_take_capture

rem %1 ... Twitterアカウント名
rem %2 ... PROJECT_DIR

start https://mobile.twitter.com/%1
del %2\images\Cap.png
del %2\images\top.png
cd %CD%
timeout /T 20 /nobreak >NUL
powershell -NoProfile -ExecutionPolicy Unrestricted .\scrshotPS.ps1

timeout /T 35 /nobreak >NUL
taskkill /F /IM chrome.exe /T

exit /B





:func_paste_img_twitter

rem %1 ... 会社名
rem %2 ... Twitterアカウント名
rem %3 ... SCREEN_SWITCH
rem %4 ... PROJECT_DIR

if %3==1 (
	set INSERT_IMG=insertimg1.r
) else if %3==2 (
	set INSERT_IMG=insertimg2.r
) else (
	echo Error Occured.
	pause
	exit
)

"C:\Program Files\R\R-3.6.1\bin\rscript.exe" %4\sources\%INSERT_IMG% %1 %2
echo Running %4\%INSERT_IMG%. Please wait...
echo:

exit /B
