@echo off
color F0
title 多合一 Office KMS 啟用小工具 By.Ray Ver.1905.03.1110

:-------------------------------------
IF "%PROCESSOR_ARCHITECTURE%" EQU "amd64" (
>nul 2>&1 "%SYSTEMROOT%\SysWOW64\cacls.exe" "%SYSTEMROOT%\SysWOW64\config\system"
) ELSE (
>nul 2>&1 "%SYSTEMROOT%\system32\cacls.exe" "%SYSTEMROOT%\system32\config\system"
)

if '%errorlevel%' NEQ '0' (
    echo 請求系統管理員權限...
    goto UACPrompt
) else ( goto gotAdmin )

:UACPrompt
    echo Set UAC = CreateObject^("Shell.Application"^) > "%temp%\getadmin.vbs"
    set params = %*:"=""
    echo UAC.ShellExecute "cmd.exe", "/c ""%~s0"" %params%", "", "runas", 1 >> "%temp%\getadmin.vbs"

    "%temp%\getadmin.vbs"
    del "%temp%\getadmin.vbs"
    exit /B

:gotAdmin
    pushd "%CD%"
    CD /D "%~dp0"

:SETSERVER
SET KmsServer=KMS_SERVER_ADDRESS
SET KmsPort=1688

:menu
cls
echo ------------------------------------
echo     多合一 Office KMS 啟用小工具
echo          版本：1905.03.1110
echo ------------------------------------
echo 請選擇 Office 版本:
echo.
echo  1.Office 2010 Standard
echo  2.Office 2010 Professional Plus
echo  3.Office 2013 Standard
echo  4.Office 2013 Professional Plus
echo  5.Office 2016 Standard
echo  6.Office 2016 Professional Plus
echo  7.Office 2019 Standard
echo  8.Office 2019 Professional Plus
echo.
echo  q.結束
echo.
echo.
SET select=
SET /P select=請選擇：
cls 

IF /I '%select%'=='1' GOTO office2010std
IF /I '%select%'=='2' GOTO office2010proplus
IF /I '%select%'=='3' GOTO office2013std
IF /I '%select%'=='4' GOTO office2013proplus
IF /I '%select%'=='5' GOTO office2016std
IF /I '%select%'=='6' GOTO office2016proplus
IF /I '%select%'=='7' GOTO office2019std
IF /I '%select%'=='8' GOTO office2019proplus
IF /I '%select%'=='q' GOTO quit

echo 選擇錯誤，請重新選擇！ 
echo 按任意鍵返回主選單
pause>nul
Goto menu

:quit
exit

:error
cls
echo 系統找不到指定路徑，請確認已選擇正確的 Office 版本
echo.
echo 按任意鍵返回主選單
pause>nul
Goto menu


REM ----- Office 2010 Standard -----
:office2010std
echo ----------------------------
echo 已選擇 Office 2010 Standard
echo ----------------------------
echo.
if NOT EXIST "%ProgramFiles%\Microsoft Office\office14\ospp.vbs" goto office2010stdx86
echo 正在設定 KMS Server...
%systemroot%\system32\cscript //B "%ProgramFiles%\Microsoft Office\office14\ospp.vbs" /sethst:%KmsServer%
%systemroot%\system32\cscript //B "%ProgramFiles%\Microsoft Office\office14\ospp.vbs" /setprt:%KmsPort%
echo.
echo 正在安裝產品金鑰...
%systemroot%\system32\cscript "%ProgramFiles%\Microsoft Office\office14\ospp.vbs" /inpkey:V7QKV-4XVVR-XYV4D-F7DFM-8R6BM
echo.
echo 正在啟用 Office...
%systemroot%\system32\cscript "%ProgramFiles%\Microsoft Office\office14\ospp.vbs" /act
echo.
echo 處理完成!
echo.
echo 按任意鍵返回主選單
pause>nul
Goto menu

:office2010stdx86
cls
echo ----------------------------
echo 已選擇 Office 2010 Standard
echo ----------------------------
echo.
if NOT EXIST "%ProgramFiles(x86)%\Microsoft Office\office14\ospp.vbs" goto error
echo 正在設定 KMS Server...
%systemroot%\system32\cscript //B "%ProgramFiles(x86)%\Microsoft Office\office14\ospp.vbs" /sethst:%KmsServer%
%systemroot%\system32\cscript //B "%ProgramFiles(x86)%\Microsoft Office\office14\ospp.vbs" /setprt:%KmsPort%
echo.
echo 正在安裝產品金鑰...
%systemroot%\system32\cscript "%ProgramFiles(x86)%\Microsoft Office\office14\ospp.vbs" /inpkey:V7QKV-4XVVR-XYV4D-F7DFM-8R6BM
echo.
echo 正在啟用 Office...
%systemroot%\system32\cscript "%ProgramFiles(x86)%\Microsoft Office\office14\ospp.vbs" /act
echo.
echo 處理完成!
echo.
echo 按任意鍵返回主選單
pause>nul
Goto menu
REM ----- Office 2010 Standard -----

REM ----- Office 2010 ProPlus -----
:office2010proplus
echo -------------------------------------
echo 已選擇 Office 2010 Professional Plus
echo -------------------------------------
echo.
if NOT EXIST "%ProgramFiles%\Microsoft Office\office14\ospp.vbs" goto office2010ppx86
echo 正在設定 KMS Server...
%systemroot%\system32\cscript //B "%ProgramFiles%\Microsoft Office\office14\ospp.vbs" /sethst:%KmsServer%
%systemroot%\system32\cscript //B "%ProgramFiles%\Microsoft Office\office14\ospp.vbs" /setprt:%KmsPort%
echo.
echo 正在安裝產品金鑰...
%systemroot%\system32\cscript "%ProgramFiles%\Microsoft Office\office14\ospp.vbs" /inpkey:VYBBJ-TRJPB-QFQRF-QFT4D-H3GVB
echo.
echo 正在啟用 Office...
%systemroot%\system32\cscript "%ProgramFiles%\Microsoft Office\office14\ospp.vbs" /act
echo.
echo 處理完成!
echo.
echo 按任意鍵返回主選單
pause>nul
Goto menu

:office2010ppx86
cls
echo -------------------------------------
echo 已選擇 Office 2010 Professional Plus
echo -------------------------------------
echo.
if NOT EXIST "%ProgramFiles(x86)%\Microsoft Office\office14\ospp.vbs" goto error
echo 正在設定 KMS Server...
%systemroot%\system32\cscript //B "%ProgramFiles(x86)%\Microsoft Office\office14\ospp.vbs" /sethst:%KmsServer%
%systemroot%\system32\cscript //B "%ProgramFiles(x86)%\Microsoft Office\office14\ospp.vbs" /setprt:%KmsPort%
echo.
echo 正在安裝產品金鑰...
%systemroot%\system32\cscript "%ProgramFiles(x86)%\Microsoft Office\office14\ospp.vbs" /inpkey:VYBBJ-TRJPB-QFQRF-QFT4D-H3GVB
echo.
echo 正在啟用 Office...
%systemroot%\system32\cscript "%ProgramFiles(x86)%\Microsoft Office\office14\ospp.vbs" /act
echo.
echo 處理完成!
echo.
echo 按任意鍵返回主選單
pause>nul
Goto menu
REM ----- Office 2010 ProPlus -----


REM ----- Office 2013 Standard -----
:office2013std
echo ----------------------------
echo 已選擇 Office 2013 Standard
echo ----------------------------
echo.
if NOT EXIST "%ProgramFiles%\Microsoft Office\office15\ospp.vbs" goto office2013stdx86
echo 正在設定 KMS Server...
%systemroot%\system32\cscript //B "%ProgramFiles%\Microsoft Office\office15\ospp.vbs" /sethst:%KmsServer%
%systemroot%\system32\cscript //B "%ProgramFiles%\Microsoft Office\office15\ospp.vbs" /setprt:%KmsPort%
echo.
echo 正在安裝產品金鑰...
%systemroot%\system32\cscript "%ProgramFiles%\Microsoft Office\office15\ospp.vbs" /inpkey:KBKQT-2NMXY-JJWGP-M62JB-92CD4
echo.
echo 正在啟用 Office...
%systemroot%\system32\cscript "%ProgramFiles%\Microsoft Office\office15\ospp.vbs" /act
echo.
echo 處理完成!
echo.
echo 按任意鍵返回主選單
pause>nul
Goto menu

:office2013stdx86
cls
echo ----------------------------
echo 已選擇 Office 2013 Standard
echo ----------------------------
echo.
if NOT EXIST "%ProgramFiles(x86)%\Microsoft Office\office15\ospp.vbs" goto error
echo 正在設定 KMS Server...
%systemroot%\system32\cscript //B "%ProgramFiles(x86)%\Microsoft Office\office15\ospp.vbs" /sethst:%KmsServer%
%systemroot%\system32\cscript //B "%ProgramFiles(x86)%\Microsoft Office\office15\ospp.vbs" /setprt:%KmsPort%
echo.
echo 正在安裝產品金鑰...
%systemroot%\system32\cscript "%ProgramFiles(x86)%\Microsoft Office\office15\ospp.vbs" /inpkey:KBKQT-2NMXY-JJWGP-M62JB-92CD4
echo.
echo 正在啟用 Office...
%systemroot%\system32\cscript "%ProgramFiles(x86)%\Microsoft Office\office15\ospp.vbs" /act
echo.
echo 處理完成!
echo.
echo 按任意鍵返回主選單
pause>nul
Goto menu
REM ----- Office 2013 Standard -----

REM ----- Office 2013 ProPlus -----
:office2013proplus
echo -------------------------------------
echo 已選擇 Office 2013 Professional Plus
echo -------------------------------------
echo.
if NOT EXIST "%ProgramFiles%\Microsoft Office\office15\ospp.vbs" goto office2013ppx86
echo 正在設定 KMS Server...
%systemroot%\system32\cscript //B "%ProgramFiles%\Microsoft Office\office15\ospp.vbs" /sethst:%KmsServer%
%systemroot%\system32\cscript //B "%ProgramFiles%\Microsoft Office\office15\ospp.vbs" /setprt:%KmsPort%
echo.
echo 正在安裝產品金鑰...
%systemroot%\system32\cscript "%ProgramFiles%\Microsoft Office\office15\ospp.vbs" /inpkey:YC7DK-G2NP3-2QQC3-J6H88-GVGXT
echo.
echo 正在啟用 Office...
%systemroot%\system32\cscript "%ProgramFiles%\Microsoft Office\office15\ospp.vbs" /act
echo.
echo 處理完成!
echo.
echo 按任意鍵返回主選單
pause>nul
Goto menu

:office2013ppx86
cls
echo -------------------------------------
echo 已選擇 Office 2013 Professional Plus
echo -------------------------------------
echo.
if NOT EXIST "%ProgramFiles(x86)%\Microsoft Office\office15\ospp.vbs" goto error
echo 正在設定 KMS Server...
%systemroot%\system32\cscript //B "%ProgramFiles(x86)%\Microsoft Office\office15\ospp.vbs" /sethst:%KmsServer%
%systemroot%\system32\cscript //B "%ProgramFiles(x86)%\Microsoft Office\office15\ospp.vbs" /setprt:%KmsPort%
echo.
echo 正在安裝產品金鑰...
%systemroot%\system32\cscript "%ProgramFiles(x86)%\Microsoft Office\office15\ospp.vbs" /inpkey:YC7DK-G2NP3-2QQC3-J6H88-GVGXT
echo.
echo 正在啟用 Office...
%systemroot%\system32\cscript "%ProgramFiles(x86)%\Microsoft Office\office15\ospp.vbs" /act
echo.
echo 處理完成!
echo.
echo 按任意鍵返回主選單
pause>nul
Goto menu
REM ----- Office 2013 ProPlus -----


REM ----- Office 2016 Standard -----
:office2016std
echo ----------------------------
echo 已選擇 Office 2016 Standard
echo ----------------------------
echo.
if NOT EXIST "%ProgramFiles%\Microsoft Office\office16\ospp.vbs" goto office2016stdx86
echo 正在設定 KMS Server...
%systemroot%\system32\cscript //B "%ProgramFiles%\Microsoft Office\office16\ospp.vbs" /sethst:%KmsServer%
%systemroot%\system32\cscript //B "%ProgramFiles%\Microsoft Office\office16\ospp.vbs" /setprt:%KmsPort%
echo.
echo 正在安裝產品金鑰...
%systemroot%\system32\cscript "%ProgramFiles%\Microsoft Office\office16\ospp.vbs" /inpkey:JNRGM-WHDWX-FJJG3-K47QV-DRTFM
echo.
echo 正在啟用 Office...
%systemroot%\system32\cscript "%ProgramFiles%\Microsoft Office\office16\ospp.vbs" /act
echo.
echo 處理完成!
echo.
echo 按任意鍵返回主選單
pause>nul
Goto menu

:office2016stdx86
cls
echo ----------------------------
echo 已選擇 Office 2016 Standard
echo ----------------------------
echo.
if NOT EXIST "%ProgramFiles(x86)%\Microsoft Office\office16\ospp.vbs" goto error
echo 正在設定 KMS Server...
%systemroot%\system32\cscript //B "%ProgramFiles(x86)%\Microsoft Office\office16\ospp.vbs" /sethst:%KmsServer%
%systemroot%\system32\cscript //B "%ProgramFiles(x86)%\Microsoft Office\office16\ospp.vbs" /setprt:%KmsPort%
echo.
echo 正在安裝產品金鑰...
%systemroot%\system32\cscript "%ProgramFiles(x86)%\Microsoft Office\office16\ospp.vbs" /inpkey:JNRGM-WHDWX-FJJG3-K47QV-DRTFM
echo.
echo 正在啟用 Office...
%systemroot%\system32\cscript "%ProgramFiles(x86)%\Microsoft Office\office16\ospp.vbs" /act
echo.
echo 處理完成!
echo.
echo 按任意鍵返回主選單
pause>nul
Goto menu
REM ----- Office 2016 Standard -----

REM ----- Office 2016 Professional Plus -----
:office2016proplus
echo -------------------------------------
echo 已選擇 Office 2016 Professional Plus
echo -------------------------------------
echo.
if NOT EXIST "%ProgramFiles%\Microsoft Office\office16\ospp.vbs" goto office2016ppx86
echo 正在設定 KMS Server...
%systemroot%\system32\cscript //B "%ProgramFiles%\Microsoft Office\office16\ospp.vbs" /sethst:%KmsServer%
%systemroot%\system32\cscript //B "%ProgramFiles%\Microsoft Office\office16\ospp.vbs" /setprt:%KmsPort%
echo.
echo 正在安裝產品金鑰...
%systemroot%\system32\cscript "%ProgramFiles%\Microsoft Office\office16\ospp.vbs" /inpkey:XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99
echo.
echo 正在啟用 Office...
%systemroot%\system32\cscript "%ProgramFiles%\Microsoft Office\office16\ospp.vbs" /act
echo.
echo 處理完成!
echo.
echo 按任意鍵返回主選單
pause>nul
Goto menu

:office2016ppx86
cls
echo -------------------------------------
echo 已選擇 Office 2016 Professional Plus
echo -------------------------------------
echo.
if NOT EXIST "%ProgramFiles(x86)%\Microsoft Office\office16\ospp.vbs" goto error
echo 正在設定 KMS Server...
%systemroot%\system32\cscript //B "%ProgramFiles(x86)%\Microsoft Office\office16\ospp.vbs" /sethst:%KmsServer%
%systemroot%\system32\cscript //B "%ProgramFiles(x86)%\Microsoft Office\office16\ospp.vbs" /setprt:%KmsPort%
echo.
echo 正在安裝產品金鑰...
%systemroot%\system32\cscript "%ProgramFiles(x86)%\Microsoft Office\office16\ospp.vbs" /inpkey:XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99
echo.
echo 正在啟用 Office...
%systemroot%\system32\cscript "%ProgramFiles(x86)%\Microsoft Office\office16\ospp.vbs" /act
echo.
echo 處理完成!
echo.
echo 按任意鍵返回主選單
pause>nul
Goto menu
REM ----- Office 2016 Professional Plus -----


REM ----- Office 2019 Standard -----
:office2019std
echo ----------------------------
echo 已選擇 Office 2019 Standard
echo ----------------------------
echo.
if NOT EXIST "%ProgramFiles%\Microsoft Office\office16\ospp.vbs" goto office2019stdx86
echo 正在設定 KMS Server...
%systemroot%\system32\cscript //B "%ProgramFiles%\Microsoft Office\office16\ospp.vbs" /sethst:%KmsServer%
%systemroot%\system32\cscript //B "%ProgramFiles%\Microsoft Office\office16\ospp.vbs" /setprt:%KmsPort%
echo.
echo 正在安裝產品金鑰...
%systemroot%\system32\cscript "%ProgramFiles%\Microsoft Office\office16\ospp.vbs" /inpkey:6NWWJ-YQWMR-QKGCB-6TMB3-9D9HK
echo.
echo 正在啟用 Office...
%systemroot%\system32\cscript "%ProgramFiles%\Microsoft Office\office16\ospp.vbs" /act
echo.
echo 處理完成!
echo.
echo 按任意鍵返回主選單
pause>nul
Goto menu

:office2019stdx86
cls
echo ----------------------------
echo 已選擇 Office 2019 Standard
echo ----------------------------
echo.
if NOT EXIST "%ProgramFiles(x86)%\Microsoft Office\office16\ospp.vbs" goto error
echo 正在設定 KMS Server...
%systemroot%\system32\cscript //B "%ProgramFiles(x86)%\Microsoft Office\office16\ospp.vbs" /sethst:%KmsServer%
%systemroot%\system32\cscript //B "%ProgramFiles(x86)%\Microsoft Office\office16\ospp.vbs" /setprt:%KmsPort%
echo.
echo 正在安裝產品金鑰...
%systemroot%\system32\cscript "%ProgramFiles(x86)%\Microsoft Office\office16\ospp.vbs" /inpkey:6NWWJ-YQWMR-QKGCB-6TMB3-9D9HK
echo.
echo 正在啟用 Office...
%systemroot%\system32\cscript "%ProgramFiles(x86)%\Microsoft Office\office16\ospp.vbs" /act
echo.
echo 處理完成!
echo.
echo 按任意鍵返回主選單
pause>nul
Goto menu
REM ----- Office 2019 Standard -----

REM ----- Office 2019 Professional Plus -----
:office2019proplus
echo -------------------------------------
echo 已選擇 Office 2019 Professional Plus
echo -------------------------------------
echo.
if NOT EXIST "%ProgramFiles%\Microsoft Office\office16\ospp.vbs" goto office2019ppx86
echo 正在設定 KMS Server...
%systemroot%\system32\cscript //B "%ProgramFiles%\Microsoft Office\office16\ospp.vbs" /sethst:%KmsServer%
%systemroot%\system32\cscript //B "%ProgramFiles%\Microsoft Office\office16\ospp.vbs" /setprt:%KmsPort%
echo.
echo 正在安裝產品金鑰...
%systemroot%\system32\cscript "%ProgramFiles%\Microsoft Office\office16\ospp.vbs" /inpkey:NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP
echo.
echo 正在啟用 Office...
%systemroot%\system32\cscript "%ProgramFiles%\Microsoft Office\office16\ospp.vbs" /act
echo.
echo 處理完成!
echo.
echo 按任意鍵返回主選單
pause>nul
Goto menu

:office2019ppx86
cls
echo -------------------------------------
echo 已選擇 Office 2019 Professional Plus
echo -------------------------------------
echo.
if NOT EXIST "%ProgramFiles(x86)%\Microsoft Office\office16\ospp.vbs" goto error
echo 正在設定 KMS Server...
%systemroot%\system32\cscript //B "%ProgramFiles(x86)%\Microsoft Office\office16\ospp.vbs" /sethst:%KmsServer%
%systemroot%\system32\cscript //B "%ProgramFiles(x86)%\Microsoft Office\office16\ospp.vbs" /setprt:%KmsPort%
echo.
echo 正在安裝產品金鑰...
%systemroot%\system32\cscript "%ProgramFiles(x86)%\Microsoft Office\office16\ospp.vbs" /inpkey:NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP
echo.
echo 正在啟用 Office...
%systemroot%\system32\cscript "%ProgramFiles(x86)%\Microsoft Office\office16\ospp.vbs" /act
echo.
echo 處理完成!
echo.
echo 按任意鍵返回主選單
pause>nul
Goto menu
REM ----- Office 2019 Professional Plus -----