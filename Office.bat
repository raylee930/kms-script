@echo off
color F0
title 多合一 Office KMS 啟用小工具 By.Ray Ver.1811.04.0326

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

:menu
cls
echo ------------------------------------
echo     多合一 Office KMS 啟用小工具
echo          版本：1811.04.0326
echo ------------------------------------
echo 請選擇 Office 版本:
echo.
echo  1.Office 2010
echo  2.Office 2013
echo  3.Office 2016/2019
echo.
echo  q.結束
echo.
echo.
SET select=
SET /P select=請選擇：
cls 

IF /I '%select%'=='1' GOTO office14
IF /I '%select%'=='2' GOTO office15
IF /I '%select%'=='3' GOTO office16
REM IF /I '%select%'=='4' GOTO office16
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

REM ----- Office 2010 -----
:office14
echo -------------------
echo 已選擇 Office 2010
echo -------------------
echo.
if NOT EXIST "%ProgramFiles%\Microsoft Office\office14\ospp.vbs" goto office14x86
echo 正在設定 KMS Server...
%systemroot%\system32\cscript //B "%ProgramFiles%\Microsoft Office\office14\ospp.vbs" /sethst:KMS_IP_ADDRESS
%systemroot%\system32\cscript //B "%ProgramFiles%\Microsoft Office\office14\ospp.vbs" /setprt:1688
echo.
echo 正在啟用 Office...
%systemroot%\system32\cscript "%ProgramFiles%\Microsoft Office\office14\ospp.vbs" /act
echo.
echo 處理完成!
echo.
echo 按任意鍵返回主選單
pause>nul
Goto menu

:office14x86
cls
echo -------------------
echo 已選擇 Office 2010
echo -------------------
echo.
if NOT EXIST "%ProgramFiles(x86)%\Microsoft Office\office14\ospp.vbs" goto error
echo 正在設定 KMS Server...
%systemroot%\system32\cscript //B "%ProgramFiles(x86)%\Microsoft Office\office14\ospp.vbs" /sethst:KMS_IP_ADDRESS
%systemroot%\system32\cscript //B "%ProgramFiles(x86)%\Microsoft Office\office14\ospp.vbs" /setprt:1688
echo.
echo 正在啟用 Office...
%systemroot%\system32\cscript "%ProgramFiles(x86)%\Microsoft Office\office14\ospp.vbs" /act
echo.
echo 處理完成!
echo.
echo 按任意鍵返回主選單
pause>nul
Goto menu
REM ----- Office 2010 -----


REM ----- Office 2013 -----
:office15
echo -------------------
echo 已選擇 Office 2013
echo -------------------
echo.
if NOT EXIST "%ProgramFiles%\Microsoft Office\office15\ospp.vbs" goto office15x86
echo 正在設定 KMS Server...
%systemroot%\system32\cscript //B "%ProgramFiles%\Microsoft Office\office15\ospp.vbs" /sethst:KMS_IP_ADDRESS
%systemroot%\system32\cscript //B "%ProgramFiles%\Microsoft Office\office15\ospp.vbs" /setprt:1688
echo.
echo 正在啟用 Office...
%systemroot%\system32\cscript "%ProgramFiles%\Microsoft Office\office15\ospp.vbs" /act
echo.
echo 處理完成!
echo.
echo 按任意鍵返回主選單
pause>nul
Goto menu

:office15x86
cls
echo -------------------
echo 已選擇 Office 2013
echo -------------------
echo.
if NOT EXIST "%ProgramFiles(x86)%\Microsoft Office\office15\ospp.vbs" goto error
echo 正在設定 KMS Server...
%systemroot%\system32\cscript //B "%ProgramFiles(x86)%\Microsoft Office\office15\ospp.vbs" /sethst:KMS_IP_ADDRESS
%systemroot%\system32\cscript //B "%ProgramFiles(x86)%\Microsoft Office\office15\ospp.vbs" /setprt:1688
echo.
echo 正在啟用 Office...
%systemroot%\system32\cscript "%ProgramFiles(x86)%\Microsoft Office\office15\ospp.vbs" /act
echo.
echo 處理完成!
echo.
echo 按任意鍵返回主選單
pause>nul
Goto menu
REM ----- Office 2013 -----


REM ----- Office 2016 or 2019 -----
:office16
echo ------------------------
echo 已選擇 Office 2016/2019
echo ------------------------
echo.
if NOT EXIST "%ProgramFiles%\Microsoft Office\office16\ospp.vbs" goto office16x86
echo 正在設定 KMS Server...
%systemroot%\system32\cscript //B "%ProgramFiles%\Microsoft Office\office16\ospp.vbs" /sethst:KMS_IP_ADDRESS
%systemroot%\system32\cscript //B "%ProgramFiles%\Microsoft Office\office16\ospp.vbs" /setprt:1688
echo.
echo 正在啟用 Office...
%systemroot%\system32\cscript "%ProgramFiles%\Microsoft Office\office16\ospp.vbs" /act
echo.
echo 處理完成!
echo.
echo 按任意鍵返回主選單
pause>nul
Goto menu

:office16x86
cls
echo ------------------------
echo 已選擇 Office 2016/2019
echo ------------------------
echo.
if NOT EXIST "%ProgramFiles(x86)%\Microsoft Office\office16\ospp.vbs" goto error
echo 正在設定 KMS Server...
%systemroot%\system32\cscript //B "%ProgramFiles(x86)%\Microsoft Office\office16\ospp.vbs" /sethst:KMS_IP_ADDRESS
%systemroot%\system32\cscript //B "%ProgramFiles(x86)%\Microsoft Office\office16\ospp.vbs" /setprt:1688
echo.
echo 正在啟用 Office...
%systemroot%\system32\cscript "%ProgramFiles(x86)%\Microsoft Office\office16\ospp.vbs" /act
echo.
echo 處理完成!
echo.
echo 按任意鍵返回主選單
pause>nul
Goto menu
REM ----- Office 2016 or 2019 -----