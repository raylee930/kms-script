@echo off
color F0
title �h�X�@ Office KMS �ҥΤp�u�� By.Ray Ver.1905.03.1110

:-------------------------------------
IF "%PROCESSOR_ARCHITECTURE%" EQU "amd64" (
>nul 2>&1 "%SYSTEMROOT%\SysWOW64\cacls.exe" "%SYSTEMROOT%\SysWOW64\config\system"
) ELSE (
>nul 2>&1 "%SYSTEMROOT%\system32\cacls.exe" "%SYSTEMROOT%\system32\config\system"
)

if '%errorlevel%' NEQ '0' (
    echo �ШD�t�κ޲z���v��...
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
echo     �h�X�@ Office KMS �ҥΤp�u��
echo          �����G1905.03.1110
echo ------------------------------------
echo �п�� Office ����:
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
echo  q.����
echo.
echo.
SET select=
SET /P select=�п�ܡG
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

echo ��ܿ��~�A�Э��s��ܡI 
echo �����N���^�D���
pause>nul
Goto menu

:quit
exit

:error
cls
echo �t�Χ䤣����w���|�A�нT�{�w��ܥ��T�� Office ����
echo.
echo �����N���^�D���
pause>nul
Goto menu


REM ----- Office 2010 Standard -----
:office2010std
echo ----------------------------
echo �w��� Office 2010 Standard
echo ----------------------------
echo.
if NOT EXIST "%ProgramFiles%\Microsoft Office\office14\ospp.vbs" goto office2010stdx86
echo ���b�]�w KMS Server...
%systemroot%\system32\cscript //B "%ProgramFiles%\Microsoft Office\office14\ospp.vbs" /sethst:%KmsServer%
%systemroot%\system32\cscript //B "%ProgramFiles%\Microsoft Office\office14\ospp.vbs" /setprt:%KmsPort%
echo.
echo ���b�w�˲��~���_...
%systemroot%\system32\cscript "%ProgramFiles%\Microsoft Office\office14\ospp.vbs" /inpkey:V7QKV-4XVVR-XYV4D-F7DFM-8R6BM
echo.
echo ���b�ҥ� Office...
%systemroot%\system32\cscript "%ProgramFiles%\Microsoft Office\office14\ospp.vbs" /act
echo.
echo �B�z����!
echo.
echo �����N���^�D���
pause>nul
Goto menu

:office2010stdx86
cls
echo ----------------------------
echo �w��� Office 2010 Standard
echo ----------------------------
echo.
if NOT EXIST "%ProgramFiles(x86)%\Microsoft Office\office14\ospp.vbs" goto error
echo ���b�]�w KMS Server...
%systemroot%\system32\cscript //B "%ProgramFiles(x86)%\Microsoft Office\office14\ospp.vbs" /sethst:%KmsServer%
%systemroot%\system32\cscript //B "%ProgramFiles(x86)%\Microsoft Office\office14\ospp.vbs" /setprt:%KmsPort%
echo.
echo ���b�w�˲��~���_...
%systemroot%\system32\cscript "%ProgramFiles(x86)%\Microsoft Office\office14\ospp.vbs" /inpkey:V7QKV-4XVVR-XYV4D-F7DFM-8R6BM
echo.
echo ���b�ҥ� Office...
%systemroot%\system32\cscript "%ProgramFiles(x86)%\Microsoft Office\office14\ospp.vbs" /act
echo.
echo �B�z����!
echo.
echo �����N���^�D���
pause>nul
Goto menu
REM ----- Office 2010 Standard -----

REM ----- Office 2010 ProPlus -----
:office2010proplus
echo -------------------------------------
echo �w��� Office 2010 Professional Plus
echo -------------------------------------
echo.
if NOT EXIST "%ProgramFiles%\Microsoft Office\office14\ospp.vbs" goto office2010ppx86
echo ���b�]�w KMS Server...
%systemroot%\system32\cscript //B "%ProgramFiles%\Microsoft Office\office14\ospp.vbs" /sethst:%KmsServer%
%systemroot%\system32\cscript //B "%ProgramFiles%\Microsoft Office\office14\ospp.vbs" /setprt:%KmsPort%
echo.
echo ���b�w�˲��~���_...
%systemroot%\system32\cscript "%ProgramFiles%\Microsoft Office\office14\ospp.vbs" /inpkey:VYBBJ-TRJPB-QFQRF-QFT4D-H3GVB
echo.
echo ���b�ҥ� Office...
%systemroot%\system32\cscript "%ProgramFiles%\Microsoft Office\office14\ospp.vbs" /act
echo.
echo �B�z����!
echo.
echo �����N���^�D���
pause>nul
Goto menu

:office2010ppx86
cls
echo -------------------------------------
echo �w��� Office 2010 Professional Plus
echo -------------------------------------
echo.
if NOT EXIST "%ProgramFiles(x86)%\Microsoft Office\office14\ospp.vbs" goto error
echo ���b�]�w KMS Server...
%systemroot%\system32\cscript //B "%ProgramFiles(x86)%\Microsoft Office\office14\ospp.vbs" /sethst:%KmsServer%
%systemroot%\system32\cscript //B "%ProgramFiles(x86)%\Microsoft Office\office14\ospp.vbs" /setprt:%KmsPort%
echo.
echo ���b�w�˲��~���_...
%systemroot%\system32\cscript "%ProgramFiles(x86)%\Microsoft Office\office14\ospp.vbs" /inpkey:VYBBJ-TRJPB-QFQRF-QFT4D-H3GVB
echo.
echo ���b�ҥ� Office...
%systemroot%\system32\cscript "%ProgramFiles(x86)%\Microsoft Office\office14\ospp.vbs" /act
echo.
echo �B�z����!
echo.
echo �����N���^�D���
pause>nul
Goto menu
REM ----- Office 2010 ProPlus -----


REM ----- Office 2013 Standard -----
:office2013std
echo ----------------------------
echo �w��� Office 2013 Standard
echo ----------------------------
echo.
if NOT EXIST "%ProgramFiles%\Microsoft Office\office15\ospp.vbs" goto office2013stdx86
echo ���b�]�w KMS Server...
%systemroot%\system32\cscript //B "%ProgramFiles%\Microsoft Office\office15\ospp.vbs" /sethst:%KmsServer%
%systemroot%\system32\cscript //B "%ProgramFiles%\Microsoft Office\office15\ospp.vbs" /setprt:%KmsPort%
echo.
echo ���b�w�˲��~���_...
%systemroot%\system32\cscript "%ProgramFiles%\Microsoft Office\office15\ospp.vbs" /inpkey:KBKQT-2NMXY-JJWGP-M62JB-92CD4
echo.
echo ���b�ҥ� Office...
%systemroot%\system32\cscript "%ProgramFiles%\Microsoft Office\office15\ospp.vbs" /act
echo.
echo �B�z����!
echo.
echo �����N���^�D���
pause>nul
Goto menu

:office2013stdx86
cls
echo ----------------------------
echo �w��� Office 2013 Standard
echo ----------------------------
echo.
if NOT EXIST "%ProgramFiles(x86)%\Microsoft Office\office15\ospp.vbs" goto error
echo ���b�]�w KMS Server...
%systemroot%\system32\cscript //B "%ProgramFiles(x86)%\Microsoft Office\office15\ospp.vbs" /sethst:%KmsServer%
%systemroot%\system32\cscript //B "%ProgramFiles(x86)%\Microsoft Office\office15\ospp.vbs" /setprt:%KmsPort%
echo.
echo ���b�w�˲��~���_...
%systemroot%\system32\cscript "%ProgramFiles(x86)%\Microsoft Office\office15\ospp.vbs" /inpkey:KBKQT-2NMXY-JJWGP-M62JB-92CD4
echo.
echo ���b�ҥ� Office...
%systemroot%\system32\cscript "%ProgramFiles(x86)%\Microsoft Office\office15\ospp.vbs" /act
echo.
echo �B�z����!
echo.
echo �����N���^�D���
pause>nul
Goto menu
REM ----- Office 2013 Standard -----

REM ----- Office 2013 ProPlus -----
:office2013proplus
echo -------------------------------------
echo �w��� Office 2013 Professional Plus
echo -------------------------------------
echo.
if NOT EXIST "%ProgramFiles%\Microsoft Office\office15\ospp.vbs" goto office2013ppx86
echo ���b�]�w KMS Server...
%systemroot%\system32\cscript //B "%ProgramFiles%\Microsoft Office\office15\ospp.vbs" /sethst:%KmsServer%
%systemroot%\system32\cscript //B "%ProgramFiles%\Microsoft Office\office15\ospp.vbs" /setprt:%KmsPort%
echo.
echo ���b�w�˲��~���_...
%systemroot%\system32\cscript "%ProgramFiles%\Microsoft Office\office15\ospp.vbs" /inpkey:YC7DK-G2NP3-2QQC3-J6H88-GVGXT
echo.
echo ���b�ҥ� Office...
%systemroot%\system32\cscript "%ProgramFiles%\Microsoft Office\office15\ospp.vbs" /act
echo.
echo �B�z����!
echo.
echo �����N���^�D���
pause>nul
Goto menu

:office2013ppx86
cls
echo -------------------------------------
echo �w��� Office 2013 Professional Plus
echo -------------------------------------
echo.
if NOT EXIST "%ProgramFiles(x86)%\Microsoft Office\office15\ospp.vbs" goto error
echo ���b�]�w KMS Server...
%systemroot%\system32\cscript //B "%ProgramFiles(x86)%\Microsoft Office\office15\ospp.vbs" /sethst:%KmsServer%
%systemroot%\system32\cscript //B "%ProgramFiles(x86)%\Microsoft Office\office15\ospp.vbs" /setprt:%KmsPort%
echo.
echo ���b�w�˲��~���_...
%systemroot%\system32\cscript "%ProgramFiles(x86)%\Microsoft Office\office15\ospp.vbs" /inpkey:YC7DK-G2NP3-2QQC3-J6H88-GVGXT
echo.
echo ���b�ҥ� Office...
%systemroot%\system32\cscript "%ProgramFiles(x86)%\Microsoft Office\office15\ospp.vbs" /act
echo.
echo �B�z����!
echo.
echo �����N���^�D���
pause>nul
Goto menu
REM ----- Office 2013 ProPlus -----


REM ----- Office 2016 Standard -----
:office2016std
echo ----------------------------
echo �w��� Office 2016 Standard
echo ----------------------------
echo.
if NOT EXIST "%ProgramFiles%\Microsoft Office\office16\ospp.vbs" goto office2016stdx86
echo ���b�]�w KMS Server...
%systemroot%\system32\cscript //B "%ProgramFiles%\Microsoft Office\office16\ospp.vbs" /sethst:%KmsServer%
%systemroot%\system32\cscript //B "%ProgramFiles%\Microsoft Office\office16\ospp.vbs" /setprt:%KmsPort%
echo.
echo ���b�w�˲��~���_...
%systemroot%\system32\cscript "%ProgramFiles%\Microsoft Office\office16\ospp.vbs" /inpkey:JNRGM-WHDWX-FJJG3-K47QV-DRTFM
echo.
echo ���b�ҥ� Office...
%systemroot%\system32\cscript "%ProgramFiles%\Microsoft Office\office16\ospp.vbs" /act
echo.
echo �B�z����!
echo.
echo �����N���^�D���
pause>nul
Goto menu

:office2016stdx86
cls
echo ----------------------------
echo �w��� Office 2016 Standard
echo ----------------------------
echo.
if NOT EXIST "%ProgramFiles(x86)%\Microsoft Office\office16\ospp.vbs" goto error
echo ���b�]�w KMS Server...
%systemroot%\system32\cscript //B "%ProgramFiles(x86)%\Microsoft Office\office16\ospp.vbs" /sethst:%KmsServer%
%systemroot%\system32\cscript //B "%ProgramFiles(x86)%\Microsoft Office\office16\ospp.vbs" /setprt:%KmsPort%
echo.
echo ���b�w�˲��~���_...
%systemroot%\system32\cscript "%ProgramFiles(x86)%\Microsoft Office\office16\ospp.vbs" /inpkey:JNRGM-WHDWX-FJJG3-K47QV-DRTFM
echo.
echo ���b�ҥ� Office...
%systemroot%\system32\cscript "%ProgramFiles(x86)%\Microsoft Office\office16\ospp.vbs" /act
echo.
echo �B�z����!
echo.
echo �����N���^�D���
pause>nul
Goto menu
REM ----- Office 2016 Standard -----

REM ----- Office 2016 Professional Plus -----
:office2016proplus
echo -------------------------------------
echo �w��� Office 2016 Professional Plus
echo -------------------------------------
echo.
if NOT EXIST "%ProgramFiles%\Microsoft Office\office16\ospp.vbs" goto office2016ppx86
echo ���b�]�w KMS Server...
%systemroot%\system32\cscript //B "%ProgramFiles%\Microsoft Office\office16\ospp.vbs" /sethst:%KmsServer%
%systemroot%\system32\cscript //B "%ProgramFiles%\Microsoft Office\office16\ospp.vbs" /setprt:%KmsPort%
echo.
echo ���b�w�˲��~���_...
%systemroot%\system32\cscript "%ProgramFiles%\Microsoft Office\office16\ospp.vbs" /inpkey:XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99
echo.
echo ���b�ҥ� Office...
%systemroot%\system32\cscript "%ProgramFiles%\Microsoft Office\office16\ospp.vbs" /act
echo.
echo �B�z����!
echo.
echo �����N���^�D���
pause>nul
Goto menu

:office2016ppx86
cls
echo -------------------------------------
echo �w��� Office 2016 Professional Plus
echo -------------------------------------
echo.
if NOT EXIST "%ProgramFiles(x86)%\Microsoft Office\office16\ospp.vbs" goto error
echo ���b�]�w KMS Server...
%systemroot%\system32\cscript //B "%ProgramFiles(x86)%\Microsoft Office\office16\ospp.vbs" /sethst:%KmsServer%
%systemroot%\system32\cscript //B "%ProgramFiles(x86)%\Microsoft Office\office16\ospp.vbs" /setprt:%KmsPort%
echo.
echo ���b�w�˲��~���_...
%systemroot%\system32\cscript "%ProgramFiles(x86)%\Microsoft Office\office16\ospp.vbs" /inpkey:XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99
echo.
echo ���b�ҥ� Office...
%systemroot%\system32\cscript "%ProgramFiles(x86)%\Microsoft Office\office16\ospp.vbs" /act
echo.
echo �B�z����!
echo.
echo �����N���^�D���
pause>nul
Goto menu
REM ----- Office 2016 Professional Plus -----


REM ----- Office 2019 Standard -----
:office2019std
echo ----------------------------
echo �w��� Office 2019 Standard
echo ----------------------------
echo.
if NOT EXIST "%ProgramFiles%\Microsoft Office\office16\ospp.vbs" goto office2019stdx86
echo ���b�]�w KMS Server...
%systemroot%\system32\cscript //B "%ProgramFiles%\Microsoft Office\office16\ospp.vbs" /sethst:%KmsServer%
%systemroot%\system32\cscript //B "%ProgramFiles%\Microsoft Office\office16\ospp.vbs" /setprt:%KmsPort%
echo.
echo ���b�w�˲��~���_...
%systemroot%\system32\cscript "%ProgramFiles%\Microsoft Office\office16\ospp.vbs" /inpkey:6NWWJ-YQWMR-QKGCB-6TMB3-9D9HK
echo.
echo ���b�ҥ� Office...
%systemroot%\system32\cscript "%ProgramFiles%\Microsoft Office\office16\ospp.vbs" /act
echo.
echo �B�z����!
echo.
echo �����N���^�D���
pause>nul
Goto menu

:office2019stdx86
cls
echo ----------------------------
echo �w��� Office 2019 Standard
echo ----------------------------
echo.
if NOT EXIST "%ProgramFiles(x86)%\Microsoft Office\office16\ospp.vbs" goto error
echo ���b�]�w KMS Server...
%systemroot%\system32\cscript //B "%ProgramFiles(x86)%\Microsoft Office\office16\ospp.vbs" /sethst:%KmsServer%
%systemroot%\system32\cscript //B "%ProgramFiles(x86)%\Microsoft Office\office16\ospp.vbs" /setprt:%KmsPort%
echo.
echo ���b�w�˲��~���_...
%systemroot%\system32\cscript "%ProgramFiles(x86)%\Microsoft Office\office16\ospp.vbs" /inpkey:6NWWJ-YQWMR-QKGCB-6TMB3-9D9HK
echo.
echo ���b�ҥ� Office...
%systemroot%\system32\cscript "%ProgramFiles(x86)%\Microsoft Office\office16\ospp.vbs" /act
echo.
echo �B�z����!
echo.
echo �����N���^�D���
pause>nul
Goto menu
REM ----- Office 2019 Standard -----

REM ----- Office 2019 Professional Plus -----
:office2019proplus
echo -------------------------------------
echo �w��� Office 2019 Professional Plus
echo -------------------------------------
echo.
if NOT EXIST "%ProgramFiles%\Microsoft Office\office16\ospp.vbs" goto office2019ppx86
echo ���b�]�w KMS Server...
%systemroot%\system32\cscript //B "%ProgramFiles%\Microsoft Office\office16\ospp.vbs" /sethst:%KmsServer%
%systemroot%\system32\cscript //B "%ProgramFiles%\Microsoft Office\office16\ospp.vbs" /setprt:%KmsPort%
echo.
echo ���b�w�˲��~���_...
%systemroot%\system32\cscript "%ProgramFiles%\Microsoft Office\office16\ospp.vbs" /inpkey:NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP
echo.
echo ���b�ҥ� Office...
%systemroot%\system32\cscript "%ProgramFiles%\Microsoft Office\office16\ospp.vbs" /act
echo.
echo �B�z����!
echo.
echo �����N���^�D���
pause>nul
Goto menu

:office2019ppx86
cls
echo -------------------------------------
echo �w��� Office 2019 Professional Plus
echo -------------------------------------
echo.
if NOT EXIST "%ProgramFiles(x86)%\Microsoft Office\office16\ospp.vbs" goto error
echo ���b�]�w KMS Server...
%systemroot%\system32\cscript //B "%ProgramFiles(x86)%\Microsoft Office\office16\ospp.vbs" /sethst:%KmsServer%
%systemroot%\system32\cscript //B "%ProgramFiles(x86)%\Microsoft Office\office16\ospp.vbs" /setprt:%KmsPort%
echo.
echo ���b�w�˲��~���_...
%systemroot%\system32\cscript "%ProgramFiles(x86)%\Microsoft Office\office16\ospp.vbs" /inpkey:NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP
echo.
echo ���b�ҥ� Office...
%systemroot%\system32\cscript "%ProgramFiles(x86)%\Microsoft Office\office16\ospp.vbs" /act
echo.
echo �B�z����!
echo.
echo �����N���^�D���
pause>nul
Goto menu
REM ----- Office 2019 Professional Plus -----