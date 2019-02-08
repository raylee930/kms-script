@echo off 
color F0
title 多合一 Windows KMS 啟用小工具 By.Ray Ver.1902.08.2313

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
echo -------------------------------------
echo     多合一 Windows KMS 啟用小工具
echo          版本：1902.08.2313
echo -------------------------------------
echo 請選擇 Windows 版本:
echo.
echo   1.Windows Vista 商用版
echo   2.Windows Vista 企業版
echo   3.Windows 7 專業版
echo   4.Windows 7 企業版
echo   5.Windows 8 專業版
echo   6.Windows 8 企業版
echo   7.Windows 8.1 專業版
echo   8.Windows 8.1 企業版
echo   9.Windows 10 教育版
echo  10.Windows 10 專業教育版
echo  11.Windows 10 專業版
echo  12.Windows 10 專業版工作站
echo  13.Windows 10 企業版
echo  14.Windows 10 企業版 LTSB 2015
echo  15.Windows 10 企業版 LTSB 2016
echo  16.Windows 10 企業版 LTSC 2019
echo  17.Windows Server 2008 Standard
echo  18.Windows Server 2008 Enterprise
echo  19.Windows Server 2008 Datacenter
echo  20.Windows Server 2008 R2 Standard
echo  21.Windows Server 2008 R2 Enterprise
echo  22.Windows Server 2008 R2 Datacenter
echo  23.Windows Server 2012
echo  24.Windows Server 2012 Server Standard
echo  25.Windows Server 2012 Datacenter
echo  26.Windows Server 2012 R2 Standard
echo  27.Windows Server 2012 R2 Datacenter
echo  28.Windows Server 2012 R2 Essentials
echo  29.Windows Server 2016 Standard
echo  30.Windows Server 2016 Datacenter
echo  31.Windows Server 2016 Essentials
echo  32.Windows Server 2019 Standard
echo  33.Windows Server 2019 Datacenter
echo  34.Windows Server 2019 Essentials
echo.
echo  i.查看目前啟用資訊
echo  r.清除啟用資訊
echo  q.結束
echo.
echo.
SET select=
SET /P select=請選擇：
cls 

IF /I '%select%'=='i' GOTO info
IF /I '%select%'=='r' GOTO reset
IF /I '%select%'=='q' GOTO quit
IF /I '%select%'=='1' GOTO vistabus
IF /I '%select%'=='2' GOTO vistaent
IF /I '%select%'=='3' GOTO win7pro
IF /I '%select%'=='4' GOTO win7ent
IF /I '%select%'=='5' GOTO win8pro
IF /I '%select%'=='6' GOTO win8ent
IF /I '%select%'=='7' GOTO win81pro
IF /I '%select%'=='8' GOTO win81ent
IF /I '%select%'=='9' GOTO win10edu
IF /I '%select%'=='10' GOTO win10proedu
IF /I '%select%'=='11' GOTO win10pro
IF /I '%select%'=='12' GOTO win10prows
IF /I '%select%'=='13' GOTO win10ent
IF /I '%select%'=='14' GOTO win10ltsb2015
IF /I '%select%'=='15' GOTO win10ltsb2016
IF /I '%select%'=='16' GOTO win10ltsc2019
IF /I '%select%'=='17' GOTO win2008std
IF /I '%select%'=='18' GOTO win2008ent
IF /I '%select%'=='19' GOTO win2008dc
IF /I '%select%'=='20' GOTO win2008r2std
IF /I '%select%'=='21' GOTO win2008r2ent
IF /I '%select%'=='22' GOTO win2008r2dc
IF /I '%select%'=='23' GOTO win2012
IF /I '%select%'=='24' GOTO win2012std
IF /I '%select%'=='25' GOTO win2012dc
IF /I '%select%'=='26' GOTO win2012r2std
IF /I '%select%'=='27' GOTO win2012r2dc
IF /I '%select%'=='28' GOTO win2012r2ess
IF /I '%select%'=='29' GOTO win2016std
IF /I '%select%'=='30' GOTO win2016dc
IF /I '%select%'=='31' GOTO win2016ess
IF /I '%select%'=='32' GOTO win2019std
IF /I '%select%'=='33' GOTO win2019dc
IF /I '%select%'=='34' GOTO win2019ess

echo 選擇錯誤，請重新選擇！ 
echo 按任意鍵返回主選單
pause>nul
Goto menu

:activation
echo.
echo 正在設定 KMS Server...
%systemroot%\system32\cscript //B %systemroot%\system32\slmgr.vbs /skms KMS_IP_ADDRESS
echo.
echo 正在啟用 Windows...
%systemroot%\system32\cscript %systemroot%\system32\slmgr.vbs /ato
echo.
echo 按任意鍵返回主選單
pause>nul
Goto menu

:quit
exit

:info
%systemroot%\system32\cscript %systemroot%\system32\slmgr.vbs /dlv
echo.
echo 按任意鍵返回主選單
pause>nul
Goto menu

:reset
%systemroot%\system32\cscript %systemroot%\system32\slmgr.vbs /upk
%systemroot%\system32\cscript %systemroot%\system32\slmgr.vbs /ckms
echo.
echo 按任意鍵返回主選單
pause>nul
Goto menu


:vistabus
echo ----------------------------
echo 已選擇 Windows Vista 商用版
echo ----------------------------
echo.
echo 正在安裝產品金鑰...
%systemroot%\system32\cscript %systemroot%\system32\slmgr.vbs /ipk YFKBB-PQJJV-G996G-VWGXY-2V3X8
Goto activation


:vistaent
echo ----------------------------
echo 已選擇 Windows Vista 企業版
echo ----------------------------
echo.
echo 正在安裝產品金鑰...
%systemroot%\system32\cscript %systemroot%\system32\slmgr.vbs /ipk VKK3X-68KWM-X2YGT-QR4M6-4BWMV
Goto activation

:win7pro
echo ------------------------
echo 已選擇 Windows 7 專業版
echo ------------------------
echo.
echo 正在安裝產品金鑰...
%systemroot%\system32\cscript %systemroot%\system32\slmgr.vbs /ipk FJ82H-XT6CR-J8D7P-XQJJ2-GPDD4
Goto activation

:win7ent
echo ------------------------
echo 已選擇 Windows 7 企業版
echo ------------------------
echo.
echo 正在安裝產品金鑰...
%systemroot%\system32\cscript %systemroot%\system32\slmgr.vbs /ipk 33PXH-7Y6KF-2VJC9-XBBR8-HVTHH
Goto activation

:win8pro
echo ------------------------
echo 已選擇 Windows 8 專業版
echo ------------------------
echo.
echo 正在安裝產品金鑰...
%systemroot%\system32\cscript %systemroot%\system32\slmgr.vbs /ipk NG4HW-VH26C-733KW-K6F98-J8CK4
Goto activation

:win8ent
echo ------------------------
echo 已選擇 Windows 8 企業版
echo ------------------------
echo.
echo 正在安裝產品金鑰...
%systemroot%\system32\cscript %systemroot%\system32\slmgr.vbs /ipk 32JNW-9KQ84-P47T8-D8GGY-CWCK7
Goto activation

:win81pro
echo --------------------------
echo 已選擇 Windows 8.1 專業版
echo --------------------------
echo.
echo 正在安裝產品金鑰...
%systemroot%\system32\cscript %systemroot%\system32\slmgr.vbs /ipk GCRJD-8NW9H-F2CDX-CCM8D-9D6T9
Goto activation

:win81ent
echo --------------------------
echo 已選擇 Windows 8.1 企業版
echo --------------------------
echo.
echo 正在安裝產品金鑰...
%systemroot%\system32\cscript %systemroot%\system32\slmgr.vbs /ipk MHF9N-XY6XB-WVXMC-BTDCT-MKKG7
Goto activation

:win10edu
echo -------------------------
echo 已選擇 Windows 10 教育版
echo -------------------------
echo.
echo 正在安裝產品金鑰...
%systemroot%\system32\cscript %systemroot%\system32\slmgr.vbs /ipk NW6C2-QMPVW-D7KKK-3GKT6-VCFB2
Goto activation

:win10proedu
echo -----------------------------
echo 已選擇 Windows 10 專業教育版
echo -----------------------------
echo.
echo 正在安裝產品金鑰...
%systemroot%\system32\cscript %systemroot%\system32\slmgr.vbs /ipk 6TP4R-GNPTD-KYYHQ-7B7DP-J447Y
Goto activation

:win10pro
echo -------------------------
echo 已選擇 Windows 10 專業版
echo -------------------------
echo.
echo 正在安裝產品金鑰...
%systemroot%\system32\cscript %systemroot%\system32\slmgr.vbs /ipk W269N-WFGWX-YVC9B-4J6C9-T83GX
Goto activation

:win10prows
echo -------------------------------
echo 已選擇 Windows 10 專業版工作站
echo -------------------------------
echo.
echo 正在安裝產品金鑰...
%systemroot%\system32\cscript %systemroot%\system32\slmgr.vbs /ipk NRG8B-VKK3Q-CXVCJ-9G2XF-6Q84J
Goto activation

:win10ent
echo -------------------------
echo 已選擇 Windows 10 企業版
echo -------------------------
echo.
echo 正在安裝產品金鑰...
%systemroot%\system32\cscript %systemroot%\system32\slmgr.vbs /ipk NPPR9-FWDCX-D2C8J-H872K-2YT43
Goto activation

:win10ltsb2015
echo -----------------------------------
echo 已選擇 Windows 10 企業版 LTSB 2015
echo -----------------------------------
echo.
echo 正在安裝產品金鑰...
%systemroot%\system32\cscript %systemroot%\system32\slmgr.vbs /ipk WNMTR-4C88C-JK8YV-HQ7T2-76DF9
Goto activation

:win10ltsb2016
echo -----------------------------------
echo 已選擇 Windows 10 企業版 LTSB 2016
echo -----------------------------------
echo.
echo 正在安裝產品金鑰...
%systemroot%\system32\cscript %systemroot%\system32\slmgr.vbs /ipk DCPHK-NFMTC-H88MJ-PFHPY-QJ4BJ
Goto activation

:win10ltsc2019
echo -----------------------------------
echo 已選擇 Windows 10 企業版 LTSC 2019
echo -----------------------------------
echo.
echo 正在安裝產品金鑰...
%systemroot%\system32\cscript %systemroot%\system32\slmgr.vbs /ipk M7XTQ-FN8P6-TTKYV-9D4CC-J462D
Goto activation

:win2008std
echo ------------------------------------
echo 已選擇 Windows Server 2008 Standard
echo ------------------------------------
echo.
echo 正在安裝產品金鑰...
%systemroot%\system32\cscript %systemroot%\system32\slmgr.vbs /ipk TM24T-X9RMF-VWXK6-X8JC9-BFGM2
Goto activation

:win2008ent
echo --------------------------------------
echo 已選擇 Windows Server 2008 Enterprise
echo --------------------------------------
echo.
echo 正在安裝產品金鑰...
%systemroot%\system32\cscript %systemroot%\system32\slmgr.vbs /ipk YQGMW-MPWTJ-34KDK-48M3W-X4Q6V
Goto activation

:win2008dc
echo --------------------------------------
echo 已選擇 Windows Server 2008 Datacenter
echo --------------------------------------
echo.
echo 正在安裝產品金鑰...
%systemroot%\system32\cscript %systemroot%\system32\slmgr.vbs /ipk 7M67G-PC374-GR742-YH8V4-TCBY3
Goto activation

:win2008r2std
echo ---------------------------------------
echo 已選擇 Windows Server 2008 R2 Standard
echo ---------------------------------------
echo.
echo 正在安裝產品金鑰...
%systemroot%\system32\cscript %systemroot%\system32\slmgr.vbs /ipk YC6KT-GKW9T-YTKYR-T4X34-R7VHC
Goto activation

:win2008r2ent
echo -----------------------------------------
echo 已選擇 Windows Server 2008 R2 Enterprise
echo -----------------------------------------
echo.
echo 正在安裝產品金鑰...
%systemroot%\system32\cscript %systemroot%\system32\slmgr.vbs /ipk 489J6-VHDMP-X63PK-3K798-CPX3Y
Goto activation

:win2008r2dc
echo -----------------------------------------
echo 已選擇 Windows Server 2008 R2 Datacenter
echo -----------------------------------------
echo.
echo 正在安裝產品金鑰...
%systemroot%\system32\cscript %systemroot%\system32\slmgr.vbs /ipk 74YFP-3QFB3-KQT8W-PMXWJ-7M648
Goto activation

:win2012
echo ---------------------------
echo 已選擇 Windows Server 2012
echo ---------------------------
echo.
echo 正在安裝產品金鑰...
%systemroot%\system32\cscript %systemroot%\system32\slmgr.vbs /ipk BN3D2-R7TKB-3YPBD-8DRP2-27GG4
Goto activation

:win2012std
echo ------------------------------------
echo 已選擇 Windows Server 2012 Standard
echo ------------------------------------
echo.
echo 正在安裝產品金鑰...
%systemroot%\system32\cscript %systemroot%\system32\slmgr.vbs /ipk XC9B7-NBPP2-83J2H-RHMBY-92BT4
Goto activation

:win2012dc
echo --------------------------------------
echo 已選擇 Windows Server 2012 Datacenter
echo --------------------------------------
echo.
echo 正在安裝產品金鑰...
%systemroot%\system32\cscript %systemroot%\system32\slmgr.vbs /ipk 48HP8-DN98B-MYWDG-T2DCC-8W83P
Goto activation

:win2012r2std
echo ---------------------------------------
echo 已選擇 Windows Server 2012 R2 Standard
echo ---------------------------------------
echo.
echo 正在安裝產品金鑰...
%systemroot%\system32\cscript %systemroot%\system32\slmgr.vbs /ipk D2N9P-3P6X9-2R39C-7RTCD-MDVJX
Goto activation

:win2012r2dc
echo -----------------------------------------
echo 已選擇 Windows Server 2012 R2 Datacenter
echo -----------------------------------------
echo.
echo 正在安裝產品金鑰...
%systemroot%\system32\cscript %systemroot%\system32\slmgr.vbs /ipk W3GGN-FT8W3-Y4M27-J84CP-Q3VJ9
Goto activation

:win2012r2ess
echo --------------------------------------
echo 已選擇 Windows Server 2012 Essentials
echo --------------------------------------
echo.
echo 正在安裝產品金鑰...
%systemroot%\system32\cscript %systemroot%\system32\slmgr.vbs /ipk KNC87-3J2TX-XB4WP-VCPJV-M4FWM
Goto activation

:win2016std
echo ------------------------------------
echo 已選擇 Windows Server 2016 Standard
echo ------------------------------------
echo.
echo 正在安裝產品金鑰...
%systemroot%\system32\cscript %systemroot%\system32\slmgr.vbs /ipk WC2BQ-8NRM3-FDDYY-2BFGV-KHKQY
Goto activation

:win2016dc
echo --------------------------------------
echo 已選擇 Windows Server 2016 Datacenter
echo --------------------------------------
echo.
echo 正在安裝產品金鑰...
%systemroot%\system32\cscript %systemroot%\system32\slmgr.vbs /ipk CB7KF-BWN84-R7R2Y-793K2-8XDDG
Goto activation

:win2016ess
echo --------------------------------------
echo 已選擇 Windows Server 2016 Essentials
echo --------------------------------------
echo.
echo 正在安裝產品金鑰...
%systemroot%\system32\cscript %systemroot%\system32\slmgr.vbs /ipk JCKRF-N37P4-C2D82-9YXRT-4M63B
Goto activation

:win2019std
echo ------------------------------------
echo 已選擇 Windows Server 2019 Standard
echo ------------------------------------
echo.
echo 正在安裝產品金鑰...
%systemroot%\system32\cscript %systemroot%\system32\slmgr.vbs /ipk N69G4-B89J2-4G8F4-WWYCC-J464C
Goto activation

:win2019dc
echo --------------------------------------
echo 已選擇 Windows Server 2019 Datacenter
echo --------------------------------------
echo.
echo 正在安裝產品金鑰...
%systemroot%\system32\cscript %systemroot%\system32\slmgr.vbs /ipk WMDGN-G9PQG-XVVXX-R3X43-63DFG
Goto activation

:win2019ess
echo --------------------------------------
echo 已選擇 Windows Server 2019 Essentials
echo --------------------------------------
echo.
echo 正在安裝產品金鑰...
%systemroot%\system32\cscript %systemroot%\system32\slmgr.vbs /ipk WVDHN-86M7X-466P6-VHXV7-YY726
Goto activation