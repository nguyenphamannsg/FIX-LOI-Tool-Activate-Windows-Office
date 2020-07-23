@echo off
title Tool upgrade tu Windows 10 Home Len Windows 10 Pro - Copyright (C) NguyenPhamAn-LDSSK-BeCo Nam. All rights reserved.
mode con: cols=122 lines=40
chcp 65001 >nul
>nul 2>&1 "%SYSTEMROOT%\system32\cacls.exe" "%SYSTEMROOT%\system32\config\system"
if '%errorlevel%' NEQ '0' (
    echo  Run CMD as Administrator...
    goto goUAC 
) else (
 goto goADMIN )

:goUAC
    echo Set UAC = CreateObject^("Shell.Application"^) > "%temp%\getadmin.vbs"
    set params = %*:"=""
    echo UAC.ShellExecute "cmd.exe", "/c %~s0 %params%", "", "runas", 1 >> "%temp%\getadmin.vbs"
    "%temp%\getadmin.vbs"
    del "%temp%\getadmin.vbs"
    exit /B

:goADMIN
    pushd "%CD%"
    CD /D "%~dp0"
color f3
cls
@echo off
@echo 1.Cong cu upgrade tu Windows 10 Home Len Windows 10 Pro
@echo 2. Nhan phim bat ki de bat dau upgrade
pause
sc config LicenseManager start= auto & net start LicenseManager
sc config wuauserv start= auto & net start wuauserv
changepk.exe /productkey VK7JG-NPHTM-C97JM-9MPGT-3V66T
@echo Vui long doi qua trinh upgrade den 100%
timeout 3
exit