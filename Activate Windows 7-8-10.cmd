chcp 65001 >nul
@echo off
title ACTIVATE WINDOWS 7 8 8.1 10 By Phone - Copyright (C) NguyenPhamAn-LDSSK-BeCo Nam-TranVinhTrung-and some other members. All rights reserved.
mode con: cols=115 lines=40
::--------------------------------------------------------------------------------------------------------------------------------------------------------
:: Elevating UAC Administrator Privileges
>nul 2>&1 "%SYSTEMROOT%\system32\cacls.exe" "%SYSTEMROOT%\system32\config\system"
if "%errorlevel%" NEQ "0" (
	echo: Set UAC = CreateObject^("Shell.Application"^) > "%temp%\getadmin.vbs"
ECHO RUN CMD AS ADMIN............................
	echo: UAC.ShellExecute "%~s0", "", "", "runas", 1 >> "%temp%\getadmin.vbs"
	"%temp%\getadmin.vbs" &	exit 
)
if exist "%temp%\getadmin.vbs" del /f /q "%temp%\getadmin.vbs"
cls

color f3
ver
setlocal EnableExtensions EnableDelayedExpansion
if exist "%ProgramFiles(x86)%" (set bit=64-bit) else (set bit=32-bit)
for /f "tokens=2*" %%c in ('"reg query "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v ProductName" 2^>nul') do set ProductName=%%d %bit%
cls
@echo.
:main
echo                                                %ProductName% 
@echo 1. Dang kiem tra ban quyen
cscript //nologo %windir%\system32\slmgr.vbs /xpr |findstr "permanently"  >nul
if %errorlevel%==0  (
@echo                                 === WINDOWS DA KICH HOAT BAN QUYEN VINH VIEN ===
echo Nhan Phim Bat Ki De Tiep Tuc
pause >nul
Exit
) else (
@echo                              === WINDOWS CHUA DUOC KICH HOAT BAN QUYEN VINH VIEN ===
@echo.
@echo Nhan phim bat ky de tiep tuc kich hoat...
pause >nul
goto 222
)
:222
@echo.
@echo 						Ban can gi???
@echo				1.Khoi Phuc Ban Quyen Da Sao Luu ( file ActivateBackup )
@echo				2.Tiep Tuc Kich Hoat Ban Quyen Windows
choice /c:12 /n /m ">_ [1,2] : "
if errorlevel 2 goto :begin                         
if errorlevel 1 goto :resbq
:remkey
@echo Dang Go Key Windows : 
sc config Winmgmt start=demand 
net start Winmgmt 
sc config LicenseManager start= auto 
net start LicenseManager 
sc config wuauserv start= auto 
net start wuauserv 
cd %windir%\system32 
cscript slmgr.vbs /cpky 
cscript slmgr.vbs /upk  
Cscript slmgr.vbs /ckms
cls & echo off & goto :begin & cls 
:begin
@echo.
set /p key= 2. Nhap Key : 
@echo.
@echo 3. Dang go bo key cu ra khoi may va xoa key trong Registry sau do nap key vao may xuat ra Installations ID
cd %windir%\system32
@echo off &mode con: cols=20 lines=2
set k1=%key%
Cscript slmgr.vbs /ipk %k1%
@echo off
@mode con: cols=100 lines=30
echo Dang tu dong get CID de kich hoat windows
for /f "tokens=3" %%i in ('cscript slmgr.vbs /dti') do set MyIID=%%i
for /f "tokens=*" %%b in ('powershell -Command "$req = [System.Net.WebRequest]::Create('https://winoffice.org/public-api/get-cid?iid=%MyIID%&token=gAAAAABfGobY9lyp_7XAz5FgrJSaf--0EUizjmFUkl2eI1EhC994zZGlkXpYMBrtJMJ_t0hVIB_KqCT7A1R4v-BNW_NOWwezv7IprnhIO4lceYJxAusOL-E=');$resp = New-Object System.IO.StreamReader $req.GetResponse().GetResponseStream(); $resp.ReadToEnd()"') do set ACID=%%b
set CID=%ACID:~1,48%
echo %CID% 
Cscript slmgr.vbs /atp %CID% & cscript slmgr.vbs /ato
@echo.
@echo Nhan phim bat ky de tiep tuc
pause >nul  
cscript //nologo %windir%\system32\slmgr.vbs /xpr |findstr "permanently" >nul
if %errorlevel%==0 (
@echo                                 === WINDOWS DA KICH HOAT BAN QUYEN VINH VIEN ===
@echo === Da kich hoat ban quyen VINH VIEN ===
echo Vui long backup ban quyen de su dung ca doi ( Tool backup o trong muc tool other)
Set /p "supporter=  - Nhap ten Supporter( neu ban khong phai la Supporter thi nhan Enter de bo qua) : "
@Echo Dang hien thi trang thai kich hoat cua Windows
echo Supporter:%supporter% >k2.txt & echo Key:%k1% >>k2.txt & echo IID:%MyIID% >>k2.txt & echo CID:%CID% >>k2.txt & cscript slmgr.vbs /dli >>k2.txt & cscript slmgr.vbs /xpr >>k2.txt & start k2.txt 
@echo Dang Backup Ban Quyen Windows
rd /s /q "%~dp0\ActivateBackup" >nul 2>&1
md "%~dp0\ActivateBackup\Windows" >nul 2>&1
if exist "%windir%\ServiceProfiles\NetworkService\AppData\Roaming\Microsoft\SoftwareProtectionPlatform\tokens.dat" xcopy "%windir%\ServiceProfiles\NetworkService\AppData\Roaming\Microsoft\SoftwareProtectionPlatform\*.*" "%~dp0\ActivateBackup\Windows" /s>nul
if exist "%windir%\System32\spp\store\2.0\tokens.dat" xcopy "%windir%\System32\spp\store\2.0\*.*" "%~dp0\ActivateBackup\Windows" /s>nul
if exist "C:\ProgramData\Microsoft\OfficeSoftwareProtectionPlatform\tokens.dat" (
md "%~dp0\ActivateBackup\Office" >nul
xcopy "C:\ProgramData\Microsoft\OfficeSoftwareProtectionPlatform" "%~dp0\ActivateBackup\Office"/s >nul
@Echo Done!^ & goto :exit ) else
(
@echo    ==WINDOWS CHUA DUOC KICH HOAT BAN QUYEN VINH VIEN==
@echo.      =DO LOI STEP 3-HOAC LOI KHONG MONG MUON KHAC==
@echo Nhan phim bat ky de thu kich hoat lai...
pause >nul
cls
goto begin
:resbq
if exist "%~dp0\ActivateBackup\Windows\tokens.dat" (
net stop sppsvc >nul 2>&1
if exist "%windir%\ServiceProfiles\NetworkService\AppData\Roaming\Microsoft\SoftwareProtectionPlatform\tokens.dat" (
xcopy "%~dp0\ActivateBackup\Windows" "%windir%\ServiceProfiles\NetworkService\AppData\Roaming\Microsoft\SoftwareProtectionPlatform" /s /y>nul
goto:nextrestore
)
if exist "%windir%\System32\spp\store\2.0\tokens.dat" (
xcopy "%~dp0\ActivateBackup\Windows" "%windir%\System32\spp\store\2.0" /s /y>nul
goto:nextrestore
)
echo Co ve nhu ban dang reset file tokens.dat
echo Chi tiet hon xem tai https://support.microsoft.com/en-us/help/2736303
echo Nhan phim bat ky de tiep tuc...
goto main
)
echo Khong tim thay bat ky 1 ban sao luu, vui long kiem tra lai
echo Nhan phim bat ky de tiep tuc...
pause>nul
cls & goto main
:nextrestore
echo Dang khoi phuc ban quyen cua ban...
if exist "%~dp0\ActivateBackup\Office\tokens.dat" (
if exist "C:\ProgramData\Microsoft\OfficeSoftwareProtectionPlatform" (
net stop osppsvc >nul 2>&1
xcopy "%~dp0\ActivateBackup\Office" "C:\ProgramData\Microsoft\OfficeSoftwareProtectionPlatform" /s /y >nul
))
echo Dang bat cac service...
net start sppsvc >nul 2>&1
net start osppsvc >nul 2>&1
sc config LicenseManager start= auto >nul 2>&1
net start LicenseManager >nul 2>&1
sc config wuauserv start= auto >nul 2>&1
net start wuauserv >nul 2>&1
echo Dang kich hoat lai Windows cua ban...
cscript //nologo %windir%\system32\slmgr.vbs /dlv | find "LICENSED" >nul
if %errorlevel%==0 (echo Kich hoat thanh cong) ELSE (echo Kich hoat khong thanh cong)
echo.
for %%v in (4,5,6) do (if exist "%ProgramFiles%\Microsoft Office\Office1%%v\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office1%%v") & if exist "%ProgramFiles% (x86)\Microsoft Office\Office1%%v\ospp.vbs" cd /d "%ProgramFiles% (x86)\Microsoft Office\Office1%%v")
if exist "ospp.vbs" (
echo Dang kich hoat lai Office cua ban...
cscript ospp.vbs /dstatus | find "LICENSED" >nul
if %errorlevel%==0 (echo Kich hoat thanh cong) ELSE (echo Kich hoat khong thanh cong)
echo.
)
echo Nhan phim bat ky de tiep tuc...
pause>nul
goto exit
:exit
@echo.
@echo =========================================================
@echo [  Cam on ban da su dung Tool Activate Windows & Office  ]
@echo [     Thanks for using Tool Activate Windows & Office    ]
@echo =========================================================

timeout 3
exit