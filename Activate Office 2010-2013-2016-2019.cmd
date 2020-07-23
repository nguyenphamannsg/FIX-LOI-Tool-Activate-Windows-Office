chcp 65001 >nul
@echo off
title ACTIVATE OFFICE 2010-2013-2016-2019 By Phone - Copyright (C) by NguyenPhamAn-LDSSK-BeCo Nam-TranVinhTrung-and some other members. All rights reserved.
mode con: cols=122 lines=38
::--------------------------------------------------------------------------------------------------------------------------------------------------------
:: Elevating UAC Administrator Privileges
>nul 2>&1 "%SYSTEMROOT%\system32\cacls.exe" "%SYSTEMROOT%\system32\config\system"
if "%errorlevel%" NEQ "0" (
	echo: Set UAC = CreateObject^("Shell.Application"^) > "%temp%\getadmin.vbs"
@echo Run CMD as ADMIN.............	
echo: UAC.ShellExecute "%~s0", "", "", "runas", 1 >> "%temp%\getadmin.vbs"
	"%temp%\getadmin.vbs" &	exit 
)
if exist "%temp%\getadmin.vbs" del /f /q "%temp%\getadmin.vbs"
cls
color f3
for %%a in (4,5,6) do (if exist "%ProgramFiles%\Microsoft Office\Office1%%a\ospp.vbs" (cd /d "%ProgramFiles%\Microsoft Office\Office1%%a")
If exist "%ProgramFiles% (x86)\Microsoft Office\Office1%%a\ospp.vbs" (cd /d "%ProgramFiles% (x86)\Microsoft Office\Office1%%a"))
@echo. ^^Dang nhan Dien Office.....
for /f "tokens=5" %%b in ('cscript ospp.vbs /dstatus ^| findstr /b /c:"LICENSE NAME:"') do set Name=%%b
goto :main
:main
@Echo                                           %Name%
@Echo.
@Echo.
@Echo Dang kiem tra ban quyen Office
for %%a in (4,5,6) do (if exist "%ProgramFiles%\Microsoft Office\Office1%%a\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office1%%a" 
If exist "%ProgramFiles% (x86)\Microsoft Office\Office1%%a\ospp.vbs" cd /d "%ProgramFiles% (x86)\Microsoft Office\Office1%%a")
cscript OSPP.VBS /dstatus |findstr "LICENSED" >nul
if %errorlevel%==0  (
@echo                        ===OFFICE DA DUOC KICH HOAT BAN QUYEN VINH VIEN===
echo Ban co muon go key Office de tiep tuc kich hoat?
choice /c:1234 /n /m ">_         Nhan 1 de go key Office , Nhan 2 de thoat
if errorlevel 1 goto :remkey
if errorlevel 2 exit
) else ( 
@Echo                       === OFFICE CHUA DUOC KICH HOAT BAN QUYEN VINH VIEN===
goto begin
)
:remkey 
for /f "tokens=8" %%c in ('cscript //nologo OSPP.VBS /dstatus ^| findstr /b /c:"Last 5"') do (cscript //nologo ospp.vbs /unpkey:%%c)
cscript ospp.vbs /remhst
goto :begin
:begin
@echo.
Echo               			   Ban Dang Can???              
echo 				1.Activate Office 2010-2013-2016-2019
echo				2.Activate Office 2019 Volume( retail-volume )
echo     			3.Khoi phuc ban quyen da sao luu truoc do ( File ActivateBackup )
echo				4.Nhan 4 de thoat
choice /c:1234 /n /m ">_ [1,2] : "
if errorlevel 4 exit
if errorlevel 3 goto :resbq
if errorlevel 2 goto :act2019                          
if errorlevel 1 goto :act

:act
for %%a in (4,5,6) do (if exist "%ProgramFiles%\Microsoft Office\Office1%%a\ospp.vbs" (cd /d "%ProgramFiles%\Microsoft Office\Office1%%a")
If exist "%ProgramFiles% (x86)\Microsoft Office\Office1%%a\ospp.vbs" (cd /d "%ProgramFiles% (x86)\Microsoft Office\Office1%%a"))
set /p key= 1. Nhap Key : 
@echo.
@echo 2. Dang cai dat Key
for /f "tokens=8" %%c in ('cscript //nologo OSPP.VBS /dstatus ^| findstr /b /c:"Last 5"') do (cscript //nologo ospp.vbs /unpkey:%%c)
cscript OSPP.VBS /inpkey:%key%
@echo.
@echo 3. Dang Xuat Installion ID va Get CID kich hoat Office
for /f "tokens=8" %%i in ('cscript ospp.vbs /dinstid') do set MyIID=%%i
for /f "tokens=*" %%b in ('powershell -Command "$req = [System.Net.WebRequest]::Create('http://huyphung.com/public-api/get-cid?username=an8amc&password=@Ann123456&iid=%MyIID%');$resp = New-Object System.IO.StreamReader $req.GetResponse().GetResponseStream(); $resp.ReadToEnd()"') do set ACID=%%b
set CID=%ACID:~1,48%
echo CID cua ban %CID% 
@echo 5. Dang kich hoat ban quyen
Cscript ospp.vbs /actcid:%CID% & cscript ospp.vbs /act 
cscript OSPP.VBS /dstatus |findstr "LICENSED" >nul
if %errorlevel%==0  (
@echo   === Da kich hoat ban quyen VINH VIEN ===
Set /p "supporter=  - Nhap ten Supporter(neu ban khong phai la Supporter thi an enter) : "
echo Dang xuat status Office
echo Supporter:%supporter% >k1.txt & echo Key:%key% >>k1.txt & echo IID:%MyIID% >>k1.txt & echo CID:%CID% >>k1.txt & cscript ospp.vbs /dstatus >>k1.txt & start k1.txt & start winword & exit
@echo Dang Backup Ban Quyen Office
rd /s /q "%~dp0\ActivateBackup" >nul 2>&1
md "%~dp0\ActivateBackup\Windows" >nul 2>&1
if exist "%windir%\ServiceProfiles\NetworkService\AppData\Roaming\Microsoft\SoftwareProtectionPlatform\tokens.dat" xcopy "%windir%\ServiceProfiles\NetworkService\AppData\Roaming\Microsoft\SoftwareProtectionPlatform\*.*" "%~dp0\ActivateBackup\Windows" /s>nul
if exist "%windir%\System32\spp\store\2.0\tokens.dat" xcopy "%windir%\System32\spp\store\2.0\*.*" "%~dp0\ActivateBackup\Windows" /s>nul
if exist "C:\ProgramData\Microsoft\OfficeSoftwareProtectionPlatform\tokens.dat" (
md "%~dp0\ActivateBackup\Office" >nul
xcopy "C:\ProgramData\Microsoft\OfficeSoftwareProtectionPlatform" "%~dp0\ActivateBackup\Office"/s >nul
pause >nul
cscript OSPP.VBS /dstatus |findstr "LICENSED" >nul
if %errorlevel%==0  (
@echo   ===Office Da duoc kich hoat ban quyen VINH VIEN ===
goto exit
) else (
@echo   === Loi khong mong muon hoac Step 3 - CID khong chinh xac ===
@echo       === Kich hoat khong thanh cong. Vui Long thu lai! ===
@echo.
pause >nul
goto begin
)
:exit
@echo.
@echo =========================================================
@echo [  Cam on ban da su dung Tool Activate Windows & Office  ]
@echo [     Thanks for using Tool Activate Windows & Office    ]
@echo =========================================================

timeout 3
cscript ospp.vbs /dstatusall >status.txt
start status.txt
exit

:act2019
@echo.
for %%a in (4,5,6) do (if exist "%ProgramFiles%\Microsoft Office\Office1%%a\ospp.vbs" (cd /d "%ProgramFiles%\Microsoft Office\Office1%%a")
If exist "%ProgramFiles% (x86)\Microsoft Office\Office1%%a\ospp.vbs" (cd /d "%ProgramFiles% (x86)\Microsoft Office\Office1%%a"))
set /p key= 1. Nhap Key : 
@echo.
@echo......................................................................
cd /d %ProgramFiles%\Microsoft Office\Office16
cd /d %ProgramFiles(x86)%\Microsoft Office\Office16
cscript ospp.vbs /inslic:"..\root\Licenses16\ProPlus2019VL_MAK_AE-ul-phn.xrm-ms"
cscript ospp.vbs /inslic:"..\root\Licenses16\ProPlus2019VL_MAK_AE-ul-oob.xrm-ms"
cscript ospp.vbs /inslic:"..\root\Licenses16\ProPlus2019VL_MAK_AE-ppd.xrm-ms"
cscript ospp.vbs /inslic:"..\root\Licenses16\ProPlus2019VL_MAK_AE-pl.xrm-ms"
@echo 2. Dang cai dat Key
cscript OSPP.VBS /inpkey:%key%
@echo.
@echo 3. Dang Xuat Installion ID va Get CID kich hoat Office
for /f "tokens=8" %%i in ('cscript ospp.vbs /dinstid') do set MyIID=%%i
for /f "tokens=*" %%b in ('powershell -Command "$req = [System.Net.WebRequest]::Create('https://huyphung.com/public-api/get-cid?username=an8amc&password=@Ann123456&iid=%MyIID%');$resp = New-Object System.IO.StreamReader $req.GetResponse().GetResponseStream(); $resp.ReadToEnd()"') do set ACID=%%b
set CID=%ACID:~1,48%
echo CID cua ban %CID% 
@echo 5. Dang kich hoat ban quyen
Cscript ospp.vbs /actcid:%CID% & cscript ospp.vbs /act 
cscript OSPP.VBS /dstatus |findstr "LICENSED" >nul
if %errorlevel%==0  (
@echo   === Da kich hoat ban quyen VINH VIEN ===
Set /p "supporter=  - Nhap ten Supporter(neu ban khong phai la Supporter thi an enter) : "
echo Dang xuat status Office
echo Supporter:%supporter% >k1.txt & echo Key:%key% >>k1.txt & echo IID:%MyIID% >>k1.txt & echo CID:%CID% >>k1.txt & cscript ospp.vbs /dstatus >>k1.txt & start k1.txt & start winword & exit
@echo Dang Backup Ban Quyen Office
rd /s /q "%~dp0\ActivateBackup" >nul 2>&1
md "%~dp0\ActivateBackup\Windows" >nul 2>&1
if exist "%windir%\ServiceProfiles\NetworkService\AppData\Roaming\Microsoft\SoftwareProtectionPlatform\tokens.dat" xcopy "%windir%\ServiceProfiles\NetworkService\AppData\Roaming\Microsoft\SoftwareProtectionPlatform\*.*" "%~dp0\ActivateBackup\Windows" /s>nul
if exist "%windir%\System32\spp\store\2.0\tokens.dat" xcopy "%windir%\System32\spp\store\2.0\*.*" "%~dp0\ActivateBackup\Windows" /s>nul
if exist "C:\ProgramData\Microsoft\OfficeSoftwareProtectionPlatform\tokens.dat" (
md "%~dp0\ActivateBackup\Office" >nul
xcopy "C:\ProgramData\Microsoft\OfficeSoftwareProtectionPlatform" "%~dp0\ActivateBackup\Office"/s >nul
pause >nul
cscript OSPP.VBS /dstatus |findstr "LICENSED" >nul
if %errorlevel%==0  (
@echo   ===Office Da duoc kich hoat ban quyen VINH VIEN ===
goto exit
) else (
@echo   === Loi khong mong muon hoac Step 3 - CID khong chinh xac ===
@echo       === Kich hoat khong thanh cong. Vui Long thu lai! ===
@echo.
pause >nul
goto begin
)

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
goto begin
)
echo Khong tim thay bat ky 1 ban sao luu, vui long kiem tra lai
echo Nhan phim bat ky de tiep tuc...
pause>nul
cls & goto:begin
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
goto:begin