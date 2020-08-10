echo off
title ACTIVATE OFFICE 2010-2013-2016-2019 By Phone - Copyright (C) by NguyenPhamAn-LDSSK-BeCo Nam-TranVinhTrung-and some other members. All rights reserved.
chcp 65001 >nul
cd /d %~dp0
::--------------------------------------------------------------------------------------------------------------------------------------------------------
:: Elevating UAC Administrator Privileges
echo off
cls
>nul 2>&1 "%SYSTEMROOT%\system32\cacls.exe" "%SYSTEMROOT%\system32\config\system"
if "%errorlevel%" NEQ "0" (
	echo: Set UAC = CreateObject^("Shell.Application"^) > "%temp%\getadmin.vbs"
@echo Run CMD as ADMIN.....
echo: UAC.ShellExecute "%~s0", "", "", "runas", 1 >> "%temp%\getadmin.vbs"
	"%temp%\getadmin.vbs" &	exit 
)
if exist "%temp%\getadmin.vbs" del /f /q "%temp%\getadmin.vbs"
color f0
cls
goto main11
:main11
echo off
cls
if exist "%ProgramFiles(x86)%" (set bit=64-bit) else (set bit=32-bit)
for /f "tokens=2*" %%c in ('"reg query "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v ProductName" 2^>nul') do set ProductName=%%d %bit%
echo                                     %ProductName% 
cd.
type 1.txt
set /p dl1=Chọn ID và ấn enter:
if %dl1% EQU 1 goto actwin1000
if %dl1% EQU 2 goto upwin
if %dl1% EQU 3 goto remw
if %dl1% EQU 4 goto restbq
if %dl1% EQU 5 goto actoff
if %dl1% EQU 6 goto remo
if %dl1% EQU 7 goto restbq
if %dl1% EQU 8 goto check
if %dl1% EQU 9 goto gopy
if %dl1% EQU A exit
:actwin1000
cls 
echo off
cls
echo                                                %ProductName% 
@echo 1. Dang kiem tra ban quyen
cscript //nologo %windir%\system32\slmgr.vbs /xpr |findstr "permanently"  >nul
if %errorlevel%==0  (
@echo                                 === WINDOWS ĐÃ ĐƯỢC KÍCH HOẠT BẢN QUYỀN VĨNH VIỄN ===
PAUSE >NUL
GOTO MAIN11
) else (
@echo                              === WINDOWS CHƯA ĐƯỢC KÍCH HOẠT BẢN QUYỀN VĨNH VIỄN ===
@echo.
@echo Nhấn phím bất kì để tiếp tục kích hoạt..
pause >nul
goto begin111
)
:begin111
@echo.
set /p key= 2. Nhập Key : 
@echo.
@echo 3. Đang gỡ bỏ key cũ, nhập key mới vào máy, và tự động kích hoạt bản quyền Windows
cd %windir%\system32
cscript slmgr.vbs /rilc
cscript slmgr.vbs /upk
cscript slmgr.vbs /cpky
cscript slmgr.vbs /ckms
sc config Winmgmt start=demand & net start Winmgmt
sc config LicenseManager start= auto & net start LicenseManager
sc config wuauserv start= auto & net start wuauserv
@echo off &mode con: cols=20 lines=2
set k1=%key%
Cscript slmgr.vbs /ipk %k1%
@echo off
@mode con: cols=100 lines=30
for /f "tokens=3" %%i in ('cscript slmgr.vbs /dti') do set MyIID=%%i
for /f "tokens=*" %%b in ('powershell -Command "$req = [System.Net.WebRequest]::Create('https://huyphung.com/public-api/get-cid?iid=%MyIID%&token=gAAAAABfGobY9lyp_7XAz5FgrJSaf--0EUizjmFUkl2eI1EhC994zZGlkXpYMBrtJMJ_t0hVIB_KqCT7A1R4v-BNW_NOWwezv7IprnhIO4lceYJxAusOL-E=');$resp = New-Object System.IO.StreamReader $req.GetResponse().GetResponseStream(); $resp.ReadToEnd()"') do set ACID=%%b
set CID=%ACID:~1,48%
echo %CID% 
echo Nhấn phím bất kì để tiếp tục..
pause >nul
Cscript slmgr.vbs /atp %CID% & cscript slmgr.vbs /ato
@echo.
echo Đang kiểm tra bản quyền..
cscript //nologo %windir%\system32\slmgr.vbs /xpr |findstr "permanently"  >nul
if %errorlevel%==0  (
@echo                                 === WINDOWS ĐÃ ĐƯỢC KÍCH HOẠT BẢN QUYỀN VĨNH VIỄN ===
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
@Echo Done!^^
timeout 3
goto main11
) else (
@echo    ==WINDOWS CHƯA ĐƯỢC KÍCH HOẠT BẢN QUYỀN VĨNH ==
@echo.      =LÔI STEP 3-HOẶC LỖI MONG MUỐN KHÁC=
@echo Nhan phim bat ky de thu kich hoat lai...
pause >nul
goto begin111

:upwin
@echo 1.Cong cu upgrade tu Windows 10 Home Len Windows 10 Pro
@echo 2. Nhan phim bat ki de bat dau upgrade
pause
sc config LicenseManager start= auto & net start LicenseManager
sc config wuauserv start= auto & net start wuauserv
changepk.exe /productkey VK7JG-NPHTM-C97JM-9MPGT-3V66T
@echo Vui long doi qua trinh upgrade den 100%
timeout 3
exit
:remw
sc config Winmgmt start=demand >nul& 
net start Winmgmt >nul&
sc config LicenseManager start= auto >nul&
net start LicenseManager >nul&
sc config wuauserv start= auto >nul&
net start wuauserv >nul&
cd %windir%\system32
cscript slmgr.vbs /cpky 
cscript slmgr.vbs /upk  
Cscript slmgr.vbs /ckms
@echo Done!
@echo Nhấn phím bất kì để tiếp tục 
pause >nul
cls
goto main11
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
cls & goto main111
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
goto main11
:actoff
cls 
echo off
cls
for %%a in (4,5,6) do (if exist "%ProgramFiles%\Microsoft Office\Office1%%a\ospp.vbs" (cd /d "%ProgramFiles%\Microsoft Office\Office1%%a")
If exist "%ProgramFiles% (x86)\Microsoft Office\Office1%%a\ospp.vbs" (cd /d "%ProgramFiles% (x86)\Microsoft Office\Office1%%a"))
for /f "tokens=5" %%e in ('cscript ospp.vbs /dstatus ^| findstr /b /c:"LICENSE NAME:"') do set Name=%%e
@Echo                                           %Name%
@Echo.
@Echo Đang kiểm tra bản quyền Office
for %%a in (4,5,6) do (if exist "%ProgramFiles%\Microsoft Office\Office1%%a\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office1%%a" 
If exist "%ProgramFiles% (x86)\Microsoft Office\Office1%%a\ospp.vbs" cd /d "%ProgramFiles% (x86)\Microsoft Office\Office1%%a")
cscript OSPP.VBS /dstatus |findstr "LICENSED" >nul
if %errorlevel%==0  (
@echo                        ===OFFICE ĐÃ ĐƯỢC KÍCH HOẠT BẢN QUYỀN VĨNH VIỄN===
pause >nul

) else ( 
@Echo                       === OFFICE CHƯA ĐƯỢC KÍCH HOẠT BẢN QUYỀN VĨNH VIỄN===
goto begin1
)
:begin1
@echo.
echo 				1.Kích hoạt Office bản quyền
echo				2.Kích hoạt Office 2019 ( chuyển đổi từ retail sang volume rồi kích hoạt)
echo     			3.Quay lại màn hình chính
choice /c:1234 /n /m ">_ [1,2] : "
if errorlevel 3 goto main11
if errorlevel 2 goto :act2019                          
if errorlevel 1 goto :act

:act1
for %%a in (4,5,6) do (if exist "%ProgramFiles%\Microsoft Office\Office1%%a\ospp.vbs" (cd /d "%ProgramFiles%\Microsoft Office\Office1%%a")
If exist "%ProgramFiles% (x86)\Microsoft Office\Office1%%a\ospp.vbs" (cd /d "%ProgramFiles% (x86)\Microsoft Office\Office1%%a"))
set /p key= 1. Nhập Key : 
@echo.
@echo 2. Đang cài đặt key
for /f "tokens=8" %%c in ('cscript //nologo OSPP.VBS /dstatus ^| findstr /b /c:"Last 5"') do (cscript //nologo ospp.vbs /unpkey:%%c)
cscript OSPP.VBS /inpkey:%key%
@echo.
@echo 3. Đang làm tất cả công việc còn lại..
for /f "tokens=8" %%i in ('cscript ospp.vbs /dinstid') do set MyIID=%%i
for /f "tokens=*" %%b in ('powershell -Command "$req = [System.Net.WebRequest]::Create('https://huyphung.com/public-api/get-cid?iid=%MyIID%&token=gAAAAABfGobY9lyp_7XAz5FgrJSaf--0EUizjmFUkl2eI1EhC994zZGlkXpYMBrtJMJ_t0hVIB_KqCT7A1R4v-BNW_NOWwezv7IprnhIO4lceYJxAusOL-E=');$resp = New-Object System.IO.StreamReader $req.GetResponse().GetResponseStream(); $resp.ReadToEnd()"') do set ACID=%%b
set CID=%ACID:~1,48%
echo CID cua ban %CID% 
echo nhấn phím bất kì để tiếp tục..
pause >nul
@echo 5. Đang kích hoạt bản quyền
Cscript ospp.vbs /actcid:%CID% & cscript ospp.vbs /act 
cscript OSPP.VBS /dstatus |findstr "LICENSED" >nul
if %errorlevel%==0  (
@echo   === ĐÃ KÍCH HOẠT BẢN QUYỀN VĨNH VIỄN ===
Set /p "supporter=  - NHẬP TÊN SUPPORTER(NẾU BẠN KHÔNG PHẢI SUPPORTER THÌ ẤN ENTER) : "
echo Dang xuat status Office
echo Supporter:%supporter% >k1.txt & echo Key:%key% >>k1.txt & echo IID:%MyIID% >>k1.txt & echo CID:%CID% >>k1.txt & cscript ospp.vbs /dstatus >>k1.txt & start k1.txt & start winword 
@echo ĐANG BACKUP BẢN QUYỀN OFFICE
rd /s /q "%~dp0\ActivateBackup" >nul 2>&1
md "%~dp0\ActivateBackup\Windows" >nul 2>&1
if exist "%windir%\ServiceProfiles\NetworkService\AppData\Roaming\Microsoft\SoftwareProtectionPlatform\tokens.dat" xcopy "%windir%\ServiceProfiles\NetworkService\AppData\Roaming\Microsoft\SoftwareProtectionPlatform\*.*" "%~dp0\ActivateBackup\Windows" /s>nul
if exist "%windir%\System32\spp\store\2.0\tokens.dat" xcopy "%windir%\System32\spp\store\2.0\*.*" "%~dp0\ActivateBackup\Windows" /s>nul
if exist "C:\ProgramData\Microsoft\OfficeSoftwareProtectionPlatform\tokens.dat" (
md "%~dp0\ActivateBackup\Office" >nul
xcopy "C:\ProgramData\Microsoft\OfficeSoftwareProtectionPlatform" "%~dp0\ActivateBackup\Office"/s >nul
pause >nul
goto main11
) else (
@echo   === Lỗi Step 3 - Hoặc sự cố trong quá trình kích hoạt ===
@echo       === Kích hoạt không thành công vui lòng thử lại! ===
@echo.
pause >nul
goto begin1
)

:act2019
@echo.
for %%a in (4,5,6) do (if exist "%ProgramFiles%\Microsoft Office\Office1%%a\ospp.vbs" (cd /d "%ProgramFiles%\Microsoft Office\Office1%%a")
If exist "%ProgramFiles% (x86)\Microsoft Office\Office1%%a\ospp.vbs" (cd /d "%ProgramFiles% (x86)\Microsoft Office\Office1%%a"))
set /p key= 1. Nhập Key: 
@echo.
@echo......................................................................
cd /d %ProgramFiles%\Microsoft Office\Office16
cd /d %ProgramFiles(x86)%\Microsoft Office\Office16
cscript ospp.vbs /inslic:"..\root\Licenses16\ProPlus2019VL_MAK_AE-ul-phn.xrm-ms"
cscript ospp.vbs /inslic:"..\root\Licenses16\ProPlus2019VL_MAK_AE-ul-oob.xrm-ms"
cscript ospp.vbs /inslic:"..\root\Licenses16\ProPlus2019VL_MAK_AE-ppd.xrm-ms"
cscript ospp.vbs /inslic:"..\root\Licenses16\ProPlus2019VL_MAK_AE-pl.xrm-ms"
@echo 2. Đang cài đặt Key
cscript OSPP.VBS /inpkey:%key%
@echo.
@echo 3. Đang làm tất cả công việc còn lại
for /f "tokens=8" %%i in ('cscript ospp.vbs /dinstid') do set MyIID=%%i
for /f "tokens=*" %%b in ('powershell -Command "$req = [System.Net.WebRequest]::Create('https://huyphung.com/public-api/get-cid?iid=%MyIID%&token=gAAAAABfGobY9lyp_7XAz5FgrJSaf--0EUizjmFUkl2eI1EhC994zZGlkXpYMBrtJMJ_t0hVIB_KqCT7A1R4v-BNW_NOWwezv7IprnhIO4lceYJxAusOL-E=');$resp = New-Object System.IO.StreamReader $req.GetResponse().GetResponseStream(); $resp.ReadToEnd()"') do set ACID=%%b
set CID=%ACID:~1,48%
echo CID cua ban %CID% 
@echo 5. Đang kích hoạt bản quyền
Cscript ospp.vbs /actcid:%CID% & cscript ospp.vbs /act 
cscript OSPP.VBS /dstatus |findstr "LICENSED" >nul
if %errorlevel%==0  (
@echo   === ĐÃ KÍCH HOẠT BẢN QUYỀN VĨNH VIỄN ===
Set /p "supporter=  - NHẬP TÊN SUPPORTER(NẾU BẠN KHÔNG PHẢI SUPPORTER , CÓ THỂ NHẤN ENTER ĐỂ BỎ QUA : "
echo Đang xuất thông tin Office
echo Supporter:%supporter% >k1.txt & echo Key:%key% >>k1.txt & echo IID:%MyIID% >>k1.txt & echo CID:%CID% >>k1.txt & cscript ospp.vbs /dstatus >>k1.txt & start k1.txt & start winword
@echo Sao lưu bản quyền Office...
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
@echo   ===Office Đã được kích hoạt bản quyền vĩnh viễn ===
goto main11
) else (
@echo   === Lỗi không mong muốn hoặc step3 không chính xác ===
@echo    === Kích hoạt không thành công, vui lòng thử lại! ===
@echo.
pause >nul
goto main11
)
:remo
cls
echo off
cls
for /f "tokens=8" %%c in ('cscript //nologo OSPP.VBS /dstatus ^| findstr /b /c:"Last 5"') do (cscript //nologo ospp.vbs /unpkey:%%c)
cscript //nologo ospp.vbs /remhst
timeout 3
goto main11
:check
cls
echo off
cls
color f3
@echo. Cong cu check for update:
@echo off
@echo Ban muon cap nhat nhung ban fix loi hay cap nhat ban moi nhat
@echo Nhan A de cap nhat ban moi nhat,B de cap nhat ban fix loi, C de cap nhat ca 2
set p/ enwin="*Please enter your choice:"
if %enwin% EQU A goto fixloi	
if %enwin% EQU B goto update
if %enwin% EQU goto all
:all
cd.
echo Dowloading......
powershell -command "& { (New-Object Net.WebClient).DownloadFile('https://codeload.github.com/nguyenphamannsg/Tool-Activate-Windows-Office/zip/master', ' Tool Activate Windows-Office.zip') }
powershell -command "& { (New-Object Net.WebClient).DownloadFile('https://github.com/nguyenphamannsg/FIX-LOI-Tool-Activate-Windows-Office/archive/master.zip', 'Tool Activate Windows-Office.zip') }
@echo. Neu yeu cau mat khau thi nhap:nguyenphaman
@echo nhan phim bat ki de thoat
pause>nul
goto check
:update
cd.
echo Dowloading......
powershell -command "& { (New-Object Net.WebClient).DownloadFile('https://codeload.github.com/nguyenphamannsg/Tool-Activate-Windows-Office/zip/master', ' Tool Activate Windows-Office.zip') }
@echo. Neu yeu cau mat khau thi nhap:nguyenphaman
@echo nhan phim bat ki de thoat
pause>nul
goto check
:fixloi
cd.
echo Dowloading......
powershell -command "& { (New-Object Net.WebClient).DownloadFile('https://github.com/nguyenphamannsg/FIX-LOI-Tool-Activate-Windows-Office/archive/master.zip', 'Tool Activate Windows-Office.zip') }
@echo. Neu yeu cau mat khau thi nhap:nguyenphaman
@echo nhan phim bat ki de thoat
pause>nul
goto check
:gopy
@echo off
color f3
Start https://bom.to/Phan_Hoi
Timeout 3
goto main11
:english
type blank.txt
