:check
cls
echo off
cls
color f3
@echo. Cong cu check for update:
@echo off
@echo Ban muon cap nhat nhung ban fix loi hay cap nhat ban moi nhat
@echo Nhan A de cap nhat ban moi nhat,B de cap nhat ban fix loi, C de cap nhat ca 2
Choice /N /C ABCDEFGHKLMNOPQRSTXY /M "*Please enter your choice:
if ERRORLEVEL 1 goto fixloi	
if ERRORLEVEL 2 goto update
if ERRORLEVER 3 goto all
:all
cd.
echo Dowloading......
powershell -command "& { (New-Object Net.WebClient).DownloadFile('https://codeload.github.com/nguyenphamannsg/Tool-Activate-Windows-Office/zip/master', ' Tool Activate Windows-Office.zip') }
powershell -command "& { (New-Object Net.WebClient).DownloadFile('https://github.com/nguyenphamannsg/FIX-LOI-Tool-Activate-Windows-Office/archive/master.zip', 'Tool Activate Windows-Office.zip') }
@echo. Neu yeu cau mat khau thi nhap:nguyenphaman
@echo nhan phim bat ki de tiep tuc
pause>nul
goto check
:update
cd.
echo Dowloading......
powershell -command "& { (New-Object Net.WebClient).DownloadFile('https://codeload.github.com/nguyenphamannsg/Tool-Activate-Windows-Office/zip/master', ' Tool Activate Windows-Office.zip') }
@echo. Neu yeu cau mat khau thi nhap:nguyenphaman
@echo nhan phim bat ki de tiep tuc
pause>nul
goto check
:fixloi
cd.
echo Dowloading......
powershell -command "& { (New-Object Net.WebClient).DownloadFile('https://github.com/nguyenphamannsg/FIX-LOI-Tool-Activate-Windows-Office/archive/master.zip', 'Tool Activate Windows-Office.zip') }
@echo. Neu yeu cau mat khau thi nhap:nguyenphaman
@echo nhan phim bat ki de tiep tuc
pause>nul
goto check