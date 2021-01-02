@echo off
set tempcd=%CD%
:mainmenu
cls
echo                       --- Office 2019 ProPlus Hebrew Installer ---
echo.
echo                                   by: RandomSun
echo.
echo   Choose your action?
echo.
echo    1. Install Office 2019 Online
echo    2. Install Office 2019 Offline + Later use
echo    3. Office Activation
echo    4. Check Office Status
echo    5. Uninstallation Methods
echo    q. Exit
echo.
echo.
echo.
set /p Input=Enter Your Number:
If /I "%Input%"=="1" goto installonline
If /I "%Input%"=="2" goto installoffline
If /I "%Input%"=="3" goto activate
If /I "%Input%"=="4" goto check
If /I "%Input%"=="5" goto uninstall
If /I "%Input%"=="q" goto q
else
goto mainmenu
:installonline
cls
echo                       --- Office 2019 ProPlus Hebrew Installer ---
echo.
echo                                   by: RandomSun
echo.
echo.
echo   which components do you like to install in Office 2019 ProPlus?
echo.
echo    1. Word, Excel, PowerPoint Only!
echo    2. All Components
echo    0. Back to menu
echo.
echo.
echo.
set /p Input=Enter Your Number:
If /I "%Input%"=="1" goto 1
If /I "%Input%"=="2" goto 2
If /I "%Input%"=="0" goto mainmenu
else
goto installonline
:1
cls
echo                       --- Office 2019 ProPlus Hebrew Installer ---
echo.
echo                                   by: RandomSun
echo.
echo.
echo   which Architecture do you want to install?
echo.
echo    1. 64x Architecture
echo    2. 86x Architecture
echo    0. Back to menu
echo.
echo.
echo.
set /p Input=Enter Your Number:
If /I "%Input%"=="1" goto 64xhalf
If /I "%Input%"=="2" goto 32xhalf
If /I "%Input%"=="0" goto installonline
else
goto 1
:64xhalf
cls
echo                       --- Office 2019 ProPlus Hebrew Installer ---
echo.
echo                                   by: RandomSun
echo.
echo.
echo    Press Any key To install your desierd office
pause
setup.exe /configure O2019(W,E,P)64xHEB.xml
pause
goto mainmenu
:32xhalf
cls
echo                       --- Office 2019 ProPlus Hebrew Installer ---
echo.
echo                                   by: RandomSun
echo.
echo.
echo    Press Any key To install your desierd office
pause
setup.exe /configure O2019(W,E,P)32xHEB.xml
pause
goto mainmenu
:2
cls
echo                       --- Office 2019 ProPlus Hebrew Installer ---
echo.
echo                                   by: RandomSun
echo.
echo.
echo   which Architecture do you want to install?
echo.
echo    1. 64x Architecture
echo    2. 86x Architecture
echo    0. Back to menu
echo.
echo.
echo.
set /p Input=Enter Your Number:
If /I "%Input%"=="1" goto 64xfull
If /I "%Input%"=="2" goto 32xfull
If /I "%Input%"=="0" goto installonline
else
goto 2
:64xfull
cls
echo                       --- Office 2019 ProPlus Hebrew Installer ---
echo.
echo                                   by: RandomSun
echo.
echo.
echo    Press Any key To install your desierd office
pause
setup.exe /configure O2019(ALL)64xHEB.xml
pause
goto mainmenu
:32xfull
cls
echo                       --- Office 2019 ProPlus Hebrew Installer ---
echo.
echo                                   by: RandomSun
echo.
echo.
echo    Press Any key To install your desierd office
pause
setup.exe /configure O2019(ALL)32xHEB.xml
pause
goto mainmenu
:installoffline
cls
echo                       --- Office 2019 ProPlus Hebrew Installer ---
echo.
echo                                   by: RandomSun
echo.
echo.
echo   which action do you want to do?
echo.
echo    1. Download And Install Office 2019 First
echo    2. Use Existed Offline Install
echo    0. Back to menu
echo.
echo.
echo.
set /p Input=Enter Your Number:
If /I "%Input%"=="1" goto firstinstalloffline
If /I "%Input%"=="2" goto existinstall
If /I "%Input%"=="0" goto mainmenu
else
goto installoffline
:firstinstalloffline
cls
echo                       --- Office 2019 ProPlus Hebrew Installer ---
echo.
echo                                   by: RandomSun
echo.
echo     Press any key to start the Download of the Offline Office 2019 Install.
echo     When the window is Done press Enter and exit the window.
pause
start %tempcd%\OfflineInstall\YAOCTRU_Generator.cmd
pause
echo When the Download Procces Finish, Press any key and exit the window.
pause
start %tempcd%\OfflineInstall\16.0.13426.20404_x86x64_he-IL_Monthly_wget.bat
pause
echo     Press Any Key Twice To Procced with the installation.
pause
start %tempcd%\OfflineInstall\YAOCTRIR_Configurator.cmd
pause
goto mainmenu
:existinstall
cls
echo                       --- Office 2019 ProPlus Hebrew Installer ---
echo.
echo                                   by: RandomSun
echo.
echo.
echo     Press any key twice to continute with the installation.
pause
start %tempcd%\OfflineInstall\YAOCTRIR_Configurator.cmd
goto mainmenu
:activate
cls
echo                       --- Office 2019 ProPlus Hebrew Installer ---
echo.
echo                                   by: RandomSun
echo.
echo     Press any key to start the Activation.
echo     Wait for the black window to show text for an successful Activation and exit the window.
pause    
start %tempcd%\kms_vl_all_aio.cmd
pause
goto mainmenu
:uninstall
cls
echo                       --- Office 2019 ProPlus Hebrew Installer ---
echo.
echo                                   by: RandomSun
echo.
echo.
echo   Chose Your action:
echo.
echo    1. Full Office Uninstall
echo    2. Remove Licenses
echo    3. Uninstall Keys
echo    0. Back to menu
echo.
echo.
echo.
set /p Input=Enter Your Number:
If /I "%Input%"=="1" goto uninstallfullscrub
If /I "%Input%"=="2" goto removelic
If /I "%Input%"=="3" goto uninstallkey
If /I "%Input%"=="0" goto mainmenu
else
goto :uninstall

:uninstallfullscrub
cls
echo                       --- Office 2019 ProPlus Hebrew Installer ---
echo.
echo                                   by: RandomSun
echo.
echo.
echo   You are about to uninstall your current office.
start %tempcd%\OfficeScrubber\Full_Scrub.cmd
pause
goto uninstall
:removelic
cls
echo                       --- Office 2019 ProPlus Hebrew Installer ---
echo.
echo                                   by: RandomSun
echo.
echo.
echo   You are about to Remove Your Office Licenses.
start %tempcd%\OfficeScrubber\Remove_Licenses.cmd
pause
goto uninstall
:uninstallkey
cls
echo                       --- Office 2019 ProPlus Hebrew Installer ---
echo.
echo                                   by: RandomSun
echo.
echo.
echo   You are about to Uninstall Your Office Keys.
start %tempcd%\OfficeScrubber\Uninstall_Keys.cmd
pause
goto uninstall
:check
IF EXIST c:\Program Files (x86)\Microsoft Office\Office16\OSPP.VBS goto 86x
IF EXIST c:\Program Files\Microsoft Office\Office16\OSPP.VBS goto 64x
else
echo Error
:64x
cls
echo                       --- Office 2019 ProPlus Hebrew Installer ---
echo.
echo                                   by: RandomSun
echo.
echo.
CD c:\Program Files\Microsoft Office\Office16\
cscript OSPP.VBS /dstatus
pause 
pushd
goto mainmenu
:86x
cls
echo                       --- Office 2019 ProPlus Hebrew Installer ---
echo.
echo                                   by: RandomSun
echo.
echo.
CD c:\Program Files (x86)\Microsoft Office\Office16\
cscript OSPP.VBS /dstatus
pause
pushd
goto mainmenu
:q
exit

