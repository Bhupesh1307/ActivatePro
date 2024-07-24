@echo off
title Activating Windows
color a

REM Setting Variables
set "versionBeb=Version: 3.1.4"
set "versionExp=ActivatePro version: 3.1.4"
set "expertMode=0"
set "mainSel=Select an option to continue: "
set "winSel=Select an option to continue: "
set "offSel=Select an option to continue: "

REM Baby Mode
:mainMenu
REM ActivatePro Logo
echo                _   _            _       _____           
echo      /\       ^| ^| (_)          ^| ^|     ^|  __ \          
echo     /  \   ___^| ^|_ ___   ____ _^| ^|_ ___^| ^|__) ^| __ ___  
echo    / /\ \ / __^| __^| \ \ / / _` ^| __/ _ \  ___/ '__/ _ \ 
echo   / ____ \ (__^| ^|_^| ^|\ V / (_^| ^| ^|^|  __/ ^|   ^| ^| ^| (_) ^|
echo  /_/    \_\___^|\__^|_^| \_/ \__,_^|\__\___^|_^|   ^|_^|  \___/ 
echo:
echo:

REM Tool Details
echo ActivatePro
echo %versionBeb%
echo by Bhupesh Chaubey
echo Github Profile: https://github.com/Bhupesh1307
echo:
echo:
echo Activate Windows 10/11 and Microsoft Office 2010-2021



REM Menu Options
echo Choose what to activate:
echo:
echo [1]	Microsoft Windows 10/11
echo [2]	Microsoft Office
echo [3]	Expert Mode
echo [99]	Exit

echo:
echo:

set /p "selA=%mainSel%"			REM Taking Input from User
set /a selA=%selA%			REM Converting the taken Input into Integer value

REM Menu Options Working Function
if %selA%==1 (
	cls
	goto windows
) else if %selA%==2 (
	cls
	goto office
) else if %selA%==3 (
	cls
	set "expertMode=1"
	goto expert
) else if %selA%==99 (
	exit
) else (
	goto errorMain
)




REM Windows
:windows

REM Windows logo
echo  __          ___           _                   
echo  \ \        / (_)         ^| ^|                  	
echo   \ \  /\  / / _ _ __   __^| ^| _____      _____ 
echo    \ \/  \/ / ^| ^| '_ \ / _` ^|/ _ \ \ /\ / / __^|
echo     \  /\  /  ^| ^| ^| ^| ^| (_^| ^| (_) \ V  V /\__ \
echo      \/  \/   ^|_^|_^| ^|_^|\__,_^|\___/ \_/\_/ ^|___/
echo:
echo:
echo Select your Microsoft Windows Edition:
echo:

REM Windows Options
echo [0]	Go Back
echo [1]	Home
echo [2]	Home N
echo [3]	Pro
echo [4]	Pro N
echo [5]	Enterprise
echo [6]	Enterprise N
echo [7]	Education
echo [8]	Education N
echo [99]	Exit

echo:
echo:

set /p "selB=%winSel%"				REM Taking input from user
set /a selB=%selB%				REM Converting taken input into integer value

REM Windows Options Working Function
if %selB%==0 (
	cls
	set "mainSel=Select an option to continue: "
	set "winSel=Select an option to continue: "
	set "offSel=Select an option to continue: "
	goto mainMenu
) else if %selB%==1 (
	cls
	goto home
) else if %selB%==2 (
	cls
	goto home-n
) else if %selB%==3 (
	cls
	goto pro
) else if %selB%==4 (
	cls
	goto pro-n
) else if %selB%==5 (
	cls
	goto enterprise
) else if %selB%==6 (
	cls
	goto enterprise-n
) else if %selB%==7 (
	cls
	goto education
) else if %selB%==8 (
	cls
	goto education-n
) else if %selB%==99 (
	exit
) else (
        goto errorWindows
)





REM Activation Commands for Windows Editions
:home
echo Activating Windows Home Edition...
timeout /t 3 > nul

slmgr -upk
slmgr.vbs /ipk TX9XD-98N7V-6WMQ6-BX7FG-H8Q99
slmgr /skms kms8.msguides.com
slmgr /ato

goto exit



:home-n
echo Activating Windows Home N Edition...
timeout /t 3 > nul

slmgr -upk
slmgr.vbs /ipk 3KHY7-WNT83-DGQKR-F7HPR-844BM
slmgr /skms kms8.msguides.com
slmgr /ato

goto exit



:pro
echo Activating Windows Pro Edition...
timeout /t 3 > nul

slmgr -upk
slmgr.vbs /ipk W269N-WFGWX-YVC9B-4J6C9-T83GX
slmgr /skms kms8.msguides.com
slmgr /ato

goto exit



:pro-n
echo Activating Windows Pro N Edition...
timeout /t 3 > nul

slmgr -upk
slmgr.vbs /ipk MH37W-N47XK-V7XM9-C7227-GCQG9
slmgr /skms kms8.msguides.com
slmgr /ato

goto exit



:enterprise
echo Activating Windows Enterprise Edition...
timeout /t 3 > nul

slmgr -upk
slmgr.vbs /ipk NPPR9-FWDCX-D2C8J-H872K-2YT43
slmgr /skms kms8.msguides.com
slmgr /ato

goto exit



:enterprise-n
echo Activating Windows Enterprise N Edition...
timeout /t 3 > nul

slmgr -upk
slmgr.vbs /ipk DPH2V-TTNVB-4X9Q3-TJR4H-KHJW4
slmgr /skms kms8.msguides.com
slmgr /ato

goto exit



:education
echo Activating Windows Education Edition...
timeout /t 3 > nul

slmgr -upk
slmgr.vbs /ipk NW6C2-QMPVW-D7KKK-3GKT6-VCFB2
slmgr /skms kms8.msguides.com
slmgr /ato

goto exit



:education-n
echo Activating Windows Education N Edition...
timeout /t 3 > nul

slmgr -upk
slmgr.vbs /ipk 2WH4N-8QGBV-H22JP-CT43Q-MDWWJ
slmgr /skms kms8.msguides.com
slmgr /ato

goto exit





REM Microsoft Office
:office

REM Office Logo
echo   __  __  _____    ____   __  __ _          
echo  ^|  \/  ^|/ ____^|  / __ \ / _^|/ _(_)         
echo  ^| \  / ^| (___   ^| ^|  ^| ^| ^|_^| ^|_ _  ___ ___ 
echo  ^| ^|\/^| ^|\___ \  ^| ^|  ^| ^|  _^|  _^| ^|/ __/ _ \
echo  ^| ^|  ^| ^|____) ^| ^| ^|__^| ^| ^| ^| ^| ^| ^| (_^|  __/
echo  ^|_^|  ^|_^|_____/   \____/^|_^| ^|_^| ^|_^|\___\___^|
echo:
echo:

REM Microsoft Office Options
echo Close any Microsoft Office App if running.
echo Select your Microsoft Office:
echo:
echo [0]	Go Back
echo [1]	Microsoft Office 2010
echo [2]	Microsoft Office 2013
echo [3]	Microsoft Office 2016
echo [4]	Microsoft Office 2019
echo [5]	Microsoft Office 2021
echo [99]	Exit

echo:
echo:
set /p "selC=%offSel%"				REM Taking Input from User
set /a selC=%selC%				REM Converting taken Input into Integer Value

REM Office Options Working Function
if %selC%==0 (
	cls
	set "mainSel=Select an option to continue: "
	set "winSel=Select an option to continue: "
	set "offSel=Select an option to continue: "
	goto mainMenu
) else if %selC%==1 (
	goto office14
) else if %selC%==2 (
	goto office15
) else if %selC%==3 (
	goto office16
) else if %selC%==4 (
	goto office16
) else if %selC%==5 (
	goto office16
) else if %selC%==99 (
	exit
) else (
	goto errorOffice
)



REM Activation Commands for Office Versions
:office14
cls
if exist "C:\Program Files (x86)\Microsoft Office\Office14\ospp.vbs" (
	cd "C:\Program Files (x86)\Microsoft Office\Office14"
) else if exist "C:\Program Files\Microsoft Office\Office14" (
	cd "c:\Program Files\Microsoft Office\Office14"
) else (
	cls
	echo Error: Microsoft Office not found!
	echo Couldn't find the path where Microsoft Office 2010 is installed.
	echo Make sure that Microsoft Office 2010 is installed at the default location.
	echo And then try again.
	echo:
	echo:
	goto office
)
cscript ospp.vbs /osppsvcauto
cscript ospp.vbs /sethst:vista-kms1.ad.gatech.edu
cscript ospp.vbs /act
cscript ospp.vbs /dstatus

goto exit



:office15
cls
if exist "C:\Program Files (x86)\Microsoft Office\Office15\ospp.vbs" (
	cd "C:\Program Files (x86)\Microsoft Office\Office15"
) else if exist "C:\Program Files\Microsoft Office\Office15" (
	cd "c:\Program Files\Microsoft Office\Office15"
) else (
	cls
	echo Error: Microsoft Office not found!
	echo Couldn't find the path where Microsoft Office 2013 is installed.
	echo Make sure that Microsoft Office 2013 is installed at the default location.
	echo And then try again.
	echo:
	echo:
	goto office
)
cscript ospp.vbs /osppsvcauto
cscript ospp.vbs /sethst:vista-kms1.ad.gatech.edu
cscript ospp.vbs /act
cscript ospp.vbs /dstatus

goto exit



:office16
cls
if exist "C:\Program Files (x86)\Microsoft Office\Office16\ospp.vbs" (
	cd "C:\Program Files (x86)\Microsoft Office\Office16"
) else if exist "C:\Program Files\Microsoft Office\Office16" (
	cd "c:\Program Files\Microsoft Office\Office16"
) else (
	cls
	echo Error: Microsoft Office not found!
	echo Couldn't find the path where Microsoft Office 2016 is installed.
	echo Make sure that the Microsoft Office 2016 is installed at the default location.
	echo And then try again.
	echo:
	echo:
	goto office
)
cscript ospp.vbs /osppsvcauto
cscript ospp.vbs /sethst:vista-kms1.ad.gatech.edu
cscript ospp.vbs /act
cscript ospp.vbs /dstatus

goto exit


REM Expert Mode
:expert
set /p "cmd=activatePro@Expert:~# "	REM Taking Input as a Command from User

REM Expert Commands Working Function
if "%cmd%"=="activatePro mode baby" (
	cls
	set "mainSel=Select an option to continue: "
	set "winSel=Select an option to continue: "
	set "offSel=Select an option to continue: "
	goto mainMenu
) else if "%cmd%"=="ap mode baby" (
	cls
	set "mainSel=Select an option to continue: "
	set "winSel=Select an option to continue: "
	set "offSel=Select an option to continue: "
	goto mainMenu
) else if "%cmd%"=="activatePro version" (
	echo %versionExp%
	echo:
	goto expert
) else if "%cmd%"=="activatePro -v" (
	echo %versionExp%
	echo:
	goto expert
) else if "%cmd%"== "ap version" (
	echo %versionExp%
	echo:
	goto expert
) else if "%cmd%"=="ap -v" (
	echo %versionExp%
	echo:
	goto expert
) else if "%cmd%"=="activatePro -h" (
	goto expertHelp
) else if "%cmd%"=="ap -h" (
	goto expertHelp
) else if "%cmd%"=="activatePro help" (
	goto expertHelp
) else if "%cmd%"=="ap help" (
	goto expertHelp
) else if "%cmd%"=="help" (
	goto expertHelp
) else if "%cmd%"=="activate windows home" (
	goto home
) else if "%cmd%"=="activate windows homeN" (
	goto home-n
) else if "%cmd%"=="activate windows pro" (
	goto pro
) else if "%cmd%"=="activate windows proN" (
	goto pro-n
) else if "%cmd%"=="activate windows enterprise" (
	goto enterprise
) else if "%cmd%"=="activate windows enterpriseN" (
	goto enterprise-n
) else if "%cmd%"=="activate windows education" (
	goto education
) else if "%cmd%"=="activate windows educationN" (
	goto education-n
) else if "%cmd%"=="activate office14" (
	goto office14
) else if "%cmd%"=="activate office15" (
	goto office15
) else if "%cmd%"=="activate office16" (
	goto office16
) else if "%cmd%"=="clear" (
	cls
	goto expert
) else if "%cmd%"=="exit" (
	exit
) else (
	echo %cmd% is not a valid command. try activatePro -h for help.
	echo:
	goto expert
)

REM Expert Commands Help
:expertHelp
echo ActivatePro is a Free and OpenSource tool, used to activate:
echo -	Microsoft Windows 10/11
echo -	Microsoft Office 2010-2021
echo:
echo ActivatePro uses KMS Services to activate your product.
echo:
echo ActivatePro Expert commands:
echo -	activatePro mode baby:		Provides user-interface for noobs
echo -	ap mode baby:			Provides user-interface for noobs
echo -	activatePro help:		Provides help for ActivatePro commands
echo -	ap help:			Provides help for ActivatePro commands
echo -	activatePro -h:			Provides help for ActivatePro commands
echo -	ap -h:				Provides help for ActivatePro commands
echo -	activatePro version:		Shows the current version of ActivatePro
echo - 	activatePro -v:			Shows the current version of ActivatePro
echo -	ap version: 			Shows the current version of ActivatePro
echo - 	ap -v:				Shows the current version of ActivatePro
echo -	activate office ^<version^>:	Activate Microsoft Office version
echo -	activate windows ^<edition^>:	Activate Microsoft Windows edition
echo -	help:				Provides help for ActivatePro commands
echo -	clear:				Clears the screen
echo:
goto expert





REM Exit Function
:exit
if %expertMode%==0 (
	cls
	echo Process Done. Exiting the program...
	timeout /t 3 > nul
	exit
) else (
	echo:
	goto expert
)




REM Error Functions
:errorMain
cls
set "mainSel=Invalid input, please try again: "
goto mainMenu


:errorWindows
cls
set "winSel=Invalid input, please try again: "
goto windows


:errorOffice
cls
set "offSel=Invalid input, please try again: "
goto office