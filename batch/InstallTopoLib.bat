@echo off
ECHO.
ECHO This batch file will install the TopoLib AddIn and some example spreadsheets
ECHO The TopoLib AddIn will be installed in your AddIns folder; in (%APPDATA%\Microsoft\AddIns)
ECHO If security restrictions make it difficult to run 3rd Party Addins from the AddIns folder, use the Templates folder instead:
ECHO The Templates folder can be found here:  (%APPDATA%\Microsoft\Templates)
ECHO The example spreadsheets will be installed in your Documents folder; in (%UserProfile%\Documents\TopoLib)
ECHO.
ECHO ** Please close MS Excel before continuing with the installation. **
ECHO.
PAUSE
:TESTAGAIN

wscript DetectExcelRunning.vbs
REM ECHO DetectExcelRunning script returned %errorlevel%

IF errorlevel 2 goto CONTINUE

ECHO. Excel is still running. Please close the application.
PAUSE

goto TESTAGAIN

:CONTINUE
ECHO We need to establish if you are using the 32-bit or 64-bit version of Excel.

wscript GetExcellBitness.vbs
REM ECHO GetExcellBitness script returned %errorlevel%

IF errorlevel 2 goto 64_BIT
:32_BIT
ECHO ** You are using 32-bit Excel **

REM define the name of the AddIn here
Set XLL=TopoLib-AddIn.xll
ECHO ** You will be using **%XLL%** as the 32-bit Excel AddIn **

PAUSE
ECHO.
ECHO Copying TopoLib files to AddIn-folder on AppData (%APPDATA%\Microsoft\AddIns\)

REM First delete the 64-bit AddIn in case it exists in the (32-bit) target folder
IF EXIST %APPDATA%\Microsoft\AddIns\TopoLib-AddIn64.xll DEL %APPDATA%\Microsoft\AddIns\TopoLib-AddIn64.xll /f /q
Copy ..\publish\x86\*.*  %APPDATA%\Microsoft\AddIns\

goto :CONTINUE2

:64_BIT
ECHO ** You are using 64-bit Excel **

REM define the name of the AddIn here
Set XLL=TopoLib-AddIn64.xll
ECHO ** You will be using **%XLL%** as as the 64-bit Excel AddIn **

PAUSE
ECHO.
ECHO Copying TopoLib files to AddIn-folder on AppData (%APPDATA%\Microsoft\AddIns\)

REM First delete the 32-bit AddIn in case it exists in the (64-bit) target folder
IF EXIST %APPDATA%\Microsoft\AddIns\TopoLib-AddIn.xll DEL %APPDATA%\Microsoft\AddIns\TopoLib-AddIn.xll /f /q
Copy ..\publish\x64\*.*  %APPDATA%\Microsoft\AddIns\

:CONTINUE2

ECHO Ready copying the TopoLib library and help file to your AddIn-folder: %APPDATA%\Microsoft\AddIns\

ECHO The next step is to install some example spreadsheets in your "Documents" area
ECHO These files will be stored under :("%UserProfile%\Documents\TopoLib)
ECHO The name of these files reflects the type of TopoLib functions being tested
PAUSE

ECHO.
ECHO Creating TopoLib folder in your "Documents" folder
IF NOT EXIST "%UserProfile%\Documents\TopoLib" MKDIR "%UserProfile%\Documents\TopoLib"

ECHO Copying example spreadsheets to (%UserProfile%\Documents\TopoLib)
Copy ..\publish\*.xlsb   "%UserProfile%\Documents\TopoLib"

ECHO.
ECHO Ready installing TopoLib...
ECHO.

ECHO To enable the TopoLib Add-In In Excel, please navigate to: "File->Options->Add-Ins".
ECHO Select Manage "Add-Ins" and press "Go..." to arrive at the Add-Ins dialog.
ECHO Select 'Browse' and move up by one folder in the path. Next descent into the AddIns folder
ECHO Double-click 'TopoLib-AddIn' to start the TopoLib Add-In.
ECHO Back in the Add-Ins dialog check that the 'TopoLib-AddIn' is now enabled. Close the dialog window.
ECHO.
ECHO Finally, you probably need to make the contents of the help-file "TopoLib.chm" accessible.
ECHO Otherwise you may see an empty help file (no information in the contents pane at the right). 
ECHO. 
ECHO Please use Explorer to navigate to (%APPDATA%\Microsoft\AddIns)
ECHO Right-click on "TopoLib.chm" and select "Properties"
ECHO On the General Page, at the bottom right, under "Security" select "Unblock" and press "OK".
ECHO This should make the contents of the help-file visible in Excel.
ECHO If you don't see security mentioned on the General Page, you do not have this issue.

ECHO.
ECHO To help you find "TopoLib-AddIn.chm", a File Explorer instance will now be launched in the AddIn folder...
ECHO Use this File Explorer to update the "TopoLib-AddIn.chm" properties by right-clicking this file.
ECHO.
PAUSE
CALL explorer %APPDATA%\Microsoft\AddIns\
ECHO.
ECHO Thank you for installing TopoLib.
ECHO.
EXIT
