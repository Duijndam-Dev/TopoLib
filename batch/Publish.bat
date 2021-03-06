@echo.
@echo.
@echo Solution folder:  %1
@echo Configuration  :  %2

if not exist "%1publish\x64" (
    mkdir "%1publish\x64" 2>nul
    if not errorlevel 1 (

		@echo Created: "%1publish\x64"
        copy "%1packages\SharpProj.9.1.183\contentFiles\any\any\proj.db" "%1publish\x64"
        copy "%1packages\SharpProj.9.1.183\contentFiles\any\any\proj.ini" "%1publish\x64"
        copy "%1packages\SharpProj.Core.9.1.183\runtimes\win-x64\lib\net462\SharpProj.dll" "%1publish\x64"
		@echo Copied database and 64-bit dll
    )
)

if not exist "%1publish\x86" (
    mkdir "%1publish\x86" 2>nul
    if not errorlevel 1 (

		@echo Created: "%1publish\x86"
        copy "%1packages\SharpProj.9.1.183\contentFiles\any\any\proj.db" "%1publish\x86"
        copy "%1packages\SharpProj.9.1.183\contentFiles\any\any\proj.ini" "%1publish\x86"
        copy "%1packages\SharpProj.Core.9.1.183\runtimes\win-x86\lib\net462\SharpProj.dll" "%1publish\x86"
		@echo Copied database and 32-bit dll
    )
)

IF "%~2" == "Release" GOTO Release
@echo Not in 'release' mode: no xxl/chm files to be published
exit

:Release

rem copy 64-bit stuff first
COPY "%1TopoLib\bin\Release\TopoLib-AddIn64-packed.xll"   "%1publish\x64\TopoLib-AddIn64.xll"
COPY "%1TopoLib\bin\Release\TopoLib-AddIn.chm"               "%1publish\x64\TopoLib-AddIn.chm"
@echo Copied 64-bit xll and chm file

rem copy 32-bit stuff next
COPY "%1TopoLib\bin\Release\TopoLib-AddIn-packed.xll"      "%1publish\x86\TopoLib-AddIn.xll"
COPY "%1TopoLib\bin\Release\TopoLib-AddIn.chm"               "%1publish\x86\TopoLib-AddIn.chm"
@echo Copied 32-bit xll and chm file

rem now copy the example spreadsheets to the publish folder
COPY /Y "%1TopoLib\Samples\*.xlsb"                                  "%1publish\"
@echo Copied example spreadsheets to publish folder

rem copy chm file back to debug folder (need it here too)
COPY "%1TopoLib\bin\Release\TopoLib-AddIn.chm"               "%1TopoLib\bin\Debug\TopoLib-AddIn.chm"
@echo Copied chm file back to debug folder (need it here too)

rem sign the 32/64 bit XLL and DLL files in the publish folder
call D:\Source\VS19\CodeSigning\SignTopoLib.bat %1 %2

rem use 7zip to create an archive containing 'publish' and 'batch' folders on H:/drive
REM just to make sure; first change current directory to project folder
CD /D %1
echo %cd%
"C:\Program Files\7-Zip\7z.exe" -y a "H:\Source\VS19\TopoLib.zip" batch\ publish\ readme.txt
@echo.

rem copy the publish folder to the home-drive to work with different PC's.
@echo Use ROBOCOPY to copy publish folder from "%1publish" to "H:\Source\VS19\TopoLib\publish"
ROBOCOPY "%1publish" "H:\Source\VS19\TopoLib\publish" /S
@echo Copied publish folder to home-drive for further distribution
@echo.

rem next line forces ERRORLEVEL = 0 upon exit
exit /b 0 