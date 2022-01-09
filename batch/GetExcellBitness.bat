@echo off

wscript GetExcellBitness.vbs

echo wscript returned %errorlevel%

IF errorlevel 2 goto 64_bit

echo bitness: 32
goto :endpoint

:64_bit
echo bitness: 64

:endpoint

pause
