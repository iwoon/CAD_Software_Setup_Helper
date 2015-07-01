@echo off
cls
setlocal
:: Check WMIC is available
WMIC.EXE Alias /? >NUL 2>&1 || GOTO s_done
:: Use WMIC to retrieve date and time
FOR /F "skip=1 tokens=1-6" %%G IN ('WMIC Path Win32_LocalTime Get Day^,Hour^,Minute^,Month^,Second^,Year /Format:table') DO (
   IF "%%~L"=="" goto s_done
      Set _yyyy=%%L
      Set _mm=00%%J
      Set _dd=00%%G
      Set _hour=00%%H
      SET _minute=00%%I
	  SET bkfile=%%L%%J%%G%%H%%I%%K
)
IF DEFINED bkfile ECHO %bkfile%
:s_done
pushd %~dp0
cd %cd%
del /f/q Js_Cr.7z fixjscr.bat
findstr /v /r /c:"REM" /c:"^$" /c:"^[\ \ ]*$" FixJsCr.source>fixjscr.bat
..\7-ZipPortable\App\7-Zip\7z.exe a Js_Cr.7z ..\dll\ ..\common\ ..\windows\ fixjscr.bat
title waiting for building installation file.
copy /b ..\7zS.sfx+config.txt+Js_Cr.7z FixJsCr-%bkfile%.exe
del /f/q Js_Cr.7z fixjscr.bat
endlocal
exit