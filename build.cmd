@echo off
setlocal
set CSC=%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\csc.exe
if not exist "%CSC%" set CSC=%WINDIR%\Microsoft.NET\Framework\v4.0.30319\csc.exe

set PIA=C:\Windows\assembly\GAC_MSIL\Microsoft.Office.Interop.OneNote\15.0.0.0__71e9bce111e9429c\Microsoft.Office.Interop.OneNote.dll
if not exist "%PIA%" (
  for /f "delims=" %%i in ('dir /b /s "C:\Windows\assembly\GAC_MSIL\Microsoft.Office.Interop.OneNote\Microsoft.Office.Interop.OneNote.dll" 2^>nul') do set PIA=%%i
)
if not exist "%PIA%" (
  echo ERROR: could not locate Microsoft.Office.Interop.OneNote.dll
  echo The OneNote interop assembly ships with desktop OneNote / Office.
  exit /b 1
)

echo Building onenote-md.exe ...
if not exist bin mkdir bin
"%CSC%" /nologo /optimize+ /platform:anycpu /r:"%PIA%" /out:bin\onenote-md.exe src\Program.cs src\OneNoteConverter.cs src\TextRunExtractor.cs
if errorlevel 1 (
  echo BUILD FAILED
  exit /b 1
)
echo Built bin\onenote-md.exe
endlocal
