@echo off
REM Build ShellTAP.dll for w11-theming-suite
REM Generic XAML injection DLL -- can target any XAML-based process
setlocal

set "SRCDIR=C:\Dev\w11-theming-suite\native\ShellTAP"
set "OUTDIR=C:\Dev\w11-theming-suite\native\bin"
set "OBJDIR=C:\Dev\w11-theming-suite\native\ShellTAP\obj"
set "VCVARS=C:\Program Files (x86)\Microsoft Visual Studio\2022\BuildTools\VC\Auxiliary\Build\vcvars64.bat"

echo [BUILD] Initializing x64 environment...
call "%VCVARS%"
if errorlevel 1 goto :fail

if not exist "%OUTDIR%\" mkdir "%OUTDIR%"
if not exist "%OBJDIR%\" mkdir "%OBJDIR%"

echo [BUILD] Compiling ShellTAP.cpp...
cl.exe /nologo /LD /EHsc /O2 /MD /W3 /D_CRT_SECURE_NO_WARNINGS /I"%SRCDIR%" /DWIN32 /DNDEBUG /D_WINDOWS /D_USRDLL "%SRCDIR%\ShellTAP.cpp" /Fe:"%OUTDIR%\ShellTAP_new.dll" /Fo:"%OBJDIR%\\" /link /DEF:"%SRCDIR%\ShellTAP.def" /NOLOGO /DLL /MACHINE:X64 ole32.lib oleaut32.lib uuid.lib user32.lib shlwapi.lib WindowsApp.lib
if errorlevel 1 goto :fail

REM Try to replace existing DLL (may be locked if injected)
del /f "%OUTDIR%\ShellTAP.dll" 2>nul
if exist "%OUTDIR%\ShellTAP.dll" (
    echo [BUILD] WARNING: ShellTAP.dll is locked, built as ShellTAP_new.dll
    echo [BUILD] Kill the target process, then rename _new to replace
) else (
    move /y "%OUTDIR%\ShellTAP_new.dll" "%OUTDIR%\ShellTAP.dll" >nul
)

echo [BUILD] SUCCESS
dir "%OUTDIR%\ShellTAP*.dll"
goto :eof

:fail
echo [BUILD] FAILED
exit /b 1
