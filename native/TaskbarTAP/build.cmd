@echo off
REM Build TaskbarTAP.dll for w11-theming-suite
setlocal

set "SRCDIR=C:\Dev\w11-theming-suite\native\TaskbarTAP"
set "OUTDIR=C:\Dev\w11-theming-suite\native\bin"
set "OBJDIR=C:\Dev\w11-theming-suite\native\TaskbarTAP\obj"
set "VCVARS=C:\Program Files (x86)\Microsoft Visual Studio\2022\BuildTools\VC\Auxiliary\Build\vcvars64.bat"

echo [BUILD] Initializing x64 environment...
call "%VCVARS%"
if errorlevel 1 goto :fail

if not exist "%OUTDIR%\" mkdir "%OUTDIR%"
if not exist "%OBJDIR%\" mkdir "%OBJDIR%"

echo [BUILD] Compiling TaskbarTAP.cpp...
cl.exe /nologo /LD /EHsc /O2 /MD /W3 /D_CRT_SECURE_NO_WARNINGS /I"%SRCDIR%" /DWIN32 /DNDEBUG /D_WINDOWS /D_USRDLL "%SRCDIR%\TaskbarTAP.cpp" /Fe:"%OUTDIR%\TaskbarTAP_new.dll" /Fo:"%OBJDIR%\\" /link /DEF:"%SRCDIR%\TaskbarTAP.def" /NOLOGO /DLL /MACHINE:X64 ole32.lib oleaut32.lib uuid.lib user32.lib shlwapi.lib WindowsApp.lib
if errorlevel 1 goto :fail

REM Try to replace existing DLL (may be locked if injected)
del /f "%OUTDIR%\TaskbarTAP.dll" 2>nul
if exist "%OUTDIR%\TaskbarTAP.dll" (
    echo [BUILD] WARNING: TaskbarTAP.dll is locked, built as TaskbarTAP_new.dll
    echo [BUILD] Restart explorer.exe then rename _new to replace
) else (
    move /y "%OUTDIR%\TaskbarTAP_new.dll" "%OUTDIR%\TaskbarTAP.dll" >nul
)

echo [BUILD] SUCCESS
dir "%OUTDIR%\TaskbarTAP*.dll"
goto :eof

:fail
echo [BUILD] FAILED
exit /b 1
