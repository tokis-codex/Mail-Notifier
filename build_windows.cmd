@echo off
setlocal EnableExtensions DisableDelayedExpansion
cd /d "%~dp0"
title Mail Notifier - Windows PySide6 Build

echo Mail Notifier - Windows PySide6 Build
echo Working directory: %CD%
echo.

if not exist "mail_title_notifier.py" goto no_source
if not exist "requirements.txt" goto no_requirements
if not exist "build_exe.py" goto no_build_script
if not exist "find_python.ps1" goto no_finder

if exist ".build_python_path.txt" del /q ".build_python_path.txt" >nul 2>&1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%CD%\find_python.ps1"
if errorlevel 1 goto no_python
if not exist ".build_python_path.txt" goto no_python
set /p PYEXE=<".build_python_path.txt"
if not defined PYEXE goto no_python
if not exist "%PYEXE%" goto no_python

echo.
echo Python executable: "%PYEXE%"
"%PYEXE%" --version
if errorlevel 1 goto failed

set "VENV=%LOCALAPPDATA%\MailNotifierBuild\venv"
set "VPY=%VENV%\Scripts\python.exe"
echo.
echo [1/4] Preparing short-path virtual environment...
echo Build venv: "%VENV%"
if exist "%VPY%" (
    "%VPY%" -c "import sys;sys.exit(0 if (3,11)<=sys.version_info[:2]<(3,14) else 1)" >nul 2>&1
    if errorlevel 1 rmdir /s /q "%VENV%"
)
if not exist "%VPY%" (
    "%PYEXE%" -m venv "%VENV%"
    if errorlevel 1 goto failed
)
if not exist "%VPY%" goto failed

echo.
echo [2/4] Checking pip...
"%VPY%" -m pip --version
if errorlevel 1 (
    "%VPY%" -m ensurepip --upgrade
    if errorlevel 1 goto failed
)

echo.
echo [3/4] Installing minimal dependencies...
"%VPY%" -m pip install --no-cache-dir -r requirements.txt
if errorlevel 1 goto failed

echo.
echo [4/4] Building EXE...
"%VPY%" build_exe.py
if errorlevel 1 goto failed

echo.
echo BUILD SUCCEEDED.
echo Check the dist folder for the EXE.
pause
exit /b 0

:no_python
echo.
echo ERROR: No usable Python 3.11 - 3.13 was found for the current Windows user.
echo Run: py -0p
 echo Run: where python
 echo Run: python --version
 echo To force a specific Python path from CMD:
echo set "PYTHON_EXE=C:\path\to\python.exe"
echo call build_windows.cmd
goto failed
:no_source
echo ERROR: mail_title_notifier.py missing.
goto failed
:no_requirements
echo ERROR: requirements.txt missing.
goto failed
:no_build_script
echo ERROR: build_exe.py missing.
goto failed
:no_finder
echo ERROR: find_python.ps1 missing.
goto failed
:failed
echo.
echo BUILD FAILED - see the error above.
pause
exit /b 1
