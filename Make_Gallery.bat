@echo off
setlocal enabledelayedexpansion
title 2026 AC Gallery

echo ============================================
echo    2026 AC Gallery
echo ============================================
echo  Builds "2026 AC Gallery.html" in the screenshot folder from the pictures
echo  already there, then opens it in your browser. Run it again whenever
echo  new screenshots, ratings or notes have been added.
echo.

REM --- locate python ---
set "PYRUN="
where py.exe >nul 2>&1 && set "PYRUN=py.exe -3"
if not defined PYRUN where python.exe >nul 2>&1 && set "PYRUN=python.exe"
if not defined PYRUN if exist "%USERPROFILE%\anaconda3\python.exe" set "PYRUN="%USERPROFILE%\anaconda3\python.exe""
if not defined PYRUN if exist "%USERPROFILE%\miniconda3\python.exe" set "PYRUN="%USERPROFILE%\miniconda3\python.exe""
if not defined PYRUN (
    echo ERROR: Python not found. Run First_Time_Setup.bat first.
    pause
    exit /b 1
)

%PYRUN% "%~dp0make_gallery.py" --open
set "RC=!errorlevel!"

echo.
if not "!RC!"=="0" pause
exit /b !RC!
