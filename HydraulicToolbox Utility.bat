@echo off
SETLOCAL ENABLEDELAYEDEXPANSION

:: Define script paths and fixed file locations
set "SCRIPT_CONVERT_HYD=.\Scripts\01_HYT-Ditches.ps1"
set "SCRIPT_WORD_TO_CSV=.\Scripts\02_ReportConversion_WordToCSV.ps1"

:MENU
cls
echo ========================================
echo   HydraulicToolbox Utility
echo ========================================
echo.
echo Please select an option:
echo.
echo 1. Convert .HYD File 
echo    (Requires input.csv in current directory)
echo.
echo 2. Convert Word Report to CSV
echo    (Extract hydraulic analysis data)
echo.
echo 3. Exit
echo.
set /p choice="Enter your choice (1-3): "

if "%choice%"=="1" goto CONVERT_HYD
if "%choice%"=="2" goto WORD_TO_CSV
if "%choice%"=="3" goto END

echo Invalid option. Please try again.
pause
goto MENU

:CONVERT_HYD
cls
echo ========================================
echo   HYD File Conversion
echo ========================================
echo.
echo This option requires an 'input.csv' file 
echo in the current directory.
echo.

if not exist "input.csv" (
    echo Error: input.csv file not found.
    pause
    goto MENU
)

powershell -ExecutionPolicy Bypass -File "%SCRIPT_CONVERT_HYD%"
echo.
echo Conversion process completed. Check the output for details.
pause
goto MENU

:WORD_TO_CSV
cls
echo ========================================
echo   Word Report to CSV Conversion
echo ========================================
echo.
powershell -ExecutionPolicy Bypass -File "%SCRIPT_WORD_TO_CSV%"
echo.
echo Conversion process completed. Check the output for details.
pause
goto MENU

:END
echo Thank you for using HydraulicToolbox Utility.
exit /b