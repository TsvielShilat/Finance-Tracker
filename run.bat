@echo off
chcp 65001 >nul
cd /d "%~dp0"
set PY=python
where py >nul 2>nul && set PY=py
%PY% --version >nul 2>nul || (echo Python is not installed. Get it from python.org and tick "Add python.exe to PATH". & pause & exit /b)
%PY% -c "import pandas, openpyxl, lxml" >nul 2>nul || %PY% -m pip install -r requirements.txt

:menu
echo.
echo ===== Family finance =====
echo  1  Import new export files
echo  2  Apply my answers from review.xlsx
echo  3  Report - all years
echo  4  Report - one year
echo  5  Status
echo  0  Exit
set c=
set /p c=Choose: 
if "%c%"=="1" %PY% finance.py import
if "%c%"=="2" %PY% finance.py apply
if "%c%"=="3" %PY% finance.py report
if "%c%"=="4" goto year
if "%c%"=="5" %PY% finance.py status
if "%c%"=="0" exit /b
goto menu

:year
set y=
set /p y=Year (e.g. 2025): 
%PY% finance.py report --year %y%
goto menu
