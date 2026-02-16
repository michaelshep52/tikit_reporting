@echo off
cls
echo ============================================================
echo              TICKET REPORTER MENU
echo ============================================================
echo.
echo Select which report to generate:
echo.
echo   1. Daily Report (IT_Daily_Report.xlsx)
echo   2. Weekly Report (Weekly_Analysis_Report.pdf)
echo   3. Both Reports
echo   4. Exit
echo.
echo ============================================================
echo.

set /p choice="Enter your choice (1-4): "

if "%choice%"=="1" goto daily
if "%choice%"=="2" goto weekly
if "%choice%"=="3" goto both
if "%choice%"=="4" goto exit
echo Invalid choice. Please try again.
pause
goto start

:daily
echo.
echo Running Daily Report...
echo.
python daily_report.py
if exist IT_Daily_Report.xlsx start IT_Daily_Report.xlsx
goto end

:weekly
echo.
echo Running Weekly Report...
echo.
python weekly_report.py
if exist Weekly_Analysis_Report.pdf start Weekly_Analysis_Report.pdf
goto end

:both
echo.
echo Running Both Reports...
echo.
python daily_report.py
python weekly_report.py
if exist IT_Daily_Report.xlsx start IT_Daily_Report.xlsx
if exist Weekly_Analysis_Report.pdf start Weekly_Analysis_Report.pdf
goto end

:exit
echo.
echo Goodbye!
exit

:end
echo.
echo Done!
pause
