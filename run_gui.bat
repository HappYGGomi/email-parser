@echo off
chcp 65001 >nul
echo ================================================
echo      Email Parser GUI Starting...
echo ================================================
echo.

python gui.py

if errorlevel 1 (
    echo.
    echo Error occurred!
    echo.
    echo Please make sure to install required libraries:
    echo   pip install openpyxl
    echo.
    pause
)
