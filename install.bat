@echo off
chcp 65001 >nul
echo ================================================
echo      Email Parser - Installing Dependencies
echo ================================================
echo.
echo Installing required libraries...
echo   - openpyxl (for Excel export)
echo   - extract-msg (for .msg file support)
echo.

pip install openpyxl extract-msg

if errorlevel 1 (
    echo.
    echo Failed with 'pip'. Trying 'python -m pip'...
    python -m pip install openpyxl extract-msg
)

if errorlevel 1 (
    echo.
    echo Failed with 'python -m pip'. Trying 'py -m pip'...
    py -m pip install openpyxl extract-msg
)

echo.
echo ================================================
echo Installation completed!
echo ================================================
echo.
echo You can now run the program:
echo   - Double-click run_gui.bat for GUI
echo   - Or run: python gui.py
echo.
echo Supported file formats:
echo   - .eml (Email Message Format)
echo   - .msg (Microsoft Outlook Message)
echo.
pause
