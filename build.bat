@echo off
echo ===================================
echo   Building Homebox Export Tool
echo ===================================
echo.

pyinstaller --onefile --windowed --name "HomeboxExport" --collect-data fpdf2 homebox_export_gui.py

echo.
if exist "dist\HomeboxExport.exe" (
    echo BUILD SUCCESSFUL!
    echo.
    echo Your .exe is at: dist\HomeboxExport.exe
) else (
    echo BUILD FAILED - check output above for errors.
)
echo.
pause
