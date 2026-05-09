@echo off
setlocal

cd /d "%~dp0"

echo ==========================================
echo UK Labour Market Dashboard Data Refresh
echo ==========================================
echo.
echo [1/2] Downloading latest source data...
python "download_uk_labour_market_data.py"
if errorlevel 1 goto :error

echo.
echo [2/2] Rebuilding model-ready tables...
python "prepare_model_data.py"
if errorlevel 1 goto :error

echo.
echo Refresh complete.
echo Next step: open Power BI Desktop and click Refresh.
echo.
pause
exit /b 0

:error
echo.
echo Refresh failed. Review the messages above.
echo.
pause
exit /b 1
