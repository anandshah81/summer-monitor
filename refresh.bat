@echo off
echo ============================================
echo   Summer Monitor - Refresh and Push to Web
echo ============================================
echo.

cd /d C:\Users\admin\claude\summer-monitor

echo Running summer_monitor.py...
python summer_monitor.py

echo.
echo Copying dashboard to repo...
python -c "import shutil; shutil.copy('summer_strength_monitor.html', 'index.html')"
if errorlevel 1 (
  echo ERROR: copy failed via python; falling back to powershell
  powershell -NoProfile -Command "Copy-Item -LiteralPath 'summer_strength_monitor.html' -Destination 'index.html' -Force"
)
REM Sanity check: index.html must match source size, otherwise abort before pushing
for %%A in (summer_strength_monitor.html) do set SRC_SIZE=%%~zA
for %%A in (index.html) do set DST_SIZE=%%~zA
if not "%SRC_SIZE%"=="%DST_SIZE%" (
  echo ERROR: index.html size %DST_SIZE% does not match source %SRC_SIZE%. Aborting push to avoid breaking live site.
  pause
  exit /b 1
)

echo Pushing to GitHub Pages...
git add index.html
git commit -m "Refresh dashboard %date% %time%"
git push

echo.
echo ============================================
echo   Done! Live at:
echo   https://anandshah81.github.io/summer-monitor/
echo ============================================
pause
