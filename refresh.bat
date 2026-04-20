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
copy /Y summer_strength_monitor.html index.html >nul

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
