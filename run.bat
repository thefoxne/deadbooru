@echo off
chcp 65001 >nul
title deadbooru - Offline Danbooru Tagging & Upload Hub
echo ========================================================
echo   deadbooru - Offline Danbooru Tagging & Upload Hub
echo   Zero-disk mode: In-memory processing
echo ========================================================
echo.
python app.py
if errorlevel 1 (
    echo.
    echo [!] An error occurred while starting the application.
    pause
)
