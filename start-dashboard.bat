@echo off
title Cornell Sales Analyzer Launcher
echo ===================================================
echo   CORNELL SALES ANALYZER - LAUNCHER
echo ===================================================
echo.
echo [1/2] Opening browser to http://localhost:5173...
start http://localhost:5173/
echo.
echo [2/2] Launching Vite development server...
echo.
npm run dev
pause
