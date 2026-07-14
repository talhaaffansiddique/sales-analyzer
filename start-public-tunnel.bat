@echo off
title Cornell Sales Analyzer (Public Internet Tunnel)
echo ===================================================
echo   CORNELL SALES ANALYZER - PUBLIC INTERNET TUNNEL
echo ===================================================
echo.
echo This script exposes your running local server to the public internet
echo so anyone in the world can access it.
echo.
echo IMPORTANT: Make sure you already started your local server first
echo (by running start-dashboard.bat or start-on-local-network.bat).
echo.
echo Launching tunnel on port 5173...
echo.
npx localtunnel --port 5173
pause
