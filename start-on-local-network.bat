@echo off
title Cornell Sales Analyzer (Local Network Mode)
echo ===================================================
echo   CORNELL SALES ANALYZER - LOCAL NETWORK MODE
echo ===================================================
echo.
echo Launching server and exposing it to your Local Network (Wi-Fi/LAN)...
echo.
echo ---------------------------------------------------
echo INSTRUCTIONS FOR SHARING WITH YOUR USERS:
echo.
echo 1. Make sure your computer and your users' devices
echo    are connected to the SAME Wi-Fi or network router.
echo.
echo 2. Once the server starts below, copy the "Network" 
echo    address (e.g., http://192.168.x.x:5173/) and send 
echo    it to your users.
echo ---------------------------------------------------
echo.
npm run dev -- --host
pause
