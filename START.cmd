@echo off
title nextcloud-cu
color 1
echo -----------------------------------------------------------------
echo -----------------------------------------------------------------
echo Starte das Programm ...
echo WARNUNG! die Passwoerter der Nutzer duerfen nicht Simple sein,
echo sonst verweigert Nextcloud die Erstellung dieser.
pause
echo -----------------------------------------------------------------
powershell -NoProfile -ExecutionPolicy Bypass -Command "& '%~dp0main.ps1'"
echo nextcloud-cu Done, BYE!
echo -----------------------------------------------------------------
timeout /t 15
exit /b