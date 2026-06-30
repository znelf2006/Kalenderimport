@echo off
powershell -ExecutionPolicy Bypass -File "%~dp0dienstplan-setup.ps1"
if %errorlevel% neq 0 (
    echo.
    echo Fehler aufgetreten. Bitte obige Meldung fotografieren und weiterschicken.
    pause
)
