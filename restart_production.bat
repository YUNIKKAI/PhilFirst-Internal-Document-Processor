@echo off
echo 🔄 Restarting PhilFirst Renewal Notice Extractor
echo =================================================================

REM 1. Stop server if running
call stop_production.bat

REM 2. Start server again
call start_production.bat %1